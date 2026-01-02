/**
 * DuplicateCleaner.gs
 * Duplicate detection and cleanup utilities for the configuration sheet
 */

/**
 * Find duplicate question IDs in a sheet
 * @param {string} sheetName - Name of the sheet to scan
 * @returns {{headers: string[], duplicates: Object}} - Headers and map of duplicate IDs to their occurrences
 */
function findDuplicatesInSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2 || lastCol === 0) {
    return { headers: [], duplicates: {} };
  }
  
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  // Track occurrences of each ID
  const idMap = {};
  
  data.forEach((row, index) => {
    const id = String(row[0]).trim();
    const realRowNumber = index + 2; // Account for 1-based indexing and header row
    
    if (id) {
      if (!idMap[id]) {
        idMap[id] = [];
      }
      idMap[id].push({
        rowNumber: realRowNumber,
        values: row
      });
    }
  });
  
  // Filter to only duplicates (more than one occurrence)
  const duplicates = {};
  Object.entries(idMap).forEach(([id, occurrences]) => {
    if (occurrences.length > 1) {
      duplicates[id] = occurrences;
    }
  });
  
  return { headers, duplicates };
}

/**
 * Delete a row from a sheet by row number
 * @param {string} sheetName - Name of the sheet
 * @param {number} rowNumber - 1-based row number to delete
 * @returns {{success: boolean}}
 */
function deleteRowFromSheet(sheetName, rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }
  
  if (rowNumber < 2) {
    throw new Error('Cannot delete header row');
  }
  
  sheet.deleteRow(rowNumber);
  return { success: true };
}

/**
 * Delete multiple rows from a sheet (in descending order to preserve row numbers)
 * @param {string} sheetName - Name of the sheet
 * @param {number[]} rowNumbers - Array of 1-based row numbers to delete
 * @returns {{success: boolean, deleted: number}}
 */
function deleteRowsFromSheet(sheetName, rowNumbers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }
  
  // Sort in descending order so we delete from bottom up
  const sortedRows = [...rowNumbers].sort((a, b) => b - a);
  
  let deleted = 0;
  sortedRows.forEach(rowNum => {
    if (rowNum >= 2) {
      sheet.deleteRow(rowNum);
      deleted++;
    }
  });
  
  return { success: true, deleted };
}

/**
 * Keep only the first occurrence of each duplicate ID
 * @param {string} sheetName - Name of the sheet
 * @returns {{success: boolean, deleted: number}}
 */
function deduplicateSheet(sheetName) {
  const { duplicates } = findDuplicatesInSheet(sheetName);
  
  if (Object.keys(duplicates).length === 0) {
    return { success: true, deleted: 0 };
  }
  
  const rowsToDelete = [];
  
  Object.values(duplicates).forEach(occurrences => {
    // Sort by row number and keep the first one
    const sorted = occurrences.sort((a, b) => a.rowNumber - b.rowNumber);
    // Add all but the first to the delete list
    sorted.slice(1).forEach(o => rowsToDelete.push(o.rowNumber));
  });
  
  return deleteRowsFromSheet(sheetName, rowsToDelete);
}

/**
 * Get duplicate data specifically for the Question Configuration sheet
 * This is the main entry point used by the DuplicateCleanerUI
 * @returns {{headers: string[], duplicates: Object}}
 */
function getConfigSheetDuplicates() {
  return findDuplicatesInSheet(CONFIG_SHEET_NAME);
}

/**
 * Delete a row from the Question Configuration sheet
 * @param {number} rowNumber - 1-based row number to delete
 * @returns {{success: boolean}}
 */
function deleteConfigSheetRow(rowNumber) {
  return deleteRowFromSheet(CONFIG_SHEET_NAME, rowNumber);
}

/**
 * Deduplicate the Question Configuration sheet
 * @returns {{success: boolean, deleted: number}}
 */
function deduplicateConfigSheet() {
  return deduplicateSheet(CONFIG_SHEET_NAME);
}

/**
 * Validate that all question IDs in the config sheet exist in the data sheets
 * @returns {{valid: string[], invalid: string[], orphaned: string[]}}
 */
function validateConfigQuestionIds() {
  const mgr = getDataManager();
  const allDataQuestions = new Set(mgr.getAllQuestions().map(q => q.questionId));
  
  const configMgr = new ConfigManager();
  const configIds = Object.keys(configMgr.loadAll());
  
  const valid = [];
  const invalid = [];
  
  configIds.forEach(id => {
    if (allDataQuestions.has(id)) {
      valid.push(id);
    } else {
      invalid.push(id);
    }
  });
  
  // Find questions in data that aren't in config (orphaned)
  const configIdSet = new Set(configIds);
  const orphaned = [];
  
  allDataQuestions.forEach(id => {
    if (!configIdSet.has(id)) {
      orphaned.push(id);
    }
  });
  
  return { valid, invalid, orphaned };
}

/**
 * Remove invalid question IDs from the config sheet
 * @returns {{success: boolean, removed: number}}
 */
function removeInvalidConfigEntries() {
  const { invalid } = validateConfigQuestionIds();
  
  if (invalid.length === 0) {
    return { success: true, removed: 0 };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (!sheet) {
    return { success: false, removed: 0 };
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { success: true, removed: 0 };
  }
  
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const invalidSet = new Set(invalid);
  
  // Find row numbers to delete (work backwards)
  const rowsToDelete = [];
  ids.forEach((id, i) => {
    if (invalidSet.has(String(id).trim())) {
      rowsToDelete.push(i + 2); // Convert to 1-based row number
    }
  });
  
  // Delete from bottom up
  rowsToDelete.sort((a, b) => b - a).forEach(row => {
    sheet.deleteRow(row);
  });
  
  return { success: true, removed: rowsToDelete.length };
}
