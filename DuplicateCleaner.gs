/**
 * Opens the Floating UI to manage duplicates.
 */
function showDuplicateManager() {
  const html = HtmlService.createTemplateFromFile('DuplicateCleanerUI')
    .evaluate()
    .setTitle('Duplicate Question Manager')
    .setWidth(800)
    .setHeight(600);
  
  // Modeless allows you to interact with the sheet while the window is open
  SpreadsheetApp.getUi().showModelessDialog(html, 'Duplicate Question Configuration Manager');
}

/**
 * Scans the Config sheet and returns IDs that appear more than once,
 * along with their full row data.
 */
function getConfigDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME); // Uses constant from Core.gs
  
  if (!sheet) {
    throw new Error(`Sheet "${CONFIG_SHEET_NAME}" not found.`);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2) return { headers: [], duplicates: {} };

  // Get headers (Row 1) and all data (Row 2 to End)
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Map to track occurrences
  const idMap = {};

  data.forEach((row, index) => {
    const id = String(row[0]).trim(); // Assuming ID is in Column A (index 0)
    const realRowNumber = index + 2; // +2 because data starts at row 2

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

  // Filter out IDs that only have 1 entry (no duplicates)
  const duplicates = {};
  Object.entries(idMap).forEach(([id, occurrences]) => {
    if (occurrences.length > 1) {
      duplicates[id] = occurrences;
    }
  });

  return {
    headers: headers,
    duplicates: duplicates
  };
}

/**
 * Deletes a specific row in the Config sheet.
 * @param {number} rowNumber - The 1-based row index to delete.
 */
function deleteConfigRow(rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (!sheet) throw new Error("Sheet not found");
  
  // Delete the row
  sheet.deleteRow(rowNumber);
  
  return { success: true };
}

// Add menu item to your existing onOpen or create a new one
function onOpenDuplicateManager() {
  SpreadsheetApp.getUi().createMenu('⚡ Cleanup Tools')
    .addItem('Manage Config Duplicates', 'showDuplicateManager')
    .addToUi();
}