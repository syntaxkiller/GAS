/**
 * ServerSync.gs
 * Centralized server synchronization operations
 * Handles all data uploads to the remote server
 */

// === CORE SERVER COMMUNICATION ===

/**
 * Send data to the remote server
 * @param {Object} payload - Data to send
 * @throws {Error} on server error
 */
function sendToServer(payload) {
  const response = UrlFetchApp.fetch(SERVER_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + SECRET_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Server error (${response.getResponseCode()}): ${response.getContentText()}`);
  }
  
  return JSON.parse(response.getContentText() || '{}');
}

/**
 * Trigger the combine operation on the server
 */
function triggerCombine() {
  sendToServer({ dataType: 'combine' });
}

// === DATA UPLOAD FUNCTIONS ===

/**
 * Upload persistent storage (links, hidden, archived)
 * @returns {{success: boolean}}
 */
function pushPersistentStorage() {
  const props = PropertiesService.getDocumentProperties();
  
  const payload = {
    dataType: 'persistentStorage',
    payload: {
      questionLinks: JSON.parse(props.getProperty(SCRIPT_PROPERTY_KEY) || '{}'),
      hiddenQuestions: JSON.parse(props.getProperty(HIDDEN_QUESTIONS_KEY) || '[]'),
      archivedQuestions: JSON.parse(props.getProperty(ARCHIVED_QUESTIONS_KEY) || '[]')
    }
  };

  sendToServer(payload);
  return { success: true };
}

/**
 * Upload question configuration only
 * @returns {{success: boolean, count: number}}
 */
function pushQuestionConfigOnly() {
  const mgr = new ConfigManager();
  const uploadData = mgr.getUploadData();
  
  sendToServer({
    dataType: 'questionConfig',
    payload: uploadData.data
  });

  // Trigger combine after config upload
  triggerCombine();

  return { success: true, count: uploadData.count };
}

/**
 * Upload survey schema (question IDs and texts)
 * @param {string} sheetName
 * @param {Array} questionIds
 * @param {Array} questionTexts
 */
function uploadSchema(sheetName, questionIds, questionTexts) {
  const props = PropertiesService.getDocumentProperties();
  const archived = JSON.parse(props.getProperty(ARCHIVED_QUESTIONS_KEY) || '[]');
  
  const schema = questionIds.reduce((acc, id, i) => {
    const qId = String(id).trim();
    if (qId && !archived.includes(qId)) {
      acc[qId] = questionTexts[i] || '';
    }
    return acc;
  }, {});

  sendToServer({ dataType: 'surveySchema', sheetName, payload: schema });
}

/**
 * Upload survey data as CSV
 * @param {string} sheetName
 * @param {Array} questionIds
 * @param {Array} rows - Data rows (excluding headers)
 */
function uploadCSV(sheetName, questionIds, rows) {
  const props = PropertiesService.getDocumentProperties();
  const archived = JSON.parse(props.getProperty(ARCHIVED_QUESTIONS_KEY) || '[]');
  
  // Filter to active (non-archived) questions
  const activeIds = [];
  const activeIndices = [];

  questionIds.forEach((id, i) => {
    const qId = String(id).trim();
    if (qId && !archived.includes(qId)) {
      activeIds.push(qId);
      activeIndices.push(i);
    }
  });

  // Build CSV
  let csv = ['respondent_id', ...activeIds].join(',') + '\n';
  let count = 0;

  rows.forEach(row => {
    // Skip empty rows
    if (!row.some(c => c !== null && c !== undefined && String(c).trim() !== '')) return;

    const values = [sheetName + '_r' + (count + 1)];
    activeIndices.forEach(i => values.push(escapeCsvValue(row[i])));
    csv += values.join(',') + '\n';
    count++;
  });

  sendToServer({ dataType: 'surveyData', sheetName, payload: csv });
}

/**
 * Escape a value for CSV format
 * @param {any} cell
 * @returns {string}
 */
function escapeCsvValue(cell) {
  if (cell === null || cell === undefined) return '';
  let val = String(cell);
  if (val.includes(',') || val.includes('"') || val.includes('\n')) {
    val = '"' + val.replace(/"/g, '""') + '"';
  }
  return val;
}

/**
 * Upload all survey data from all sheets
 * @returns {{success: boolean, sheetsProcessed: number}}
 */
function pushAllSurveyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let count = 0;

  ss.getSheets()
    .filter(isDataSheet)
    .forEach(sheet => {
      const lastCol = sheet.getLastColumn();
      const lastRow = sheet.getLastRow();
      if (lastCol === 0 || lastRow < 3) return;

      const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      const sheetName = sheet.getName();
      
      // Upload schema (row 0 = texts, row 1 = ids)
      uploadSchema(sheetName, data[1], data[0]);
      
      // Upload data (rows 2+)
      uploadCSV(sheetName, data[1], data.slice(2));
      count++;
    });

  return { success: true, sheetsProcessed: count };
}

/**
 * Push all data to server (full sync)
 * @returns {{success: boolean, storage: Object, surveys: Object, config: Object}}
 */
function pushAllDataToServer() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Confirm with user
    if (ui.alert('Confirm Upload', 'Upload all data to server?', ui.ButtonSet.YES_NO) !== ui.Button.YES) {
      return { cancelled: true };
    }

    ui.alert('Upload Started', 'Uploading...', ui.ButtonSet.OK);

    // Upload all components
    const storage = pushPersistentStorage();
    const surveys = pushAllSurveyData();
    
    let config = null;
    try {
      config = pushQuestionConfigOnly();
    } catch (e) {
      Logger.log('No config: ' + e.message);
      config = { success: false, message: 'No configuration' };
    }

    // Trigger final combine
    triggerCombine();

    // Show success
    ui.alert(
      'Success',
      `Upload complete!\n\nStorage: ✓\nSurveys: ${surveys.sheetsProcessed || 0} years\nConfig: ${config.success ? config.count + ' questions' : 'None'}`,
      ui.ButtonSet.OK
    );

    return { success: true, storage, surveys, config };
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error: ' + err.message);
    throw err;
  }
}

/**
 * Save config and upload in one operation
 * @param {Object} configObj
 * @returns {{success: boolean, saved: number, uploaded: number}}
 */
function saveAndUploadConfig(configObj) {
  const saveResult = saveConfigToSheetBulk(configObj);
  const uploadResult = pushQuestionConfigOnly();
  return { success: true, saved: saveResult.rowCount, uploaded: uploadResult.count };
}

// === DEBUG FUNCTIONS ===

/**
 * View all stored data in a debug sheet
 */
function viewStoredData() {
  const props = PropertiesService.getDocumentProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let debug = ss.getSheetByName('Stored_Data_Debug');
  if (!debug) {
    debug = ss.insertSheet('Stored_Data_Debug');
  } else {
    debug.clear();
  }

  const dataItems = [
    ['Links JSON:', props.getProperty(SCRIPT_PROPERTY_KEY) || '{}'],
    ['Hidden JSON:', props.getProperty(HIDDEN_QUESTIONS_KEY) || '[]'],
    ['Archived JSON:', props.getProperty(ARCHIVED_QUESTIONS_KEY) || '[]']
  ];

  dataItems.forEach(([label, data], i) => {
    debug.getRange(i * 2 + 1, 1).setValue(label);
    debug.getRange(i * 2 + 2, 1).setValue(data);
  });

  SpreadsheetApp.getUi().alert('Stored data written to "Stored_Data_Debug" sheet.');
}

/**
 * Clear all stored data
 */
function clearAllStoredData() {
  const props = PropertiesService.getDocumentProperties();
  
  [SCRIPT_PROPERTY_KEY, HIDDEN_QUESTIONS_KEY, ARCHIVED_QUESTIONS_KEY, IDS_VERIFIED_FLAG]
    .forEach(key => props.deleteProperty(key));
  
  SpreadsheetApp.getUi().alert('All stored data cleared.');
}

/**
 * Preview upload data for current sheet
 */
function previewUploadData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const ui = SpreadsheetApp.getUi();

  if (data.length < 3) {
    ui.alert('Preview failed', 'Need at least 3 rows', ui.ButtonSet.OK);
    return;
  }

  const props = PropertiesService.getDocumentProperties();
  const archived = JSON.parse(props.getProperty(ARCHIVED_QUESTIONS_KEY) || '[]');
  const ids = data[1];
  const activeIds = ids.filter((id, i) => id && String(id).trim() && !archived.includes(id));

  Logger.log(`=== PREVIEW ===\n${activeIds.length} active questions (${ids.filter(Boolean).length - activeIds.length} archived)`);
  ui.alert(
    'Preview',
    `${activeIds.length} active questions\n${ids.filter(Boolean).length - activeIds.length} archived`,
    ui.ButtonSet.OK
  );
}

/**
 * Get persistent storage data for debugging
 * @returns {{links: Object, hidden: Array, archived: Array}}
 */
function getPersistentStorageData() {
  const props = PropertiesService.getDocumentProperties();
  
  try {
    return {
      links: JSON.parse(props.getProperty(SCRIPT_PROPERTY_KEY) || '{}'),
      hidden: JSON.parse(props.getProperty(HIDDEN_QUESTIONS_KEY) || '[]'),
      archived: JSON.parse(props.getProperty(ARCHIVED_QUESTIONS_KEY) || '[]')
    };
  } catch (e) {
    throw new Error('Failed to parse storage JSON.');
  }
}
