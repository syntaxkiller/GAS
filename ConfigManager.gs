/**
 * ConfigManager.gs
 * Centralized manager for question configuration operations
 * Consolidates all config save/load functionality
 */

class ConfigManager {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = null;
    this.existingIds = null;
  }

  /**
   * Get or create the config sheet
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   */
  getSheet() {
    if (!this.sheet) {
      this.sheet = this.ss.getSheetByName(CONFIG_SHEET_NAME);
      
      if (!this.sheet) {
        this.sheet = this.ss.insertSheet(CONFIG_SHEET_NAME);
        this._setupSheet();
      }
    }
    return this.sheet;
  }

  /**
   * Set up a new config sheet with headers and formatting
   * @private
   */
  _setupSheet() {
    const sheet = this.sheet;
    
    // Set headers
    sheet.getRange(1, 1, 1, CONFIG_HEADERS.length)
      .setValues([CONFIG_HEADERS])
      .setBackground('#4A86E8')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setWrap(true);
    
    sheet.setFrozenRows(1);
    
    // Set column widths
    CONFIG_COLUMN_WIDTHS.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
    
    // Set up validation
    this._setupValidation();
  }

  /**
   * Set up data validation rules for the config sheet
   * @private
   */
  _setupValidation() {
    const sheet = this.sheet;
    const rules = [
      [3, SpreadsheetApp.newDataValidation().requireValueInList(getQuestionTypes()).setAllowInvalid(false).build()],
      [6, SpreadsheetApp.newDataValidation().requireCheckbox().build()],
      [7, SpreadsheetApp.newDataValidation().requireValueInList(getSortTypes()).setAllowInvalid(false).build()],
      [8, SpreadsheetApp.newDataValidation().requireValueInList(getSortTypes()).setAllowInvalid(false).build()],
      [11, SpreadsheetApp.newDataValidation().requireValueInList(['', ...COLOR_GROUP_EMOJIS]).setAllowInvalid(true).build()]
    ];
    
    rules.forEach(([col, rule]) => {
      sheet.getRange(2, col, MAX_VALIDATION_ROWS, 1).setDataValidation(rule);
    });
  }

  /**
   * Load existing question IDs and their row numbers
   * @returns {Object} - { questionId: rowNumber }
   */
  getExistingIds() {
    if (this.existingIds) return this.existingIds;
    
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();
    
    this.existingIds = {};
    
    if (lastRow > 1) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      ids.forEach(([id], i) => {
        if (id) {
          this.existingIds[String(id).trim()] = i + 2;
        }
      });
    }
    
    return this.existingIds;
  }

  /**
   * Invalidate the cached existing IDs
   */
  clearIdCache() {
    this.existingIds = null;
  }

  /**
   * Build a row array from a config object
   * @param {string} questionId
   * @param {Object} config
   * @param {string} questionText
   * @returns {Array}
   */
  buildRow(questionId, config, questionText = '') {
    // Handle customOrder - convert array to string if needed
    let customOrder = config.customOrder || '';
    if (Array.isArray(customOrder)) {
      customOrder = customOrder.join(',');
    }
    
    let customOrderMobile = config.customOrderMobile || '';
    if (Array.isArray(customOrderMobile)) {
      customOrderMobile = customOrderMobile.join(',');
    }
    
    return [
      questionId,
      questionText,
      config.questionType || '',
      config.chartType || '',
      config.mobileChartType || '',
      config.hasOther || false,
      config.sortType || 'none',
      config.sortTypeMobile || '',
      customOrder,
      customOrderMobile,
      config.colorGroup || ''
    ];
  }

  /**
   * Get question text for a question ID
   * @param {string} questionId
   * @returns {string}
   */
  getQuestionText(questionId) {
    try {
      const mgr = getDataManager();
      const result = mgr.findQuestionColumn(questionId);
      return result.questionText || '';
    } catch (e) {
      Logger.log(`Could not get question text for ${questionId}: ${e.message}`);
      return '';
    }
  }

  /**
   * Save a single question configuration
   * @param {string} questionId
   * @param {Object} config
   * @returns {{success: boolean}}
   */
  save(questionId, config) {
    const sheet = this.getSheet();
    const existingIds = this.getExistingIds();
    const questionText = this.getQuestionText(questionId);
    const row = this.buildRow(questionId, config, questionText);
    
    if (existingIds[questionId]) {
      // Update existing row
      sheet.getRange(existingIds[questionId], 1, 1, CONFIG_COLUMN_COUNT).setValues([row]);
    } else {
      // Add new row
      const newRowNum = sheet.getLastRow() + 1;
      sheet.getRange(newRowNum, 1, 1, CONFIG_COLUMN_COUNT).setValues([row]);
      this.existingIds[questionId] = newRowNum;
    }
    
    return { success: true };
  }

  /**
   * Save multiple question configurations
   * @param {Array<{questionId: string, config: Object}>} configsArray
   * @returns {{success: boolean, count: number}}
   */
  saveMultiple(configsArray) {
    const sheet = this.getSheet();
    const existingIds = this.getExistingIds();
    
    const updates = [];
    const newRows = [];
    
    configsArray.forEach(({ questionId, config }) => {
      const questionText = this.getQuestionText(questionId);
      const row = this.buildRow(questionId, config, questionText);
      
      if (existingIds[questionId]) {
        updates.push({ rowNumber: existingIds[questionId], data: row });
      } else {
        newRows.push(row);
      }
    });
    
    // Apply updates
    updates.forEach(({ rowNumber, data }) => {
      sheet.getRange(rowNumber, 1, 1, CONFIG_COLUMN_COUNT).setValues([data]);
    });
    
    // Add new rows in batch
    if (newRows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, CONFIG_COLUMN_COUNT).setValues(newRows);
    }
    
    this.clearIdCache();
    
    return { success: true, count: updates.length + newRows.length };
  }

  /**
   * Bulk save - clears existing and writes all new
   * @param {Object} configObj - { questionId: config, ... }
   * @returns {{success: boolean, rowCount: number}}
   */
  saveBulk(configObj) {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();
    
    // Clear existing data (keep headers)
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }
    
    // Build all rows
    const rows = Object.entries(configObj).map(([questionId, config]) => {
      const questionText = this.getQuestionText(questionId);
      return this.buildRow(questionId, config, questionText);
    });
    
    // Write in one batch
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, CONFIG_COLUMN_COUNT).setValues(rows);
    }
    
    this.clearIdCache();
    
    return { success: true, rowCount: rows.length };
  }

  /**
   * Save group configurations (all questions in groups get the same config)
   * @param {Array<{questionIds: string[], questionType: string, chartType: string, ...}>} groupConfigs
   * @returns {{success: boolean, count: number, updates: number, new: number}}
   */
  saveGroups(groupConfigs) {
    const sheet = this.getSheet();
    const existingIds = this.getExistingIds();
    
    const updates = [];
    const newRows = [];
    
    groupConfigs.forEach(groupConfig => {
      groupConfig.questionIds.forEach(questionId => {
        const questionText = this.getQuestionText(questionId);
        const row = this.buildRow(questionId, groupConfig, questionText);
        
        if (existingIds[questionId]) {
          updates.push({ rowNumber: existingIds[questionId], data: row });
        } else {
          newRows.push(row);
        }
      });
    });
    
    // Apply updates
    updates.forEach(({ rowNumber, data }) => {
      sheet.getRange(rowNumber, 1, 1, CONFIG_COLUMN_COUNT).setValues([data]);
    });
    
    // Add new rows in batch
    if (newRows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, CONFIG_COLUMN_COUNT).setValues(newRows);
    }
    
    this.clearIdCache();
    
    Logger.log(`Saved ${updates.length + newRows.length} configs (${updates.length} updates, ${newRows.length} new)`);
    
    return {
      success: true,
      count: updates.length + newRows.length,
      updates: updates.length,
      new: newRows.length
    };
  }

  /**
   * Load all configurations from the sheet
   * @returns {Object} - { questionId: config, ... }
   */
  loadAll() {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) return {};
    
    const data = sheet.getRange(2, 1, lastRow - 1, CONFIG_COLUMN_COUNT).getValues();
    
    return data.reduce((acc, row) => {
      const questionId = String(row[0]).trim();
      if (questionId) {
        acc[questionId] = {
          questionText: row[1] || '',
          questionType: row[2],
          chartType: row[3],
          mobileChartType: row[4] || '',
          hasOther: row[5] === true,
          sortType: row[6] || 'none',
          sortTypeMobile: row[7] || '',
          customOrder: row[8] || '',
          customOrderMobile: row[9] || '',
          colorGroup: row[10] || ''
        };
      }
      return acc;
    }, {});
  }

  /**
   * Delete a row from the config sheet
   * @param {number} rowNumber - 1-based row number
   * @returns {{success: boolean}}
   */
  deleteRow(rowNumber) {
    const sheet = this.getSheet();
    sheet.deleteRow(rowNumber);
    this.clearIdCache();
    return { success: true };
  }

  /**
   * Find and return duplicate question IDs in the config
   * @returns {{headers: string[], duplicates: Object}}
   */
  findDuplicates() {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 2) return { headers: [], duplicates: {} };
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // Map to track occurrences
    const idMap = {};
    
    data.forEach((row, index) => {
      const id = String(row[0]).trim();
      const realRowNumber = index + 2;
      
      if (id) {
        if (!idMap[id]) idMap[id] = [];
        idMap[id].push({ rowNumber: realRowNumber, values: row });
      }
    });
    
    // Filter to only duplicates
    const duplicates = {};
    Object.entries(idMap).forEach(([id, occurrences]) => {
      if (occurrences.length > 1) {
        duplicates[id] = occurrences;
      }
    });
    
    return { headers, duplicates };
  }

  /**
   * Get config data for upload to server
   * @returns {{headers: string[], data: Object[]}}
   */
  getUploadData() {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      throw new Error('No configuration data to upload.');
    }
    
    const [headers] = sheet.getRange(1, 1, 1, CONFIG_COLUMN_COUNT).getValues();
    const data = sheet.getRange(2, 1, lastRow - 1, CONFIG_COLUMN_COUNT).getValues();
    
    const payload = data.map(row => 
      headers.reduce((obj, h, i) => ({ ...obj, [h]: row[i] }), {})
    );
    
    return { headers, data: payload, count: payload.length };
  }
}

// === EXPORTED FUNCTIONS ===

/**
 * Ensure config sheet exists and return it
 */
function ensureConfigSheetExists() {
  const mgr = new ConfigManager();
  return mgr.getSheet();
}

/**
 * Load all configurations from the sheet
 */
function loadConfigFromSheet() {
  const mgr = new ConfigManager();
  return mgr.loadAll();
}

/**
 * Save a single question configuration
 */
function saveConfigToSheet(questionId, config) {
  const mgr = new ConfigManager();
  return mgr.save(questionId, config);
}

/**
 * Save multiple configurations at once
 */
function saveMultipleConfigsToSheet(configsArray) {
  const mgr = new ConfigManager();
  return mgr.saveMultiple(configsArray);
}

/**
 * Bulk save - replaces all existing config
 */
function saveConfigToSheetBulk(configObj) {
  const mgr = new ConfigManager();
  return mgr.saveBulk(configObj);
}

/**
 * Save group configurations
 */
function saveGroupConfigsToSheet(groupConfigs) {
  const mgr = new ConfigManager();
  return mgr.saveGroups(groupConfigs);
}

/**
 * Get duplicates in config sheet (for DuplicateCleaner)
 */
function getConfigDuplicates() {
  const mgr = new ConfigManager();
  return mgr.findDuplicates();
}

/**
 * Delete a config row (for DuplicateCleaner)
 */
function deleteConfigRow(rowNumber) {
  const mgr = new ConfigManager();
  return mgr.deleteRow(rowNumber);
}

/**
 * Get question configuration data for initialization
 */
function getQuestionConfigData() {
  const mgr = getDataManager();
  const props = PropertiesService.getDocumentProperties();
  const archived = JSON.parse(props.getProperty(ARCHIVED_QUESTIONS_KEY) || '[]');
  
  return {
    questions: mgr.getAllQuestions(archived),
    links: JSON.parse(props.getProperty(SCRIPT_PROPERTY_KEY) || '{}'),
    config: loadConfigFromSheet()
  };
}
