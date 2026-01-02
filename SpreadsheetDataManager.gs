/**
 * SpreadsheetDataManager.gs
 * Centralized manager for all spreadsheet data operations
 * Provides caching, consistent error handling, and simplified API
 */

class SpreadsheetDataManager {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheetCache = {};
    this.idCache = {};
    this.colIndexCache = {};
    this.headerCache = {};
  }

  // === SHEET ACCESS ===

  /**
   * Get sheet by name with caching
   * @param {string} sheetName
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   * @throws {Error} if sheet not found
   */
  getSheet(sheetName) {
    if (!this.sheetCache[sheetName]) {
      this.sheetCache[sheetName] = this.ss.getSheetByName(sheetName);
      if (!this.sheetCache[sheetName]) {
        throw new Error(`Sheet not found: ${sheetName}`);
      }
    }
    return this.sheetCache[sheetName];
  }

  /**
   * Get sheet by name, returning null if not found (no throw)
   * @param {string} sheetName
   * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
   */
  getSheetSafe(sheetName) {
    try {
      return this.getSheet(sheetName);
    } catch (e) {
      return null;
    }
  }

  /**
   * Get all data sheets (excluding Links, Config, and internal sheets)
   * @returns {GoogleAppsScript.Spreadsheet.Sheet[]}
   */
  getDataSheets() {
    return this.ss.getSheets().filter(isDataSheet);
  }

  // === QUESTION ID ACCESS ===

  /**
   * Get all question IDs from a sheet (row 2) with caching
   * @param {string} sheetName
   * @returns {string[]}
   */
  getSheetQuestionIds(sheetName) {
    if (!this.idCache[sheetName]) {
      const sheet = this.getSheet(sheetName);
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0 || sheet.getLastRow() < 2) {
        this.idCache[sheetName] = [];
      } else {
        this.idCache[sheetName] = sheet.getRange(2, 1, 1, lastCol)
          .getValues()[0]
          .map(id => String(id).trim());
      }
    }
    return this.idCache[sheetName];
  }

  /**
   * Get question headers (row 1) from a sheet with caching
   * @param {string} sheetName
   * @returns {string[]}
   */
  getSheetHeaders(sheetName) {
    if (!this.headerCache[sheetName]) {
      const sheet = this.getSheet(sheetName);
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0 || sheet.getLastRow() < 1) {
        this.headerCache[sheetName] = [];
      } else {
        this.headerCache[sheetName] = sheet.getRange(1, 1, 1, lastCol)
          .getValues()[0]
          .map(h => String(h || ''));
      }
    }
    return this.headerCache[sheetName];
  }

  /**
   * Extract sheet name from question ID (assumes format: sheetName_columnNumber)
   * @param {string} questionId
   * @returns {string}
   */
  getSheetNameFromId(questionId) {
    return questionId.split('_')[0];
  }

  /**
   * Find column info for a question ID
   * @param {string} questionId
   * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, sheetName: string, colIndex: number, questionText: string}}
   * @throws {Error} if question not found
   */
  findQuestionColumn(questionId) {
    if (this.colIndexCache[questionId]) {
      return this.colIndexCache[questionId];
    }

    const sheetName = this.getSheetNameFromId(questionId);
    const ids = this.getSheetQuestionIds(sheetName);
    const colIndex = ids.findIndex(id => id === questionId);

    if (colIndex === -1) {
      throw new Error(`Question not found: ${questionId}`);
    }

    const sheet = this.getSheet(sheetName);
    const headers = this.getSheetHeaders(sheetName);
    const questionText = headers[colIndex] || '';

    const result = { sheet, sheetName, colIndex, questionText };
    this.colIndexCache[questionId] = result;
    return result;
  }

  /**
   * Find column info without throwing (returns null if not found)
   * @param {string} questionId
   * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, sheetName: string, colIndex: number, questionText: string}|null}
   */
  findQuestionColumnSafe(questionId) {
    try {
      return this.findQuestionColumn(questionId);
    } catch (e) {
      return null;
    }
  }

  // === DATA ACCESS ===

  /**
   * Read all values from a question column (excluding header rows 1-2)
   * @param {string} questionId
   * @returns {any[]}
   */
  getQuestionValues(questionId) {
    const { sheet, colIndex } = this.findQuestionColumn(questionId);
    const lastRow = sheet.getLastRow();

    if (lastRow < 3) return [];

    return sheet.getRange(3, colIndex + 1, lastRow - 2, 1)
      .getValues()
      .map(row => row[0]);
  }

  /**
   * Get unique answers for a question
   * @param {string} questionId
   * @param {Object} options - { sorted: true, withCounts: false, trim: true, excludeEmpty: true }
   * @returns {string[]|{value: string, count: number}[]}
   */
  getUniqueAnswers(questionId, options = {}) {
    const opts = {
      sorted: true,
      withCounts: false,
      trim: true,
      excludeEmpty: true,
      ...options
    };

    const values = this.getQuestionValues(questionId);
    const counts = {};

    values.forEach(val => {
      if (val === null || val === undefined) return;

      let v = String(val);
      if (opts.trim) v = v.trim();
      if (opts.excludeEmpty && !v) return;

      counts[v] = (counts[v] || 0) + 1;
    });

    if (opts.withCounts) {
      const result = Object.entries(counts).map(([value, count]) => ({ value, count }));
      return opts.sorted ? result.sort((a, b) => b.count - a.count) : result;
    } else {
      const result = Object.keys(counts);
      return opts.sorted ? result.sort() : result;
    }
  }

  /**
   * Get numerical statistics for a question
   * @param {string} questionId
   * @returns {{values: {value: number, count: number}[], stats: {min: number, max: number, mean: number, median: number, count: number}}}
   */
  getNumericalStats(questionId) {
    const rawValues = this.getQuestionValues(questionId);
    const values = rawValues
      .map(v => parseFloat(v))
      .filter(v => !isNaN(v))
      .sort((a, b) => a - b);

    if (values.length === 0) {
      return { values: [], stats: {} };
    }

    const sum = values.reduce((a, b) => a + b, 0);
    const counts = values.reduce((acc, v) => ({ ...acc, [v]: (acc[v] || 0) + 1 }), {});

    return {
      values: Object.entries(counts)
        .map(([value, count]) => ({ value: parseFloat(value), count }))
        .sort((a, b) => a.value - b.value),
      stats: {
        min: values[0],
        max: values[values.length - 1],
        mean: sum / values.length,
        median: values[Math.floor(values.length / 2)],
        count: values.length
      }
    };
  }

  // === DATA MODIFICATION ===

  /**
   * Replace values in a question column
   * @param {string} questionId
   * @param {Object} replaceMap - { oldValue: newValue }
   * @param {Object} options - { splitByComma: false, dryRun: false }
   * @returns {{replacements: number, modified: Array, dryRun: boolean}}
   */
  replaceValues(questionId, replaceMap, options = {}) {
    const opts = { splitByComma: false, dryRun: false, ...options };
    const { sheet, colIndex } = this.findQuestionColumn(questionId);
    const lastRow = sheet.getLastRow();

    if (lastRow < 3) return { replacements: 0, modified: [], dryRun: opts.dryRun };

    const dataRange = sheet.getRange(3, colIndex + 1, lastRow - 2, 1);
    const values = dataRange.getValues();
    let changeCount = 0;
    const modified = [];

    values.forEach((row, i) => {
      const cellValue = String(row[0]).trim();
      if (!cellValue) return;

      let newValue;

      if (opts.splitByComma) {
        const parts = cellValue.split(',').map(p => p.trim());
        const newParts = parts
          .map(part => replaceMap.hasOwnProperty(part) ? replaceMap[part] : part)
          .filter(p => p !== '');
        newValue = newParts.join(', ');
      } else {
        newValue = replaceMap.hasOwnProperty(cellValue) ? replaceMap[cellValue] : cellValue;
      }

      if (newValue !== cellValue) {
        values[i][0] = newValue;
        changeCount++;
        modified.push({ row: i + 3, old: cellValue, new: newValue });
      }
    });

    if (changeCount > 0 && !opts.dryRun) {
      dataRange.setValues(values);
    }

    return {
      replacements: changeCount,
      modified: opts.dryRun ? modified : [],
      dryRun: opts.dryRun
    };
  }

  /**
   * Batch replace values across multiple questions with rate limiting
   * @param {string[]} questionIds
   * @param {Object} replaceMap
   * @param {Object} options
   * @returns {{totalReplacements: number, questionsProcessed: number, results: Array}}
   */
  batchReplaceValues(questionIds, replaceMap, options = {}) {
    const opts = {
      batchSize: RATE_LIMIT_BATCH_SIZE,
      delayBetweenBatches: BATCH_WRITE_DELAY_MS,
      ...options
    };

    let totalReplacements = 0;
    const results = [];

    for (let i = 0; i < questionIds.length; i += opts.batchSize) {
      const batch = questionIds.slice(i, i + opts.batchSize);

      batch.forEach(questionId => {
        try {
          const result = this.replaceValues(questionId, replaceMap, options);
          totalReplacements += result.replacements;
          results.push({ questionId, ...result });
        } catch (err) {
          Logger.log(`Error processing ${questionId}: ${err.message}`);
          results.push({ questionId, error: err.message });
        }
      });

      if (i + opts.batchSize < questionIds.length) {
        Utilities.sleep(opts.delayBetweenBatches);
      }
    }

    return {
      totalReplacements,
      questionsProcessed: questionIds.length,
      results
    };
  }

  // === COLUMN OPERATIONS ===

  /**
   * Hide a question column in its sheet
   * @param {string} questionId
   * @returns {{success: boolean, message?: string}}
   */
  hideColumn(questionId) {
    try {
      const { sheet, colIndex } = this.findQuestionColumn(questionId);
      sheet.hideColumns(colIndex + 1);
      return { success: true };
    } catch (e) {
      return { success: false, message: e.message };
    }
  }

  /**
   * Show a hidden question column
   * @param {string} questionId
   * @returns {{success: boolean, message?: string}}
   */
  showColumn(questionId) {
    try {
      const { sheet, colIndex } = this.findQuestionColumn(questionId);
      sheet.showColumns(colIndex + 1);
      return { success: true };
    } catch (e) {
      return { success: false, message: e.message };
    }
  }

  /**
   * Delete a question column from its sheet
   * @param {string} questionId
   * @returns {{success: boolean, message?: string}}
   */
  deleteColumn(questionId) {
    try {
      const { sheet, colIndex, sheetName } = this.findQuestionColumn(questionId);
      sheet.deleteColumn(colIndex + 1);
      
      // Invalidate caches for this sheet
      delete this.idCache[sheetName];
      delete this.headerCache[sheetName];
      delete this.colIndexCache[questionId];
      
      return { success: true };
    } catch (e) {
      return { success: false, message: e.message };
    }
  }

  /**
   * Format a column as plain text
   * @param {string} questionId
   * @returns {{success: boolean}}
   */
  formatAsPlainText(questionId) {
    const { sheet, colIndex } = this.findQuestionColumn(questionId);
    const lastRow = sheet.getLastRow();
    if (lastRow >= 3) {
      sheet.getRange(3, colIndex + 1, lastRow - 2, 1).setNumberFormat('@');
    }
    return { success: true };
  }

  // === BULK OPERATIONS ===

  /**
   * Process multiple questions efficiently, grouping by sheet
   * @param {Array<{id: string, sheet: string}>} questions - Array of question objects
   * @param {boolean} readOnly - If true, don't write changes
   * @param {Function} cellCallback - (value, rowIndex) => newValue (for writes) or void (for reads)
   * @param {Function} perQuestionCallback - (question, value, rowIndex) => void (for reads only)
   */
  processDataInColumns(questions, readOnly, cellCallback, perQuestionCallback) {
    const bySheet = questions.reduce((acc, q) => {
      if (!acc[q.sheet]) acc[q.sheet] = [];
      acc[q.sheet].push(q);
      return acc;
    }, {});

    Object.entries(bySheet).forEach(([sheetName, qs]) => {
      const sheet = this.getSheetSafe(sheetName);
      if (!sheet) return;

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 3 || lastCol === 0) return;

      const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      const ids = allData[1];
      const numRows = lastRow - 2;

      // Collect pending writes for this sheet
      const pendingWrites = [];

      qs.forEach(q => {
        const colIndex = ids.indexOf(q.id);
        if (colIndex === -1) return;

        const colValues = allData.slice(2).map(row => [row[colIndex]]);
        let modified = false;

        colValues.forEach((cell, i) => {
          const val = cell[0];
          if (val === null || val === undefined || val === '') return;

          if (readOnly) {
            if (perQuestionCallback) perQuestionCallback(q, val, i);
            else if (cellCallback) cellCallback(val, i);
          } else {
            const newVal = cellCallback(val, i);
            if (newVal !== val) {
              cell[0] = newVal;
              modified = true;
            }
          }
        });

        if (!readOnly && modified) {
          pendingWrites.push({
            range: sheet.getRange(3, colIndex + 1, numRows, 1),
            values: colValues
          });
        }
      });

      // Batch write with rate limiting
      if (!readOnly && pendingWrites.length > 0) {
        for (let i = 0; i < pendingWrites.length; i += RATE_LIMIT_BATCH_SIZE) {
          const chunk = pendingWrites.slice(i, i + RATE_LIMIT_BATCH_SIZE);
          chunk.forEach(write => write.range.setValues(write.values));
          SpreadsheetApp.flush();

          if (i + RATE_LIMIT_BATCH_SIZE < pendingWrites.length) {
            Utilities.sleep(RATE_LIMIT_DELAY_MS);
          }
        }
      }
    });
  }

  /**
   * Get all questions from all data sheets
   * @param {string[]} excludeIds - Question IDs to exclude (e.g., archived)
   * @returns {Array<{questionId: string, questionText: string, year: string}>}
   */
  getAllQuestions(excludeIds = []) {
    const excludeSet = new Set(excludeIds);
    
    return this.getDataSheets().flatMap(sheet => {
      const sheetName = sheet.getName();
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0 || sheet.getLastRow() < 2) return [];

      const [texts, ids] = sheet.getRange(1, 1, 2, lastCol).getValues();
      
      return ids.map((id, i) => {
        const qId = String(id).trim();
        return qId && !excludeSet.has(qId)
          ? { questionId: qId, questionText: texts[i] || '', year: sheetName }
          : null;
      }).filter(Boolean);
    });
  }

  // === CACHE MANAGEMENT ===

  /**
   * Clear all caches (call when sheet structure changes)
   */
  clearCache() {
    this.sheetCache = {};
    this.idCache = {};
    this.colIndexCache = {};
    this.headerCache = {};
  }

  /**
   * Clear cache for a specific sheet
   * @param {string} sheetName
   */
  clearSheetCache(sheetName) {
    delete this.sheetCache[sheetName];
    delete this.idCache[sheetName];
    delete this.headerCache[sheetName];
    
    // Clear column cache for any questions from this sheet
    Object.keys(this.colIndexCache).forEach(qId => {
      if (this.getSheetNameFromId(qId) === sheetName) {
        delete this.colIndexCache[qId];
      }
    });
  }
}

// === SINGLETON INSTANCE ===
// Use this for better performance across multiple function calls in the same execution

let _dataManagerInstance = null;

/**
 * Get or create the singleton SpreadsheetDataManager instance
 * @returns {SpreadsheetDataManager}
 */
function getDataManager() {
  if (!_dataManagerInstance) {
    _dataManagerInstance = new SpreadsheetDataManager();
  }
  return _dataManagerInstance;
}

/**
 * Reset the singleton instance (use when spreadsheet structure changes)
 */
function resetDataManager() {
  if (_dataManagerInstance) {
    _dataManagerInstance.clearCache();
  }
  _dataManagerInstance = null;
}
