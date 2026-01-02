/**
 * Centralized manager for spreadsheet data operations
 * Provides caching, consistent error handling, and simplified API
 */
class SpreadsheetDataManager {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheetCache = {};
    this.idCache = {};
    this.colIndexCache = {};
  }
  
  /**
   * Get sheet by name (with caching)
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
   * Get all question IDs from a sheet (with caching)
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
   * Find column index for a question ID
   * Returns: { sheet, sheetName, colIndex, questionText }
   */
  findQuestionColumn(questionId) {
    if (this.colIndexCache[questionId]) {
      return this.colIndexCache[questionId];
    }
    
    const [sheetName] = questionId.split('_');
    const ids = this.getSheetQuestionIds(sheetName);
    const colIndex = ids.findIndex(id => id === questionId);
    
    if (colIndex === -1) {
      throw new Error(`Question not found: ${questionId}`);
    }
    
    const sheet = this.getSheet(sheetName);
    const questionText = sheet.getRange(1, colIndex + 1).getValue();
    
    const result = { sheet, sheetName, colIndex, questionText };
    this.colIndexCache[questionId] = result;
    return result;
  }
  
  /**
   * Read all values from a question column
   * Returns: array of values (excluding header rows 1-2)
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
   * Options: { sorted: true, withCounts: false, trim: true, excludeEmpty: true }
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
   * Replace values in a question column
   * replaceMap: { oldValue: newValue }
   * Options: { splitByComma: false, dryRun: false }
   */
  replaceValues(questionId, replaceMap, options = {}) {
    const opts = { splitByComma: false, dryRun: false, ...options };
    const { sheet, colIndex } = this.findQuestionColumn(questionId);
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 3) return { replacements: 0, modified: [] };
    
    const dataRange = sheet.getRange(3, colIndex + 1, lastRow - 2, 1);
    const values = dataRange.getValues();
    let changeCount = 0;
    const modified = [];
    
    values.forEach((row, i) => {
      const cellValue = String(row[0]).trim();
      if (!cellValue) return;
      
      let newValue;
      
      if (opts.splitByComma) {
        const parts = cellValue.split(',').map(p => p.trimEnd());
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
   * Batch replace values across multiple questions
   */
  batchReplaceValues(questionIds, replaceMap, options = {}) {
    const opts = { 
      batchSize: 5, 
      delayBetweenBatches: 500,
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
  
  /**
   * Process multiple questions efficiently by grouping by sheet
   * callback: (questionId, values, metadata) => { modified: boolean, newValues: [] }
   */
  processQuestions(questions, callback) {
    const bySheet = questions.reduce((acc, qId) => {
      const [sheetName] = qId.split('_');
      if (!acc[sheetName]) acc[sheetName] = [];
      acc[sheetName].push(qId);
      return acc;
    }, {});
    
    const results = {};
    
    Object.entries(bySheet).forEach(([sheetName, qIds]) => {
      const sheet = this.getSheet(sheetName);
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      if (lastRow < 3 || lastCol === 0) return;
      
      const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      const [texts, ids] = [allData[0], allData[1]];
      
      qIds.forEach(qId => {
        const colIndex = ids.findIndex(id => String(id).trim() === qId);
        if (colIndex === -1) return;
        
        const values = allData.slice(2).map(row => row[colIndex]);
        const metadata = {
          questionId: qId,
          questionText: texts[colIndex],
          sheet: sheetName,
          colIndex
        };
        
        const result = callback(qId, values, metadata);
        results[qId] = result;
        
        if (result.modified && result.newValues) {
          const range = sheet.getRange(3, colIndex + 1, result.newValues.length, 1);
          range.setValues(result.newValues.map(v => [v]));
        }
      });
    });
    
    return results;
  }
  
  /**
   * Clear all caches (call when structure changes)
   */
  clearCache() {
    this.sheetCache = {};
    this.idCache = {};
    this.colIndexCache = {};
  }
}