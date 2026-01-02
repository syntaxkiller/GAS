/**
 * DataOperations.gs
 * Data manipulation operations (faceting, mass editing, column combining)
 */

// === FACET DATA OPERATIONS ===

/**
 * Get facet data for a group or question
 * @param {Object} options - { groupName?, questionId?, splitByComma? }
 * @returns {{uniqueValues: string[], valueCounts: Array}}
 */
function getFacetData(options) {
  const { splitByComma } = options;
  const questions = getQuestionsToProcess_(options);

  Logger.log('getFacetData - Options: ' + JSON.stringify(options));
  Logger.log('getFacetData - Questions found: ' + questions.length);

  if (!questions || questions.length === 0) {
    return { error: 'No questions found' };
  }

  const mgr = getDataManager();
  const valueCounts = {};
  let totalValuesProcessed = 0;

  mgr.processDataInColumns(questions, true, null, (q, val) => {
    const v = String(val || '').trim();
    if (!v) return;

    totalValuesProcessed++;

    if (splitByComma) {
      const parts = v.split(',').map(part => part.trim()).filter(part => part);
      parts.forEach(part => {
        valueCounts[part] = (valueCounts[part] || 0) + 1;
      });
    } else {
      valueCounts[v] = (valueCounts[v] || 0) + 1;
    }
  });

  Logger.log('getFacetData - Total values processed: ' + totalValuesProcessed);
  Logger.log('getFacetData - Unique values found: ' + Object.keys(valueCounts).length);

  const uniqueValues = Object.keys(valueCounts).sort();
  const valueCountsArray = Object.entries(valueCounts)
    .map(([text, count]) => ({ text, count }))
    .sort((a, b) => b.count - a.count);

  return { uniqueValues, valueCounts: valueCountsArray };
}

/**
 * Get full preview data organized by year
 * @param {Object} options - { groupName?, questionId? }
 * @returns {Object} - { year: [answers], ... }
 */
function getFullPreviewData(options) {
  const questions = getQuestionsToProcess_(options);

  if (!questions || questions.length === 0) {
    return {};
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const previewData = {};

  questions.forEach(q => {
    const sheet = ss.getSheetByName(q.sheet);
    if (!sheet) return;

    const year = q.sheet;
    if (!previewData[year]) {
      previewData[year] = [];
    }

    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    if (lastCol === 0 || lastRow < 2) return;

    const ids = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    const colIndex = ids.findIndex(id => String(id).trim() === q.id);

    if (colIndex === -1) return;

    if (lastRow >= 3) {
      const answers = sheet.getRange(3, colIndex + 1, lastRow - 2, 1).getValues();
      const uniqueAnswers = new Set();

      answers.forEach(([val]) => {
        const v = String(val || '').trim();
        if (v) uniqueAnswers.add(v);
      });

      Array.from(uniqueAnswers).forEach(answer => {
        if (!previewData[year].includes(answer)) {
          previewData[year].push(answer);
        }
      });
    }
  });

  // Sort answers within each year
  Object.keys(previewData).forEach(year => {
    previewData[year].sort();
  });

  return previewData;
}

// === MASS EDIT OPERATIONS ===

/**
 * Apply mass find/replace edits
 * @param {Object} options
 * @returns {{success: boolean, message: string}}
 */
function applyMassEdit(options) {
  const { groupName, questionId, findReplace, caseInsensitive, splitByComma } = options;

  if (!findReplace || Object.keys(findReplace).length === 0) {
    return { success: false, message: 'No replacements provided.' };
  }

  const questions = getQuestionsToProcess_({ groupName, questionId });
  if (!questions || questions.length === 0) {
    return { success: false, message: 'No questions found.' };
  }

  let totalChanged = 0;
  const normalized = Object.entries(findReplace).reduce((acc, [k, v]) => {
    acc[caseInsensitive ? k.toLowerCase() : k] = v;
    return acc;
  }, {});

  const mgr = getDataManager();
  
  mgr.processDataInColumns(questions, false, (val) => {
    const str = String(val).trim();

    if (splitByComma) {
      const parts = str.split(',').map(p => p.trim());
      let changed = false;

      const newParts = parts.map(part => {
        const key = caseInsensitive ? part.toLowerCase() : part;
        if (normalized.hasOwnProperty(key)) {
          changed = true;
          totalChanged++;
          return normalized[key];
        }
        return part;
      });

      return changed ? newParts.join(', ') : val;
    } else {
      const key = caseInsensitive ? str.toLowerCase() : str;
      if (normalized.hasOwnProperty(key)) {
        totalChanged++;
        return normalized[key];
      }
      return val;
    }
  });

  return {
    success: true,
    message: `Changed ${totalChanged} ${splitByComma ? 'value(s)' : 'cell(s)'} across ${questions.length} question(s).`
  };
}

/**
 * Replace text values with options
 * @param {string} questionId
 * @param {Array} replacements - [{ oldValue, newValue }, ...]
 * @param {boolean} splitByComma
 * @returns {{success: boolean, replacements: number}}
 */
function replaceTextValuesWithOptions(questionId, replacements, splitByComma) {
  const mgr = getDataManager();

  const replaceMap = {};
  replacements.forEach(r => {
    replaceMap[r.oldValue] = r.newValue;
  });

  const result = mgr.replaceValues(questionId, replaceMap, { splitByComma });
  return { success: true, replacements: result.replacements };
}

/**
 * Replace text values across multiple questions with rate limiting
 * @param {string[]} questionIds
 * @param {Array} replacements
 * @param {boolean} splitByComma
 * @returns {{success: boolean, replacements: number, questionsProcessed: number}}
 */
function replaceTextValuesRateOptimized(questionIds, replacements, splitByComma) {
  const mgr = getDataManager();

  const replaceMap = {};
  replacements.forEach(r => {
    replaceMap[r.oldValue] = r.newValue;
  });

  const result = mgr.batchReplaceValues(questionIds, replaceMap, {
    splitByComma,
    batchSize: RATE_LIMIT_BATCH_SIZE,
    delayBetweenBatches: BATCH_WRITE_DELAY_MS
  });

  return {
    success: true,
    replacements: result.totalReplacements,
    questionsProcessed: result.questionsProcessed
  };
}

/**
 * Format columns as plain text
 * @param {Object} options - { groupName?, questionId? }
 * @returns {{success: boolean, message: string}}
 */
function formatAsPlainText(options) {
  const { groupName, questionId } = options;

  const questions = getQuestionsToProcess_({ groupName, questionId });
  if (!questions || questions.length === 0) {
    return { success: false, message: 'No questions found.' };
  }

  const mgr = getDataManager();
  let totalFormatted = 0;

  questions.forEach(q => {
    try {
      mgr.formatAsPlainText(q.id);
      totalFormatted++;
    } catch (e) {
      Logger.log(`Could not format ${q.id}: ${e.message}`);
    }
  });

  return {
    success: true,
    message: `Formatted ${totalFormatted} column(s) as plain text.`
  };
}

// === COLUMN COMBINING OPERATIONS ===

/**
 * Combine checkbox columns into a single column
 * @param {Object} options
 * @returns {{success: boolean, message: string}}
 */
function combineCheckboxColumns(options) {
  const { replacements, otherColumnHeader, newColumnHeader, sheetName, idsToArchive } = options;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { success: false, message: 'Sheet not found: ' + sheetName };
  }

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow < 3 || lastCol === 0) {
    return { success: false, message: 'Sheet has no data to combine.' };
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const ids = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  // Build column index map
  const colMap = {};
  replacements.forEach(r => {
    const idx = headers.indexOf(r.header);
    if (idx >= 0) colMap[r.header] = { index: idx, text: r.text };
  });

  // Find other column if specified
  let otherColIndex = -1;
  if (otherColumnHeader) {
    otherColIndex = headers.indexOf(otherColumnHeader);
  }

  // Create combined column values
  const combinedValues = data.map(row => {
    const values = [];

    Object.entries(colMap).forEach(([header, info]) => {
      const cellVal = row[info.index];
      if (cellVal === true || String(cellVal).toLowerCase() === 'true' || cellVal === 1) {
        values.push(info.text);
      }
    });

    // Add "other" value if present
    if (otherColIndex >= 0 && row[otherColIndex]) {
      const otherVal = String(row[otherColIndex]).trim();
      if (otherVal && otherVal.toLowerCase() !== 'false') {
        values.push(otherVal);
      }
    }

    return [values.join(', ')];
  });

  // Add new column
  const newColIndex = lastCol + 1;
  sheet.getRange(1, newColIndex).setValue(newColumnHeader);
  sheet.getRange(2, newColIndex).setValue(`${sheetName}_${newColIndex}`);
  sheet.getRange(3, newColIndex, combinedValues.length, 1).setValues(combinedValues);

  // Archive source columns
  if (idsToArchive && idsToArchive.length > 0) {
    archiveQuestionIds_(idsToArchive);
  }

  resetDataManager();

  return {
    success: true,
    message: `Created "${newColumnHeader}" column with ${combinedValues.length} rows. ${idsToArchive.length} source column(s) archived.`
  };
}

/**
 * Combine two single-choice columns
 * @param {Object} options
 * @returns {{success: boolean, message: string}}
 */
function combineSingleChoiceColumns(options) {
  const { selectedHeaders, sheetName, idsToArchive } = options;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { success: false, message: 'Sheet not found: ' + sheetName };
  }

  if (!selectedHeaders || selectedHeaders.length !== 2) {
    return { success: false, message: 'Exactly two columns must be selected.' };
  }

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow < 3 || lastCol === 0) {
    return { success: false, message: 'Sheet has no data to combine.' };
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  const col1Idx = headers.indexOf(selectedHeaders[0]);
  const col2Idx = headers.indexOf(selectedHeaders[1]);

  if (col1Idx === -1 || col2Idx === -1) {
    return { success: false, message: 'Could not find selected columns in sheet.' };
  }

  // Combine: prefer first column, fallback to second
  const combinedValues = data.map(row => {
    const val1 = String(row[col1Idx] || '').trim();
    const val2 = String(row[col2Idx] || '').trim();
    return [val1 || val2];
  });

  // Write to first column (overwrite)
  sheet.getRange(3, col1Idx + 1, combinedValues.length, 1).setValues(combinedValues);

  // Archive second column
  if (idsToArchive && idsToArchive.length > 1) {
    archiveQuestionIds_([idsToArchive[1]]);
  }

  resetDataManager();

  return {
    success: true,
    message: `Combined columns into "${selectedHeaders[0]}". Second column archived.`
  };
}

/**
 * Archive question IDs by adding them to the archived list
 * @param {string[]} ids
 * @private
 */
function archiveQuestionIds_(ids) {
  const props = PropertiesService.getDocumentProperties();
  const archived = JSON.parse(props.getProperty(ARCHIVED_QUESTIONS_KEY) || '[]');

  ids.forEach(id => {
    if (!archived.includes(id)) {
      archived.push(id);
    }
  });

  props.setProperty(ARCHIVED_QUESTIONS_KEY, JSON.stringify(archived));

  // Also remove from hidden if present
  const hidden = JSON.parse(props.getProperty(HIDDEN_QUESTIONS_KEY) || '[]');
  const newHidden = hidden.filter(id => !ids.includes(id));
  props.setProperty(HIDDEN_QUESTIONS_KEY, JSON.stringify(newHidden));
}

// === ALL QUESTIONS LIST ===

/**
 * Get all question list (for QuestionConfig)
 * @returns {Array}
 */
function getAllQuestionList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const archived = JSON.parse(props.getProperty(ARCHIVED_QUESTIONS_KEY) || '[]');

  return ss.getSheets()
    .filter(isDataSheet)
    .flatMap(sheet => {
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0 || sheet.getLastRow() < 2) return [];

      const [texts, ids] = sheet.getRange(1, 1, 2, lastCol).getValues();
      
      return ids.map((id, i) => {
        const qId = String(id).trim();
        return qId && !archived.includes(qId)
          ? { questionId: qId, questionText: texts[i] || '', year: sheet.getName() }
          : null;
      }).filter(Boolean);
    });
}
