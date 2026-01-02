/**
 * Core.gs
 * Main application logic for survey processing tool
 * @OnlyCurrentDoc
 */

// === MENU & UI ===

/**
 * Create the Survey Tools menu on spreadsheet open
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Survey Tools')
    .addItem('🔗 Link Questions', 'showLinkingDialog')
    .addItem('📝 Chart Config', 'showQuestionConfigDialog')
    .addItem('⚡ Manage Duplicates', 'showDuplicateManager')
    .addSeparator()
    .addItem('Push All Data to Website', 'pushAllDataToServer')
    .addItem('Push Question Config Only', 'pushQuestionConfigOnly')
    .addSeparator()
    .addItem('Debug: View Stored Data', 'viewStoredData')
    .addToUi();
}

/**
 * Show the question linking dialog
 */
function showLinkingDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('Sidebar').evaluate().setWidth(1600).setHeight(900),
    'Survey Tools'
  );
}

/**
 * Show the question configuration dialog
 */
function showQuestionConfigDialog() {
  ensureConfigSheetExists();
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('QuestionConfig').evaluate().setWidth(1600).setHeight(900),
    'Question Configuration'
  );
}

/**
 * Show the duplicate manager dialog
 */
function showDuplicateManager() {
  const html = HtmlService.createTemplateFromFile('DuplicateCleanerUI')
    .evaluate()
    .setTitle('Duplicate Question Manager')
    .setWidth(800)
    .setHeight(600);

  SpreadsheetApp.getUi().showModelessDialog(html, 'Duplicate Question Configuration Manager');
}

// === DATA INITIALIZATION ===

/**
 * Ensure question IDs exist in row 2 of all data sheets
 * @returns {{success: boolean, message: string}}
 */
function ensureQuestionIdsExist() {
  const props = PropertiesService.getDocumentProperties();
  
  if (props.getProperty(IDS_VERIFIED_FLAG)) {
    return { success: true, message: 'IDs already verified.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ss.getSheets()
    .filter(isDataSheet)
    .forEach(sheet => {
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0 || sheet.getLastRow() === 0) return;

      const row2 = sheet.getMaxRows() > 1 
        ? sheet.getRange(2, 1, 1, lastCol).getValues()[0] 
        : [];
      const idPattern = new RegExp(`^${sheet.getName()}_\\d+$`);

      if (row2.length === 0 || !idPattern.test(String(row2[0]))) {
        sheet.insertRowBefore(2);
        const questions = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const ids = questions.map((q, i) => q ? `${sheet.getName()}_${i + 1}` : '');
        sheet.getRange(2, 1, 1, ids.length).setValues([ids]);
      }
    });

  props.setProperty(IDS_VERIFIED_FLAG, 'true');
  SpreadsheetApp.flush();
  return { success: true, message: 'IDs verified.' };
}

/**
 * Get all initial data for the sidebar
 * @returns {Array} [sheetsData, linksJson, hiddenJson, archivedJson]
 */
function getSidebarInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsData = {};

  ss.getSheets()
    .filter(isDataSheet)
    .forEach(sheet => {
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0 || sheet.getLastRow() < 2) return;

      const [questions, ids] = sheet.getRange(1, 1, 2, lastCol).getValues();
      let prevText = null;

      sheetsData[sheet.getName()] = questions.map((text, i) => {
        const result = text && ids[i] ? {
          question: text,
          id: String(ids[i]),
          sheet: sheet.getName(),
          previousQuestion: prevText
        } : null;
        
        prevText = text || null;
        return result;
      }).filter(Boolean);
    });

  const props = PropertiesService.getDocumentProperties();
  
  return [
    sheetsData,
    props.getProperty(SCRIPT_PROPERTY_KEY) || '{}',
    props.getProperty(HIDDEN_QUESTIONS_KEY) || '[]',
    props.getProperty(ARCHIVED_QUESTIONS_KEY) || '[]'
  ];
}

/**
 * Get active sheet data
 * @returns {{sheetName: string, headers: string[]}}
 */
function getActiveSheetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return {
    sheetName: sheet.getName(),
    headers: sheet.getLastRow() === 0 ? [] : sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
  };
}

/**
 * Set the active sheet
 * @param {string} sheetName
 */
function setActiveSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet) {
    SpreadsheetApp.setActiveSheet(sheet);
  }
}

// === QUESTION OPERATIONS ===

/**
 * Get answer options for a question
 * @param {string} questionId
 * @returns {Array}
 */
function getAnswerOptions(questionId) {
  const mgr = getDataManager();
  return mgr.getUniqueAnswers(questionId, { withCounts: true, sorted: true });
}

/**
 * Get unique answers for a question
 * @param {string} questionId
 * @returns {string[]}
 */
function getUniqueAnswers(questionId) {
  const mgr = getDataManager();
  return mgr.getUniqueAnswers(questionId);
}

/**
 * Get numerical statistics for a question
 * @param {string} questionId
 * @returns {Object}
 */
function getNumericalStats(questionId) {
  const mgr = getDataManager();
  return mgr.getNumericalStats(questionId);
}

/**
 * Get answer options for multiple questions
 * @param {string[]} questionIds
 * @returns {Object}
 */
function getMultipleAnswerOptions(questionIds) {
  return questionIds.reduce((result, qId) => {
    try {
      result[qId] = { answers: getAnswerOptions(qId) };
    } catch (err) {
      result[qId] = { error: err.message, answers: [] };
    }
    return result;
  }, {});
}

/**
 * Compare answers across multiple questions
 * @param {string[]} questionIds
 * @returns {Object}
 */
function compareQuestionAnswers(questionIds) {
  const mgr = getDataManager();
  
  const results = questionIds.reduce((acc, qId) => {
    try {
      acc[qId] = { answers: mgr.getUniqueAnswers(qId) };
    } catch (err) {
      acc[qId] = { error: err.message, answers: [] };
    }
    return acc;
  }, {});

  if (questionIds.length > 1) {
    const sets = questionIds.map(qId => new Set(results[qId].answers));
    const common = Array.from(sets[0]).filter(ans => sets.every(s => s.has(ans)));

    results._comparison = {
      totalQuestions: questionIds.length,
      commonAnswers: common,
      commonCount: common.length,
      pairwiseMatches: {}
    };

    // Calculate pairwise matches
    for (let i = 0; i < questionIds.length; i++) {
      for (let j = i + 1; j < questionIds.length; j++) {
        const stats = getIntersectionStats(results[questionIds[i]].answers, results[questionIds[j]].answers);
        const percentage = Math.round((stats.matchCount / Math.max(stats.countA, stats.countB)) * 100);

        results._comparison.pairwiseMatches[`${questionIds[i]}|${questionIds[j]}`] = {
          matchCount: stats.matchCount,
          q1Count: stats.countA,
          q2Count: stats.countB,
          matchPercentage: percentage
        };
      }
    }
  }

  return results;
}

/**
 * Get preview data for questions
 * @param {string[]} questionIds
 * @returns {Array}
 */
function getQuestionPreviewData(questionIds) {
  const mgr = getDataManager();
  
  return questionIds.map(qId => {
    try {
      const { questionText } = mgr.findQuestionColumn(qId);
      const answers = mgr.getUniqueAnswers(qId);
      return { questionId: qId, questionText, answers };
    } catch (err) {
      return { questionId: qId, error: err.message, answers: [] };
    }
  });
}

/**
 * Get group match percentage
 * @param {string[]} questionIds
 * @returns {number}
 */
function getGroupMatchPercentage(questionIds) {
  if (!questionIds || questionIds.length < 2) return 0;

  const mgr = getDataManager();
  const questionData = questionIds.map(qId => {
    try {
      return { questionId: qId, uniqueAnswers: mgr.getUniqueAnswers(qId) };
    } catch (err) {
      return { questionId: qId, uniqueAnswers: [] };
    }
  });

  const sets = questionData.map(qd => new Set(qd.uniqueAnswers));
  const commonAnswers = Array.from(sets[0]).filter(ans => sets.every(s => s.has(ans)));

  let totalPercentage = 0;
  let validQuestions = 0;

  questionData.forEach(qd => {
    if (qd.uniqueAnswers.length > 0) {
      const matchCount = qd.uniqueAnswers.filter(ans => commonAnswers.includes(ans)).length;
      totalPercentage += (matchCount / qd.uniqueAnswers.length) * 100;
      validQuestions++;
    }
  });

  return validQuestions > 0 ? Math.round(totalPercentage / validQuestions) : 0;
}

// === UTILITY FUNCTIONS ===

/**
 * Get intersection statistics between two lists
 * @param {Array} listA
 * @param {Array} listB
 * @returns {{matchCount: number, countA: number, countB: number}}
 */
function getIntersectionStats(listA, listB) {
  const setA = new Set(listA.map(String));
  const setB = new Set(listB.map(String));
  const matchCount = Array.from(setA).filter(a => setB.has(a)).length;

  return { matchCount, countA: setA.size, countB: setB.size };
}

/**
 * Validate custom sort match against question answers
 * @param {string[]} questionIds
 * @param {string} customOrderString
 * @returns {{valid: boolean, lowMatchQuestions: Array}}
 */
function validateCustomSortMatch(questionIds, customOrderString) {
  if (!customOrderString || !questionIds || questionIds.length === 0) {
    return { valid: true, lowMatchQuestions: [] };
  }

  const customValues = customOrderString.split(',').map(s => s.trim()).filter(Boolean);
  if (customValues.length === 0) return { valid: true, lowMatchQuestions: [] };

  const lowMatchQuestions = [];

  questionIds.forEach(qId => {
    try {
      const answers = getUniqueAnswers(qId);
      const stats = getIntersectionStats(answers, customValues);
      const percentage = Math.round((stats.matchCount / customValues.length) * 100);

      if (percentage < LOW_MATCH_THRESHOLD_PERCENT) {
        lowMatchQuestions.push({ id: qId, percentage, matchCount: stats.matchCount });
      }
    } catch (e) {
      Logger.log(`Error validating sort for ${qId}: ${e.message}`);
    }
  });

  return { valid: lowMatchQuestions.length === 0, lowMatchQuestions };
}

// === STORAGE OPERATIONS ===

/**
 * Save question links
 * @param {string} linksJson
 * @returns {{success: boolean}}
 */
function saveLinks(linksJson) {
  PropertiesService.getDocumentProperties().setProperty(SCRIPT_PROPERTY_KEY, linksJson);
  return { success: true };
}

/**
 * Get existing links
 * @returns {string}
 */
function getExistingLinks() {
  return PropertiesService.getDocumentProperties().getProperty(SCRIPT_PROPERTY_KEY) || '{}';
}

/**
 * Save hidden questions
 * @param {string} hiddenJson
 * @returns {{success: boolean}}
 */
function saveHiddenQuestions(hiddenJson) {
  PropertiesService.getDocumentProperties().setProperty(HIDDEN_QUESTIONS_KEY, hiddenJson);
  return { success: true };
}

/**
 * Get hidden questions
 * @returns {string}
 */
function getHiddenQuestions() {
  return PropertiesService.getDocumentProperties().getProperty(HIDDEN_QUESTIONS_KEY) || '[]';
}

/**
 * Save archived questions
 * @param {string} archivedJson
 * @returns {{success: boolean}}
 */
function saveArchivedQuestions(archivedJson) {
  PropertiesService.getDocumentProperties().setProperty(ARCHIVED_QUESTIONS_KEY, archivedJson);
  return { success: true };
}

/**
 * Get archived questions
 * @returns {string}
 */
function getArchivedQuestions() {
  return PropertiesService.getDocumentProperties().getProperty(ARCHIVED_QUESTIONS_KEY) || '[]';
}

/**
 * Save both hidden and archived questions
 * @param {string} hiddenJson
 * @param {string} archivedJson
 * @returns {{success: boolean}}
 */
function saveHiddenAndArchivedQuestions(hiddenJson, archivedJson) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(HIDDEN_QUESTIONS_KEY, hiddenJson);
  props.setProperty(ARCHIVED_QUESTIONS_KEY, archivedJson);
  return { success: true };
}

// === COLUMN OPERATIONS ===

/**
 * Hide a question column
 * @param {string} sheetName
 * @param {string} questionId
 * @returns {{success: boolean, message?: string}}
 */
function hideQuestion(sheetName, questionId) {
  const mgr = getDataManager();
  return mgr.hideColumn(questionId);
}

/**
 * Show a hidden question column
 * @param {string} sheetName
 * @param {string} questionId
 * @returns {{success: boolean, message?: string}}
 */
function unhideQuestion(sheetName, questionId) {
  const mgr = getDataManager();
  return mgr.showColumn(questionId);
}

/**
 * Delete a question column and clean up links
 * @param {string} sheetName
 * @param {string} questionId
 * @returns {{success: boolean, message: string}}
 */
function deleteQuestion(sheetName, questionId) {
  const mgr = getDataManager();
  const result = mgr.deleteColumn(questionId);
  
  if (!result.success) {
    return result;
  }

  // Clean up links
  const props = PropertiesService.getDocumentProperties();
  let allLinks;
  
  try {
    allLinks = JSON.parse(props.getProperty(SCRIPT_PROPERTY_KEY) || '{}');
  } catch (e) {
    allLinks = {};
  }

  // Remove questionId from all groups
  for (const groupName in allLinks) {
    allLinks[groupName] = allLinks[groupName].filter(id => id !== questionId);
    if (allLinks[groupName].length === 0) {
      delete allLinks[groupName];
    }
  }

  props.setProperty(SCRIPT_PROPERTY_KEY, JSON.stringify(allLinks));
  resetDataManager();  // Clear caches

  return { success: true, message: 'Question deleted and links cleaned.' };
}

// === HELPER FUNCTION FOR QUESTION PROCESSING ===

/**
 * Get questions to process from options
 * @param {Object} options - { groupName?: string, questionId?: string }
 * @returns {Array}
 * @private
 */
function getQuestionsToProcess_(options) {
  const { groupName, questionId } = options;
  const allLinks = JSON.parse(getExistingLinks());
  const groupIds = groupName ? allLinks[groupName] : (questionId ? [questionId] : []);

  if (!groupIds || groupIds.length === 0) return [];

  const allQuestions = getAllQuestionsFlat_();
  return groupIds.map(id => allQuestions.find(q => q.id === id)).filter(Boolean);
}

/**
 * Get all questions as flat array
 * @returns {Array}
 * @private
 */
function getAllQuestionsFlat_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  return ss.getSheets()
    .filter(isDataSheet)
    .flatMap(sheet => {
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0 || sheet.getLastRow() < 2) return [];

      const [headers, ids] = sheet.getRange(1, 1, 2, lastCol).getValues();
      return headers.map((text, i) => 
        text && ids[i] ? { question: text, id: String(ids[i]), sheet: sheet.getName() } : null
      ).filter(Boolean);
    });
}
