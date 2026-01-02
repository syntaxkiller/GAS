/**
 * @OnlyCurrentDoc
 * Refactored survey tool - consolidated server-side logic
 */

// === CONSTANTS ===
const [LINKS_SHEET_NAME, CONFIG_SHEET_NAME, SCRIPT_PROPERTY_KEY, HIDDEN_QUESTIONS_KEY, 
       ARCHIVED_QUESTIONS_KEY, IDS_VERIFIED_FLAG, SERVER_URL, SECRET_KEY] = [
  'Links', 'Question Configuration', 'questionLinks', 'hiddenQuestionIds',
  'archivedQuestionIds', 'survey_tool_ids_verified_v1',
  'https://cripplingalcoholism.com/api/sync.php', 'AAAAB3NzaC1yc2EAAAADAQABAAABAQCcehTkTH5iAywXFTIHgbJITg+zTkcjX6IDAotX3hZDPG0m0WTEQNMoGhKmk/Es9CgTtpRwqWYuSIYAEb76y7jvN3ACPaytOuTXyNzv6cEo9NctZUx9tOo4N7XQLnNYWnr8nwCKg98y3iAkMiRN5a4+G+oNGOij6aVZVOF1yjctivMrjmfx5Sj41fZqRPsZi/3nDBO6fEomtU3aAQSHhag6L5xlyPB4Z59Fo5YEfdjgxLLixqVH09R5pggiKG3H4oatMGvOXiqJMvRnKy9Gp7W6+YX/NsnjTzTnLRWNPgbLTcuROQaAGvuTbShTRdqkldDWBxN1liot3Nzj7RtvT7CD'
];

const QUESTION_TYPES = {
  YES_NO_NA: 'Yes No N/A', NUMERICAL: 'Numerical input', NUMERICAL_BINNED: 'Numerical (Pre-binned)',
  MULTIPLE_CHOICE_NUM: 'Multiple Choice (Numerical)', MULTIPLE_CHOICE: 'Multiple Choice',
  TEXT: 'Text input', HEIGHT: 'Height', SCALE_1_5: '1-5 Scale',
  COUNTRY: 'Country', STATE: 'State/Province', ARRAY: 'Array (comma-separated)'
};

const CHART_TYPES = {
  TABLE: 'table', BAR: 'bar', COLUMN: 'column', PIE: 'pie', DOUGHNUT: 'doughnut',
  LINE: 'line', AREA: 'area', SCATTER: 'scatter', BUBBLE: 'bubble', RADAR: 'radar',
  POLAR: 'polarArea', HISTOGRAM: 'histogram', BOX_PLOT: 'boxPlot', COUNTRY_MAP: 'country_map',
  STATE_MAP: 'state_map', AGE_PYRAMID: 'age_pyramid', CHORD: 'chord', WORD_CLOUD: 'word_cloud',
  PICTOGRAM: 'pictogram'
};

const SORT_TYPES = {
  NONE: 'none', ALPHABETICAL: 'alphabetical', FREQUENCY_DESC: 'frequency_desc',
  FREQUENCY_ASC: 'frequency_asc', HEIGHT: 'height', CUSTOM: 'custom'
}; 

const CHART_COMPATIBILITY = {
  [QUESTION_TYPES.YES_NO_NA]: ['TABLE', 'BAR', 'COLUMN', 'PIE', 'DOUGHNUT', 'POLAR'],
  [QUESTION_TYPES.NUMERICAL]: ['TABLE', 'HISTOGRAM', 'LINE', 'SCATTER', 'BOX_PLOT'],
  [QUESTION_TYPES.NUMERICAL_BINNED]: ['TABLE', 'BAR', 'COLUMN', 'LINE', 'AREA'],
  [QUESTION_TYPES.MULTIPLE_CHOICE_NUM]: ['TABLE', 'BAR', 'COLUMN', 'PIE', 'DOUGHNUT', 'RADAR'],
  [QUESTION_TYPES.MULTIPLE_CHOICE]: ['TABLE', 'BAR', 'COLUMN', 'PIE', 'DOUGHNUT', 'PICTOGRAM'],
  [QUESTION_TYPES.TEXT]: ['TABLE', 'WORD_CLOUD'],
  [QUESTION_TYPES.HEIGHT]: ['TABLE', 'HISTOGRAM', 'AGE_PYRAMID'],
  [QUESTION_TYPES.SCALE_1_5]: ['TABLE', 'BAR', 'COLUMN', 'RADAR', 'LINE'],
  [QUESTION_TYPES.COUNTRY]: ['TABLE', 'COUNTRY_MAP', 'BAR', 'PIE'],
  [QUESTION_TYPES.STATE]: ['TABLE', 'STATE_MAP', 'BAR', 'PIE'],
  [QUESTION_TYPES.ARRAY]: ['TABLE', 'BAR', 'COLUMN', 'RADAR', 'CHORD']
};

// === UTILITY FUNCTIONS ===
const getCompatibleCharts = (type) => (CHART_COMPATIBILITY[type] || ['TABLE']).map(k => CHART_TYPES[k] || k);
const getQuestionTypes = () => Object.values(QUESTION_TYPES);
const getChartTypes = () => Object.values(CHART_TYPES);
const getSortTypes = () => Object.values(SORT_TYPES);
const isChartCompatible = (qType, cType) => getCompatibleCharts(qType).includes(cType);
function getIntersectionStats(listA, listB) {
  const setA = new Set(listA.map(String));
  const setB = new Set(listB.map(String));
  
  // Calculate intersection
  const matchCount = Array.from(setA).filter(a => setB.has(a)).length;
  
  return {
    matchCount: matchCount,
    countA: setA.size,
    countB: setB.size
  };
}

// === SHEET SETUP ===
function ensureConfigSheetExists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
    const headers = ['question_id', 'question_text', 'question_type', 'chart_type', 'mobile_chart_type',
                     'has_other', 'sort_type_desktop', 'sort_type_mobile', 'custom_order', 'custom_order_mobile'];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#4A86E8').setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
    sheet.setFrozenRows(1);
    
    [120, 300, 180, 150, 150, 100, 150, 150, 250, 250].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    setupColumnValidation(sheet);
  }
  return sheet;
}

function setupColumnValidation(sheet) {
  const maxRows = 1000;
  const rules = [
    [3, SpreadsheetApp.newDataValidation().requireValueInList(getQuestionTypes()).setAllowInvalid(false).build()],
    [6, SpreadsheetApp.newDataValidation().requireCheckbox().build()],
    [7, SpreadsheetApp.newDataValidation().requireValueInList(getSortTypes()).setAllowInvalid(false).build()],
    [8, SpreadsheetApp.newDataValidation().requireValueInList(getSortTypes()).setAllowInvalid(false).build()]
  ];
  rules.forEach(([col, rule]) => sheet.getRange(2, col, maxRows, 1).setDataValidation(rule));
}

// === DATA LOADING ===
function getQuestionConfigData() {
  return {
    questions: getAllQuestionList(),
    links: JSON.parse(getExistingLinks()),
    config: loadConfigFromSheet()
  };
}

function getAllQuestionList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archived = JSON.parse(PropertiesService.getDocumentProperties().getProperty(ARCHIVED_QUESTIONS_KEY) || '[]');
  
  return ss.getSheets()
    .filter(s => s.getName() !== LINKS_SHEET_NAME && s.getName() !== CONFIG_SHEET_NAME && !s.getName().startsWith('_'))
    .flatMap(sheet => {
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0 || sheet.getLastRow() < 2) return [];
      
      const [texts, ids] = sheet.getRange(1, 1, 2, lastCol).getValues();
      return ids.map((id, i) => {
        const qId = String(id).trim();
        return qId && !archived.includes(qId) ? 
          { questionId: qId, questionText: texts[i] || '', year: sheet.getName() } : null;
      }).filter(Boolean);
    });
}

function loadConfigFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return {};
  
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues().reduce((acc, row) => {
    const qId = String(row[0]).trim();
    if (qId) acc[qId] = {
      questionText: row[1] || '',
      questionType: row[2],
      chartType: row[3],
      mobileChartType: row[4] || '',
      hasOther: row[5] === true,
      sortType: row[6] || 'none',
      sortTypeMobile: row[7] || '',
      customOrder: row[8] || '',
      customOrderMobile: row[9] || ''
    };
    return acc;
  }, {});
}

function getAnswerOptions(questionId) {
  const mgr = new SpreadsheetDataManager();
  return mgr.getUniqueAnswers(questionId, { withCounts: true, sorted: true });
}

function compareQuestionAnswers(questionIds) {
  const mgr = new SpreadsheetDataManager();
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
      totalQuestions: questionIds.length, commonAnswers: common, commonCount: common.length,
      pairwiseMatches: {}
    };

    for (let i = 0; i < questionIds.length; i++) {
      for (let j = i + 1; j < questionIds.length; j++) {
        const q1Answers = results[questionIds[i]].answers;
        const q2Answers = results[questionIds[j]].answers;
        
        // Use helper
        const stats = getIntersectionStats(q1Answers, q2Answers);
        
        // For pairwise comparison, we still use Math.max to measure overall similarity
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

      // Use Custom Values count as denominator to avoid skew from "Other" answers
      const percentage = Math.round((stats.matchCount / customValues.length) * 100);

      if (percentage < 20) {
        lowMatchQuestions.push({
          id: qId,
          percentage: percentage,
          matchCount: stats.matchCount
        });
      }
    } catch (e) {
      Logger.log(`Error validating sort for ${qId}: ${e.message}`);
    }
  });

  return {
    valid: lowMatchQuestions.length === 0,
    lowMatchQuestions: lowMatchQuestions
  };
}

function getGroupMatchPercentage(questionIds) {
  if (!questionIds || questionIds.length < 2) return 0;
  
  const mgr = new SpreadsheetDataManager();
  const questionData = [];
  
  // Collect unique answers for each question
  questionIds.forEach(qId => {
    try {
      const uniqueAnswers = mgr.getUniqueAnswers(qId);
      questionData.push({ questionId: qId, uniqueAnswers });
    } catch (err) {
      questionData.push({ questionId: qId, uniqueAnswers: [] });
    }
  });
  
  // Find common answers across all questions
  const sets = questionData.map(qd => new Set(qd.uniqueAnswers));
  const commonAnswers = Array.from(sets[0]).filter(ans => sets.every(s => s.has(ans)));
  
  // Calculate average match percentage
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

function getQuestionPreviewData(questionIds) {
  const mgr = new SpreadsheetDataManager();
  const results = [];
  
  questionIds.forEach(qId => {
    try {
      const { questionText } = mgr.findQuestionColumn(qId);
      const answers = mgr.getUniqueAnswers(qId);
      results.push({ questionId: qId, questionText, answers });
    } catch (err) {
      results.push({ questionId: qId, error: err.message, answers: [] });
    }
  });
  
  return results;
}

function getUniqueAnswers(qId) {
  const mgr = new SpreadsheetDataManager();
  return mgr.getUniqueAnswers(qId);
}

function getNumericalStats(qId) {
  const [year] = qId.split('_');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(year);
  if (!sheet) throw new Error('Sheet not found: ' + year);
  
  const ids = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = ids.findIndex(id => String(id).trim() === qId);
  if (colIndex === -1) throw new Error('Question not found: ' + qId);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return { values: [], stats: {} };
  
  const values = sheet.getRange(3, colIndex + 1, lastRow - 2, 1).getValues()
    .map(([v]) => parseFloat(v)).filter(v => !isNaN(v)).sort((a, b) => a - b);
  
  if (values.length === 0) return { values: [], stats: {} };
  
  const sum = values.reduce((a, b) => a + b, 0);
  const counts = values.reduce((acc, v) => ({ ...acc, [v]: (acc[v] || 0) + 1 }), {});
  
  return {
    values: Object.entries(counts).map(([value, count]) => ({ value: parseFloat(value), count }))
      .sort((a, b) => a.value - b.value),
    stats: {
      min: values[0], max: values[values.length - 1],
      mean: sum / values.length, median: values[Math.floor(values.length / 2)],
      count: values.length
    }
  };
}

function getMultipleAnswerOptions(questionIds) {
  const result = {};
  questionIds.forEach(qId => {
    try {
      result[qId] = { answers: getAnswerOptions(qId) };
    } catch (err) {
      result[qId] = { error: err.message, answers: [] };
    }
  });
  return result;
}

//Question Configuration Columns: A: question_id B: question_text C: question_type D: chart_type E: mobile_chart_type F: has_other g: sort_type_desktop H: sort_type_mobile I: custom_order J: custom_order_mobile

function saveGroupConfigsToSheet(groupConfigs) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    throw new Error('Question Configuration sheet not found');
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lastRow = sheet.getLastRow();
  
  // Load existing question IDs
  const existingIds = {};
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(([id], i) => {
      if (id) existingIds[String(id).trim()] = i + 2;
    });
  }
  
  const updates = [];
  const newRows = [];
  
  // Process each group
  groupConfigs.forEach(groupConfig => {
    groupConfig.questionIds.forEach(qId => {
      // Get question text
      const [year] = qId.split('_');
      let questionText = '';
      try {
        const yearSheet = ss.getSheetByName(year);
        if (yearSheet) {
          const [texts, ids] = yearSheet.getRange(1, 1, 2, yearSheet.getLastColumn()).getValues();
          const colIndex = ids.findIndex(id => String(id).trim() === qId);
          if (colIndex >= 0) questionText = texts[colIndex] || '';
        }
      } catch (e) {
        Logger.log('Could not get question text for ' + qId + ': ' + e.message);
      }
      
      let customOrder = '';
      if (Array.isArray(groupConfig.customOrder)) {
        customOrder = groupConfig.customOrder.join(',');
      } else {
        customOrder = groupConfig.customOrder || '';
      }
      
      const row = [
        qId,
        questionText,
        groupConfig.questionType || '',
        groupConfig.chartType || '',
        groupConfig.mobileChartType || '',
        groupConfig.hasOther || false,
        groupConfig.sortType || 'none',
        '', // sortTypeMobile
        customOrder,
        '' // customOrderMobile
      ];
      
      if (existingIds[qId]) {
        updates.push({ rowNumber: existingIds[qId], data: row });
      } else {
        newRows.push(row);
      }
    });
  });
  
  // Write updates
  updates.forEach(({ rowNumber, data }) => {
    const range = sheet.getRange(rowNumber, 1, 1, 10);
    range.setValues([data]);
  });
  
  // Write new rows in one batch
  if (newRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    const range = sheet.getRange(startRow, 1, newRows.length, 10);
    range.setValues(newRows);
  }
  
  Logger.log('Saved ' + (updates.length + newRows.length) + ' configs (' + updates.length + ' updates, ' + newRows.length + ' new)');
  
  return { success: true, count: updates.length + newRows.length, updates: updates.length, new: newRows.length };
}

// === SAVE & UPLOAD ===
function saveConfigToSheet(questionId, config) {
  const sheet = ensureConfigSheetExists();
  const lastRow = sheet.getLastRow();
  
  // Get question text using cached manager
  const mgr = new SpreadsheetDataManager();
  let questionText = '';
  try {
    const result = mgr.findQuestionColumn(questionId);
    questionText = result.questionText || '';
  } catch (e) {
    Logger.log('Could not get question text: ' + e.message);
  }
  
  // Find existing row or add new
  let rowIndex = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    rowIndex = ids.findIndex(([id]) => String(id).trim() === questionId);
  }
  
  const row = [
    questionId,
    questionText,
    config.questionType || '',
    config.chartType || '',
    config.mobileChartType || '',
    config.hasOther || false,
    config.sortType || 'none',
    config.sortTypeMobile || '',
    config.customOrder || '',
    config.customOrderMobile || ''
  ];
  
  if (rowIndex >= 0) {
    sheet.getRange(rowIndex + 2, 1, 1, 10).setValues([row]);
  } else {
    sheet.getRange(lastRow + 1, 1, 1, 10).setValues([row]);
  }
  
  return { success: true };
}

function saveConfigToSheetBulk(configObj) {
  const sheet = ensureConfigSheetExists();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  const rows = Object.entries(configObj).map(([qId, cfg]) => {
    // Get question text
    const [year] = qId.split('_');
    let questionText = '';
    try {
      const yearSheet = ss.getSheetByName(year);
      if (yearSheet) {
        const [texts, ids] = yearSheet.getRange(1, 1, 2, yearSheet.getLastColumn()).getValues();
        const colIndex = ids.findIndex(id => String(id).trim() === qId);
        if (colIndex >= 0) questionText = texts[colIndex] || '';
      }
    } catch (e) {
      Logger.log('Could not get question text for ' + qId);
    }
    
    return [
      qId,
      questionText,
      cfg.questionType || '',
      cfg.chartType || '',
      cfg.mobileChartType || '',
      cfg.hasOther || false,
      cfg.sortType || 'none',
      cfg.sortTypeMobile || '',
      cfg.customOrder || '',
      cfg.customOrderMobile || ''
    ];
  });
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 10).setValues(rows);
  }
  
  return { success: true, rowCount: rows.length };
}

function saveMultipleConfigsToSheet(configsArray) {
  const sheet = ensureConfigSheetExists();
  const lastRow = sheet.getLastRow();
  
  // Load existing question IDs
  const existingIds = {};
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(([id], i) => {
      if (id) existingIds[String(id).trim()] = i + 2; // row number
    });
  }
  
  // Update or add each config
  configsArray.forEach(({ questionId, config }) => {
    const row = [
      questionId,
      config.questionType || '',
      config.chartType || '',
      config.mobileChartType || '',
      config.hasOther || false,
      config.sortType || 'none',
      config.sortTypeMobile || '',
      config.customOrder || '',
      config.binDefinition || ''
    ];
    
    if (existingIds[questionId]) {
      sheet.getRange(existingIds[questionId], 1, 1, 9).setValues([row]);
    } else {
      const newRow = sheet.getLastRow() + 1;
      sheet.getRange(newRow, 1, 1, 9).setValues([row]);
      existingIds[questionId] = newRow;
    }
  });
  
  return { success: true, count: configsArray.length };
}

function pushQuestionConfigOnly() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) throw new Error('Question Configuration sheet not found.');
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No configuration data to upload.');
  
  const [headers] = sheet.getRange(1, 1, 1, 10).getValues();
  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  
  const payload = {
    dataType: 'questionConfig',
    payload: data.map(row => headers.reduce((obj, h, i) => ({ ...obj, [h]: row[i] }), {}))
  };
  
  _sendToServer(payload);
  
  // NEW: Trigger the Node.js script to actually import the data into MySQL
  triggerCombine();
  
  return { success: true, count: data.length };
}

function saveAndUploadConfig(configObj) {
  const saveResult = saveConfigToSheetBulk(configObj);
  const uploadResult = pushQuestionConfigOnly();
  return { success: true, saved: saveResult.rowCount, uploaded: uploadResult.count };
}

function showQuestionConfigDialog() {
  ensureConfigSheetExists();
  HtmlService.createTemplateFromFile('QuestionConfig').evaluate()
    .setWidth(1600).setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('QuestionConfig').evaluate().setWidth(1600).setHeight(900),
    'Question Configuration'
  );
}

// === MENU & UI ===
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Survey Tools')
    .addItem('🔗 Link Questions', 'showLinkingDialog')
    .addItem('📝 Chart Config', 'showQuestionConfigDialog')
    .addItem('⚡ Manage Duplicates', 'showDuplicateManager') // <--- Added this line
    .addSeparator()
    .addItem('Push All Data to Website', 'pushAllDataToServer')
    .addItem('Push Question Config Only', 'pushQuestionConfigOnly')
    .addSeparator()
    .addItem('Debug: View Stored Data', 'viewStoredData')
    .addToUi();
}

function showLinkingDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('Sidebar').evaluate().setWidth(1600).setHeight(900),
    'Survey Tools'
  );
}

// === DATA SERVICE ===
function ensureQuestionIdsExist() {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty(IDS_VERIFIED_FLAG)) {
    return { success: true, message: 'IDs already verified.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets()
    .filter(s => s.getName() !== LINKS_SHEET_NAME && !s.getName().startsWith('_'))
    .forEach(sheet => {
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0 || sheet.getLastRow() === 0) return;

      const row2 = sheet.getMaxRows() > 1 ? sheet.getRange(2, 1, 1, lastCol).getValues()[0] : [];
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

function getSidebarInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsData = {};
  
  ss.getSheets()
    .filter(s => s.getName() !== LINKS_SHEET_NAME && !s.getName().startsWith('_'))
    .forEach(sheet => {
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0 || sheet.getLastRow() < 2) return;
      
      const [questions, ids] = sheet.getRange(1, 1, 2, lastCol).getValues();
      let prevText = null;
      
      sheetsData[sheet.getName()] = questions.map((text, i) => {
        const result = text && ids[i] ? {
          question: text, id: String(ids[i]), sheet: sheet.getName(), previousQuestion: prevText
        } : null;
        if (text) prevText = text;
        else prevText = null;
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

function getActiveSheetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return {
    sheetName: sheet.getName(),
    headers: sheet.getLastRow() === 0 ? [] : sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
  };
}

function getAllQuestionsFlat_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()
    .filter(s => s.getName() !== LINKS_SHEET_NAME && !s.getName().startsWith('_'))
    .flatMap(sheet => {
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0 || sheet.getLastRow() < 2) return [];
      
      const [headers, ids] = sheet.getRange(1, 1, 2, lastCol).getValues();
      return headers.map((text, i) => text && ids[i] ? 
        { question: text, id: String(ids[i]), sheet: sheet.getName() } : null
      ).filter(Boolean);
    });
}

function getQuestionsToProcess_(options) {
  const { groupName, questionId } = options;
  const allLinks = JSON.parse(getExistingLinks());
  const groupIds = groupName ? allLinks[groupName] : (questionId ? [questionId] : []);
  
  if (!groupIds || groupIds.length === 0) return [];
  
  const allQuestions = getAllQuestionsFlat_();
  return groupIds.map(id => allQuestions.find(q => q.id === id)).filter(Boolean);
}


function getFacetData(options) {
  const { splitByComma } = options;
  const questions = getQuestionsToProcess_(options);
  
  Logger.log('getFacetData - Options: ' + JSON.stringify(options));
  Logger.log('getFacetData - Questions found: ' + questions.length);
  
  if (!questions || questions.length === 0) {
    return { error: 'No questions found' };
  }

  // Aggregate all values across all questions in the group
  const valueCounts = {};
  let totalValuesProcessed = 0;

  _processDataInColumns_(questions, true, null, (q, val) => {
    const v = String(val || '').trim();
    if (!v) return;
    
    totalValuesProcessed++;
    
    // Handle split by comma if enabled
    if (splitByComma) {
      // Split by comma and process each part
      const parts = v.split(',').map(part => part.trim()).filter(part => part);
      parts.forEach(part => {
        valueCounts[part] = (valueCounts[part] || 0) + 1;
      });
    } else {
      // Process as single value
      valueCounts[v] = (valueCounts[v] || 0) + 1;
    }
  });

  Logger.log('getFacetData - Total values processed: ' + totalValuesProcessed);
  Logger.log('getFacetData - Unique values found: ' + Object.keys(valueCounts).length);

  // Convert to the expected format
  const uniqueValues = Object.keys(valueCounts).sort();
  const valueCountsArray = Object.entries(valueCounts)
    .map(([text, count]) => ({ text, count }))
    .sort((a, b) => b.count - a.count);

  return {
    uniqueValues: uniqueValues,
    valueCounts: valueCountsArray
  };
}

// Code.gs
function getCustomRemovalWords() {
  try {
    const userProps = PropertiesService.getUserProperties();
    const customWords = userProps.getProperty('wordcloud_custom_removal_words');
    return customWords ? JSON.parse(customWords) : [];
  } catch (err) {
    console.error('Error getting custom words:', err);
    return [];
  }
}

function setCustomRemovalWords(words) {
  try {
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty('wordcloud_custom_removal_words', JSON.stringify(words));
    return true;
  } catch (err) {
    console.error('Error setting custom words:', err);
    return false;
  }
}

function processTextForWordCloud(text, options) {
  options = options || {};
  let processed = text.toLowerCase();
  
  // Remove punctuation if requested
  if (options.removePunctuation) {
    processed = processed.replace(/[^\w\s]/g, ' ');
  }
  
  // Split into words
  let words = processed.split(/\s+/).filter(w => w.length > 0);
  
  // Remove stop words if requested
  if (options.removeStopWords) {
    const stopWords = new Set([
      'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'for', 'from', 'has', 'he',
      'in', 'is', 'it', 'its', 'of', 'on', 'that', 'the', 'to', 'was', 'will', 'with'
    ]);
    words = words.filter(w => !stopWords.has(w));
  }
  
  // Remove custom words if requested
  if (options.removeCustom && options.customWords) {
    const customSet = new Set(options.customWords.map(w => w.toLowerCase()));
    words = words.filter(w => !customSet.has(w));
  }
  
  // Count word frequencies
  const frequencies = {};
  words.forEach(word => {
    frequencies[word] = (frequencies[word] || 0) + 1;
  });
  
  return frequencies;
}

function replaceTextValuesWithOptions(questionId, replacements, splitByComma) {
  const mgr = new SpreadsheetDataManager();
  
  // Build replacement map
  const replaceMap = {};
  replacements.forEach(r => {
    replaceMap[r.oldValue] = r.newValue;
  });
  
  const result = mgr.replaceValues(questionId, replaceMap, { splitByComma });
  return { success: true, replacements: result.replacements };
}

function replaceTextValuesRateOptimized(questionIds, replacements, splitByComma) {
  const mgr = new SpreadsheetDataManager();
  
  // Build replacement map
  const replaceMap = {};
  replacements.forEach(r => {
    replaceMap[r.oldValue] = r.newValue;
  });
  
  const result = mgr.batchReplaceValues(questionIds, replaceMap, { 
    splitByComma,
    batchSize: 5,
    delayBetweenBatches: 500
  });
  
  return {
    success: true,
    replacements: result.totalReplacements,
    questionsProcessed: result.questionsProcessed
  };
}



function formatAsPlainText(options) {
  const { groupName, questionId } = options;
  
  const questions = getQuestionsToProcess_({ groupName, questionId });
  if (!questions || questions.length === 0) {
    return { success: false, message: 'No questions found.' };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalFormatted = 0;
  
  // Group questions by sheet for batch operations
  const bySheet = questions.reduce((acc, q) => {
    if (!acc[q.sheet]) acc[q.sheet] = [];
    acc[q.sheet].push(q);
    return acc;
  }, {});
  
  Object.entries(bySheet).forEach(([sheetName, qs]) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    
    const ids = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    qs.forEach(q => {
      const colIndex = ids.findIndex(id => String(id).trim() === q.id);
      if (colIndex === -1) return;
      
      // Format the entire column (from row 3 to last row) as plain text
      const range = sheet.getRange(3, colIndex + 1, lastRow - 2, 1);
      range.setNumberFormat('@');
      totalFormatted++;
    });
  });
  
  return {
    success: true,
    message: `Formatted ${totalFormatted} column(s) as plain text.`
  };
}

function _processDataInColumns_(questions, readOnly, cellCallback, perQuestionCallback) {
  const bySheet = questions.reduce((acc, q) => {
    if (!acc[q.sheet]) acc[q.sheet] = [];
    acc[q.sheet].push(q);
    return acc;
  }, {});

  Object.entries(bySheet).forEach(([name, qs]) => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 3 || lastCol === 0) return;

    const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const ids = allData[1];
    const numRows = lastRow - 2;
    
    // Collect all pending writes for this sheet (only used when not readOnly)
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
    
    // Batch write all pending changes for this sheet to avoid rate limits
    if (!readOnly && pendingWrites.length > 0) {
      // Write in chunks of 5 to avoid HTTP 429 errors
      const CHUNK_SIZE = 5;
      for (let i = 0; i < pendingWrites.length; i += CHUNK_SIZE) {
        const chunk = pendingWrites.slice(i, i + CHUNK_SIZE);
        chunk.forEach(write => {
          write.range.setValues(write.values);
        });
        SpreadsheetApp.flush();
        
        // Add small delay between chunks to respect rate limits
        if (i + CHUNK_SIZE < pendingWrites.length) {
          Utilities.sleep(100);
        }
      }
    }
  });
}

/**
 * Gets preview data for a group or single question.
 * Returns an object where keys are year/sheet names and values are arrays of unique answers.
 * 
 * @param {Object} options - { groupName: string|null, questionId: string|null }
 * @returns {Object} - { "2023": ["answer1", "answer2"], "2024": [...], ... }
 */
function getFullPreviewData(options) {
  const questions = getQuestionsToProcess_(options);
  
  if (!questions || questions.length === 0) {
    return {}; // Return empty object if no questions found
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const previewData = {};

  // Process each question
  questions.forEach(q => {
    const sheet = ss.getSheetByName(q.sheet);
    if (!sheet) return; // Skip if sheet not found
    
    const year = q.sheet; // Use sheet name as the year key
    
    // Initialize the array for this year if it doesn't exist
    if (!previewData[year]) {
      previewData[year] = [];
    }
    
    // Find the column for this question
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    if (lastCol === 0 || lastRow < 2) return; // Skip if sheet is empty
    
    const ids = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    const colIndex = ids.findIndex(id => String(id).trim() === q.id);
    
    if (colIndex === -1) return; // Skip if question not found
    
    // Get unique answers for this question
    if (lastRow >= 3) {
      const answers = sheet.getRange(3, colIndex + 1, lastRow - 2, 1).getValues();
      const uniqueAnswers = new Set();
      
      answers.forEach(([val]) => {
        const v = String(val || '').trim();
        if (v) uniqueAnswers.add(v);
      });
      
      // Add unique answers to the preview data for this year
      // Only add if not already present (in case multiple questions from same year)
      Array.from(uniqueAnswers).forEach(answer => {
        if (!previewData[year].includes(answer)) {
          previewData[year].push(answer);
        }
      });
    }
  });
  
  // Sort answers alphabetically within each year
  Object.keys(previewData).forEach(year => {
    previewData[year].sort();
  });
  
  return previewData;
}

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

  _processDataInColumns_(questions, false, (val) => {
    const str = String(val).trim();
    
    if (splitByComma) {
      // Handle comma-separated values
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
      // Handle whole cell values
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

// === SHEET ACTIONS ===
function hideQuestion(sheetName, qId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { success: false, message: 'Sheet not found' };
  
  const ids = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = ids.findIndex(id => String(id).trim() === qId);
  if (colIndex === -1) return { success: false, message: 'Question not found' };
  
  sheet.hideColumns(colIndex + 1);
  return { success: true };
}

function unhideQuestion(sheetName, qId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { success: false, message: 'Sheet not found' };
  
  const ids = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = ids.findIndex(id => String(id).trim() === qId);
  if (colIndex === -1) return { success: false, message: 'Question not found' };
  
  sheet.showColumns(colIndex + 1);
  return { success: true };
}

function deleteQuestion(sheetName, questionId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { success: false, message: 'Sheet not found: ' + sheetName };
  }
  
  // Find the column
  const ids = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = ids.findIndex(id => String(id).trim() === questionId);
  if (colIndex === -1) {
    return { success: false, message: 'Question not found: ' + questionId };
  }
  
  // 1. Delete the column from the sheet
  sheet.deleteColumn(colIndex + 1);
  
  // 2. Clean up links
  const props = PropertiesService.getDocumentProperties();
  const linksJson = props.getProperty(SCRIPT_PROPERTY_KEY) || '{}';
  let allLinks;
  
  try {
    allLinks = JSON.parse(linksJson);
  } catch (e) {
    allLinks = {};
  }
  
  // Remove this questionId from all groups
  for (const groupName in allLinks) {
    allLinks[groupName] = allLinks[groupName].filter(id => id !== questionId);
    // If the group is now empty, delete the group itself
    if (allLinks[groupName].length === 0) {
      delete allLinks[groupName];
    }
  }
  
  // Save the cleaned links back to properties
  props.setProperty(SCRIPT_PROPERTY_KEY, JSON.stringify(allLinks));
  
  return { success: true, message: 'Question deleted and links cleaned.' };
}

function getHiddenQuestions() {
  return PropertiesService.getDocumentProperties().getProperty(HIDDEN_QUESTIONS_KEY) || '[]';
}

function getArchivedQuestions() {
  return PropertiesService.getDocumentProperties().getProperty(ARCHIVED_QUESTIONS_KEY) || '[]';
}

// === STORAGE ===
function saveLinks(linksJson) {
  PropertiesService.getDocumentProperties().setProperty(SCRIPT_PROPERTY_KEY, linksJson);
  return { success: true };
}

function getExistingLinks() {
  return PropertiesService.getDocumentProperties().getProperty(SCRIPT_PROPERTY_KEY) || '{}';
}

function saveHiddenQuestions(hiddenJson) {
  PropertiesService.getDocumentProperties().setProperty(HIDDEN_QUESTIONS_KEY, hiddenJson);
  return { success: true };
}

function saveArchivedQuestions(archivedJson) {
  PropertiesService.getDocumentProperties().setProperty(ARCHIVED_QUESTIONS_KEY, archivedJson);
  return { success: true };
}

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
  
  _sendToServer(payload);
  return { success: true };
}

function pushAllSurveyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let count = 0;
  
  ss.getSheets()
    .filter(s => s.getName() !== LINKS_SHEET_NAME && !s.getName().startsWith('_'))
    .forEach(sheet => {
      const lastCol = sheet.getLastColumn();
      const lastRow = sheet.getLastRow();
      if (lastCol === 0 || lastRow < 3) return;
      
      const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      _uploadSchema(sheet.getName(), data[1], data[0]);
      _uploadCSV(sheet.getName(), data[1], data.slice(2));
      count++;
    });
  
  return { success: true, sheetsProcessed: count };
}

function triggerCombine() {
  _sendToServer({ dataType: 'combine' });
}

function viewStoredData() {
  const props = PropertiesService.getDocumentProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let debug = ss.getSheetByName('Stored_Data_Debug');
  
  if (!debug) debug = ss.insertSheet('Stored_Data_Debug');
  else debug.clear();

  [['Links JSON:', props.getProperty(SCRIPT_PROPERTY_KEY) || '{}'],
   ['Hidden JSON:', props.getProperty(HIDDEN_QUESTIONS_KEY) || '[]'],
   ['Archived JSON:', props.getProperty(ARCHIVED_QUESTIONS_KEY) || '[]']]
    .forEach(([label, data], i) => {
      debug.getRange(i * 2 + 1, 1).setValue(label);
      debug.getRange(i * 2 + 2, 1).setValue(data);
    });

  SpreadsheetApp.getUi().alert('Stored data written to "Stored_Data_Debug" sheet.');
}

function clearAllStoredData() {
  const props = PropertiesService.getDocumentProperties();
  [SCRIPT_PROPERTY_KEY, HIDDEN_QUESTIONS_KEY, ARCHIVED_QUESTIONS_KEY, IDS_VERIFIED_FLAG]
    .forEach(key => props.deleteProperty(key));
  SpreadsheetApp.getUi().alert('All stored data cleared.');
}

// === HELPER FUNCTIONS ===
function _processDataInColumns_(questions, readOnly, cellCallback, perQuestionCallback) {
  const bySheet = questions.reduce((acc, q) => {
    if (!acc[q.sheet]) acc[q.sheet] = [];
    acc[q.sheet].push(q);
    return acc;
  }, {});

  Object.entries(bySheet).forEach(([name, qs]) => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 3 || lastCol === 0) return;

    const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const ids = allData[1];
    const numRows = lastRow - 2;

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
        sheet.getRange(3, colIndex + 1, numRows, 1).setValues(colValues);
      }
    });
  });
}

// === SERVER SYNC ===
function pushAllDataToServer() {
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui.alert('Confirm Upload', 'Upload all data to server?', ui.ButtonSet.YES_NO) !== ui.Button.YES) {
      return { cancelled: true };
    }
    
    ui.alert('Upload Started', 'Uploading...', ui.ButtonSet.OK);
    
    const storage = pushPersistentStorage();
    const surveys = pushAllSurveyData();
    let config = null;
    
    try {
      config = pushQuestionConfigOnly();
    } catch (e) {
      Logger.log('No config: ' + e.message);
      config = { success: false, message: 'No configuration' };
    }
    
    triggerCombine();
    
    ui.alert('Success', `Upload complete!\n\nStorage: ✓\nSurveys: ${surveys.sheetsProcessed || 0} years\nConfig: ${config.success ? config.count + ' questions' : 'None'}`, ui.ButtonSet.OK);
    
    return { success: true, storage, surveys, config };
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error: ' + err.message);
    throw err;
  }
}

function _uploadSchema(sheetName, qIds, qTexts) {
  const archived = JSON.parse(PropertiesService.getDocumentProperties().getProperty(ARCHIVED_QUESTIONS_KEY) || '[]');
  const schema = qIds.reduce((acc, id, i) => {
    const qId = String(id).trim();
    if (qId && !archived.includes(qId)) acc[qId] = qTexts[i] || '';
    return acc;
  }, {});
  
  _sendToServer({ dataType: 'surveySchema', sheetName, payload: schema });
}

function _uploadCSV(sheetName, qIds, rows) {
  const archived = JSON.parse(PropertiesService.getDocumentProperties().getProperty(ARCHIVED_QUESTIONS_KEY) || '[]');
  const activeIds = [];
  const activeIndices = [];
  
  qIds.forEach((id, i) => {
    const qId = String(id).trim();
    if (qId && !archived.includes(qId)) {
      activeIds.push(qId);
      activeIndices.push(i);
    }
  });
  
  let csv = ['respondent_id', ...activeIds].join(',') + '\n';
  let count = 0;
  
  rows.forEach(row => {
    if (!row.some(c => c !== null && c !== undefined && String(c).trim() !== '')) return;
    
    const values = [sheetName + '_r' + (count + 1)];
    activeIndices.forEach(i => values.push(_escapeCsv(row[i])));
    csv += values.join(',') + '\n';
    count++;
  });
  
  _sendToServer({ dataType: 'surveyData', sheetName, payload: csv });
}

function _sendToServer(payload) {
  const response = UrlFetchApp.fetch(SERVER_URL, {
    method: 'post', contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + SECRET_KEY },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`Server error (${response.getResponseCode()}): ${response.getContentText()}`);
  }
}

function _escapeCsv(cell) {
  if (cell === null || cell === undefined) return '';
  let val = String(cell);
  if (val.includes(',') || val.includes('"') || val.includes('\n')) {
    val = '"' + val.replace(/"/g, '""') + '"';
  }
  return val;
}

function _getPersistentStorageData() {
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

function previewUploadData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 3) {
    SpreadsheetApp.getUi().alert('Preview failed', 'Need at least 3 rows', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const archived = JSON.parse(PropertiesService.getDocumentProperties().getProperty(ARCHIVED_QUESTIONS_KEY) || '[]');
  const [texts, ids] = [data[0], data[1]];
  const activeIds = ids.filter((id, i) => id && String(id).trim() && !archived.includes(id));
  
  Logger.log(`=== PREVIEW ===\n${activeIds.length} active questions (${ids.filter(Boolean).length - activeIds.length} archived)`);
  SpreadsheetApp.getUi().alert('Preview', `${activeIds.length} active questions\n${ids.filter(Boolean).length - activeIds.length} archived`, SpreadsheetApp.getUi().ButtonSet.OK);
}