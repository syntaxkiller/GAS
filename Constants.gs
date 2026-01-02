/**
 * Constants.gs
 * Centralized configuration and constants for the survey tool
 * @OnlyCurrentDoc
 */

// === SHEET NAMES ===
const LINKS_SHEET_NAME = 'Links';
const CONFIG_SHEET_NAME = 'Question Configuration';
const INTERNAL_SHEET_PREFIX = '_';

// === PROPERTY KEYS ===
const SCRIPT_PROPERTY_KEY = 'questionLinks';
const HIDDEN_QUESTIONS_KEY = 'hiddenQuestionIds';
const ARCHIVED_QUESTIONS_KEY = 'archivedQuestionIds';
const IDS_VERIFIED_FLAG = 'survey_tool_ids_verified_v1';
const WORDCLOUD_CUSTOM_WORDS_KEY = 'wordcloud_custom_removal_words';

// === SERVER CONFIG ===
const SERVER_URL = 'https://cripplingalcoholism.com/api/sync.php';
const SECRET_KEY = 'AAAAB3NzaC1yc2EAAAADAQABAAABAQCcehTkTH5iAywXFTIHgbJITg+zTkcjX6IDAotX3hZDPG0m0WTEQNMoGhKmk/Es9CgTtpRwqWYuSIYAEb76y7jvN3ACPaytOuTXyNzv6cEo9NctZUx9tOo4N7XQLnNYWnr8nwCKg98y3iAkMiRN5a4+G+oNGOij6aVZVOF1yjctivMrjmfx5Sj41fZqRPsZi/3nDBO6fEomtU3aAQSHhag6L5xlyPB4Z59Fo5YEfdjgxLLixqVH09R5pggiKG3H4oatMGvOXiqJMvRnKy9Gp7W6+YX/NsnjTzTnLRWNPgbLTcuROQaAGvuTbShTRdqkldDWBxN1liot3Nzj7RtvT7CD';

// === RATE LIMITING ===
const RATE_LIMIT_BATCH_SIZE = 5;
const RATE_LIMIT_DELAY_MS = 100;
const BATCH_WRITE_DELAY_MS = 500;

// === CONFIG SHEET STRUCTURE ===
// The Question Configuration sheet stores chart/display settings for each question
// Column K (color_group) stores the Color Group emoji for categorizing questions by their chart config
const CONFIG_COLUMN_COUNT = 11;
const CONFIG_HEADERS = [
  'question_id', 'question_text', 'question_type', 'chart_type', 'mobile_chart_type',
  'has_other', 'sort_type_desktop', 'sort_type_mobile', 'custom_order', 'custom_order_mobile', 'color_group'
];
const CONFIG_COLUMN_WIDTHS = [120, 300, 180, 150, 150, 100, 150, 150, 250, 250, 80];

// === COLOR GROUPS ===
// Color Groups are used to categorize questions by their chart configuration.
// Questions with the same color group share similar chart settings (type, sort, etc.)
// This is different from "Linked Questions" which group related questions across survey years.
// Color Groups are stored in column K of the Question Configuration sheet.
const COLOR_GROUP_EMOJIS = ['🔴', '🟡', '🟢', '🔵', '🟣', '🟤', '⚫', '⚪', '🔶', '🟠'];

// === VALIDATION SETTINGS ===
const MAX_VALIDATION_ROWS = 1000;
const LOW_MATCH_THRESHOLD_PERCENT = 20;

// === QUESTION TYPES ===
const QUESTION_TYPES = {
  YES_NO_NA: 'Yes No N/A',
  NUMERICAL: 'Numerical input',
  NUMERICAL_BINNED: 'Numerical (Pre-binned)',
  MULTIPLE_CHOICE_NUM: 'Multiple Choice (Numerical)',
  MULTIPLE_CHOICE: 'Multiple Choice',
  TEXT: 'Text input',
  HEIGHT: 'Height',
  SCALE_1_5: '1-5 Scale',
  COUNTRY: 'Country',
  STATE: 'State/Province',
  ARRAY: 'Array (comma-separated)'
};

// === CHART TYPES ===
const CHART_TYPES = {
  TABLE: 'table',
  BAR: 'bar',
  COLUMN: 'column',
  PIE: 'pie',
  DOUGHNUT: 'doughnut',
  LINE: 'line',
  AREA: 'area',
  SCATTER: 'scatter',
  BUBBLE: 'bubble',
  RADAR: 'radar',
  POLAR: 'polarArea',
  HISTOGRAM: 'histogram',
  BOX_PLOT: 'boxPlot',
  COUNTRY_MAP: 'country_map',
  STATE_MAP: 'state_map',
  AGE_PYRAMID: 'age_pyramid',
  CHORD: 'chord',
  WORD_CLOUD: 'word_cloud',
  PICTOGRAM: 'pictogram'
};

// === SORT TYPES ===
const SORT_TYPES = {
  NONE: 'none',
  ALPHABETICAL: 'alphabetical',
  FREQUENCY_DESC: 'frequency_desc',
  FREQUENCY_ASC: 'frequency_asc',
  HEIGHT: 'height',
  CUSTOM: 'custom'
};

// === CHART COMPATIBILITY MATRIX ===
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

// === UTILITY FUNCTIONS FOR CONSTANTS ===

/**
 * Get compatible chart types for a question type
 * @param {string} questionType - The question type
 * @returns {string[]} Array of compatible chart type values
 */
function getCompatibleCharts(questionType) {
  const chartKeys = CHART_COMPATIBILITY[questionType] || ['TABLE'];
  return chartKeys.map(key => CHART_TYPES[key] || key);
}

/**
 * Get all question type values
 * @returns {string[]}
 */
function getQuestionTypes() {
  return Object.values(QUESTION_TYPES);
}

/**
 * Get all chart type values
 * @returns {string[]}
 */
function getChartTypes() {
  return Object.values(CHART_TYPES);
}

/**
 * Get all sort type values
 * @returns {string[]}
 */
function getSortTypes() {
  return Object.values(SORT_TYPES);
}

/**
 * Check if a chart type is compatible with a question type
 * @param {string} questionType
 * @param {string} chartType
 * @returns {boolean}
 */
function isChartCompatible(questionType, chartType) {
  return getCompatibleCharts(questionType).includes(chartType);
}

/**
 * Check if a sheet should be processed (not internal/config)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {boolean}
 */
function isDataSheet(sheet) {
  const name = sheet.getName();
  return name !== LINKS_SHEET_NAME && 
         name !== CONFIG_SHEET_NAME && 
         !name.startsWith(INTERNAL_SHEET_PREFIX);
}

/**
 * Get all color group emoji values
 * @returns {string[]}
 */
function getColorGroupEmojis() {
  return COLOR_GROUP_EMOJIS;
}
