/**
 * Utils.gs
 * General utility functions and HTML template helpers
 */

/**
 * Include an HTML file in a template
 * Used by HtmlService.createTemplateFromFile() with <?!= include('filename') ?>
 * @param {string} filename - Name of the HTML file to include (without .html extension)
 * @returns {string} - The content of the HTML file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get content from an HTML template file
 * Alternative to include() that returns HtmlOutput for chaining
 * @param {string} filename
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function getHtmlContent(filename) {
  return HtmlService.createHtmlOutputFromFile(filename);
}

/**
 * Deep clone an object (JSON-safe only)
 * @param {any} obj
 * @returns {any}
 */
function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

/**
 * Safely parse JSON with a default value
 * @param {string} jsonString
 * @param {any} defaultValue
 * @returns {any}
 */
function safeJsonParse(jsonString, defaultValue = null) {
  try {
    return jsonString ? JSON.parse(jsonString) : defaultValue;
  } catch (e) {
    Logger.log('JSON parse error: ' + e.message);
    return defaultValue;
  }
}

/**
 * Format a date for display
 * @param {Date} date
 * @param {string} format - 'short', 'medium', 'long', 'iso'
 * @returns {string}
 */
function formatDate(date, format = 'medium') {
  if (!date) return '';
  
  const d = date instanceof Date ? date : new Date(date);
  
  switch (format) {
    case 'short':
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yy');
    case 'medium':
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d, yyyy');
    case 'long':
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM d, yyyy h:mm a');
    case 'iso':
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
    default:
      return d.toString();
  }
}

/**
 * Truncate a string with ellipsis
 * @param {string} str
 * @param {number} maxLength
 * @returns {string}
 */
function truncateString(str, maxLength = 50) {
  if (!str || str.length <= maxLength) return str;
  return str.substring(0, maxLength - 3) + '...';
}

/**
 * Generate a simple unique ID
 * @param {string} prefix
 * @returns {string}
 */
function generateId(prefix = 'id') {
  return prefix + '_' + Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
}

/**
 * Chunk an array into smaller arrays
 * @param {Array} array
 * @param {number} size
 * @returns {Array[]}
 */
function chunkArray(array, size) {
  const chunks = [];
  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }
  return chunks;
}

/**
 * Flatten a nested array one level
 * @param {Array} array
 * @returns {Array}
 */
function flattenArray(array) {
  return array.reduce((acc, val) => acc.concat(val), []);
}

/**
 * Remove duplicates from an array
 * @param {Array} array
 * @returns {Array}
 */
function uniqueArray(array) {
  return [...new Set(array)];
}

/**
 * Sort an array of objects by a key
 * @param {Object[]} array
 * @param {string} key
 * @param {boolean} descending
 * @returns {Object[]}
 */
function sortByKey(array, key, descending = false) {
  return [...array].sort((a, b) => {
    const valA = a[key];
    const valB = b[key];
    const cmp = valA < valB ? -1 : valA > valB ? 1 : 0;
    return descending ? -cmp : cmp;
  });
}

/**
 * Group an array of objects by a key
 * @param {Object[]} array
 * @param {string} key
 * @returns {Object}
 */
function groupByKey(array, key) {
  return array.reduce((acc, item) => {
    const groupKey = item[key];
    if (!acc[groupKey]) acc[groupKey] = [];
    acc[groupKey].push(item);
    return acc;
  }, {});
}

/**
 * Create a map from an array using a key function
 * @param {Array} array
 * @param {Function} keyFn
 * @returns {Object}
 */
function arrayToMap(array, keyFn) {
  return array.reduce((acc, item) => {
    acc[keyFn(item)] = item;
    return acc;
  }, {});
}

/**
 * Debounce a function (for use in Apps Script triggers)
 * Uses PropertiesService to track last execution
 * @param {string} key - Unique key for this debounced function
 * @param {number} delayMs - Minimum time between executions
 * @returns {boolean} - True if function should execute, false if debounced
 */
function shouldExecute(key, delayMs = 1000) {
  const props = PropertiesService.getScriptProperties();
  const lastRunKey = 'debounce_' + key;
  const lastRun = parseInt(props.getProperty(lastRunKey) || '0', 10);
  const now = Date.now();
  
  if (now - lastRun < delayMs) {
    return false;
  }
  
  props.setProperty(lastRunKey, now.toString());
  return true;
}

/**
 * Retry a function with exponential backoff
 * @param {Function} fn - Function to execute
 * @param {number} maxRetries - Maximum number of retries
 * @param {number} baseDelayMs - Base delay in milliseconds
 * @returns {any} - Result of the function
 */
function retryWithBackoff(fn, maxRetries = 3, baseDelayMs = 1000) {
  let lastError;
  
  for (let i = 0; i <= maxRetries; i++) {
    try {
      return fn();
    } catch (e) {
      lastError = e;
      if (i < maxRetries) {
        const delay = baseDelayMs * Math.pow(2, i);
        Utilities.sleep(delay);
      }
    }
  }
  
  throw lastError;
}

/**
 * Measure execution time of a function
 * @param {string} label - Label for logging
 * @param {Function} fn - Function to measure
 * @returns {any} - Result of the function
 */
function measureTime(label, fn) {
  const start = Date.now();
  try {
    return fn();
  } finally {
    const elapsed = Date.now() - start;
    Logger.log(`${label}: ${elapsed}ms`);
  }
}

/**
 * Validate required parameters
 * @param {Object} params - Object of parameter names to values
 * @throws {Error} If any required parameter is missing
 */
function validateRequired(params) {
  const missing = Object.entries(params)
    .filter(([_, value]) => value === undefined || value === null || value === '')
    .map(([name]) => name);
  
  if (missing.length > 0) {
    throw new Error('Missing required parameters: ' + missing.join(', '));
  }
}

/**
 * Create a standard success response
 * @param {any} data - Data to include in response
 * @param {string} message - Success message
 * @returns {Object}
 */
function successResponse(data, message = 'Success') {
  return {
    success: true,
    message,
    data,
    timestamp: new Date().toISOString()
  };
}

/**
 * Create a standard error response
 * @param {string|Error} error - Error message or Error object
 * @returns {Object}
 */
function errorResponse(error) {
  const message = error instanceof Error ? error.message : String(error);
  return {
    success: false,
    message,
    error: message,
    timestamp: new Date().toISOString()
  };
}
