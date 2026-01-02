/**
 * WordCloudProcessor.gs
 * Text processing utilities for word cloud generation
 * Consolidates all word cloud related functionality
 */

// === STOP WORDS LIST ===

/**
 * Get the comprehensive list of stop words to filter from word clouds
 * @returns {string[]} Array of lowercase words to filter out
 */
function getWordCloudStopWords() {
  return [
    // Articles
    'a', 'an', 'the',

    // Common pronouns
    'i', 'me', 'my', 'myself', 'we', 'our', 'ours', 'ourselves',
    'you', 'your', 'yours', 'yourself', 'yourselves',
    'he', 'him', 'his', 'himself', 'she', 'her', 'hers', 'herself',
    'it', 'its', 'itself', 'they', 'them', 'their', 'theirs', 'themselves',

    // Common verbs
    'am', 'is', 'are', 'was', 'were', 'be', 'been', 'being',
    'have', 'has', 'had', 'having', 'do', 'does', 'did', 'doing',
    'would', 'should', 'could', 'ought', 'can', 'may', 'might', 'must',
    'will', 'shall',

    // Prepositions
    'of', 'at', 'by', 'for', 'with', 'about', 'against', 'between',
    'into', 'through', 'during', 'before', 'after', 'above', 'below',
    'to', 'from', 'up', 'down', 'in', 'out', 'on', 'off', 'over',
    'under', 'again', 'further', 'then', 'once',

    // Conjunctions
    'and', 'but', 'or', 'nor', 'so', 'yet', 'if', 'because', 'as',
    'until', 'while', 'when', 'where', 'why', 'how',

    // Other common words
    'all', 'any', 'both', 'each', 'few', 'more', 'most', 'other',
    'some', 'such', 'no', 'not', 'only', 'own', 'same', 'than',
    'too', 'very', 'just', 'now', 'also', 'here', 'there',

    // Common contractions (already split)
    'don', 't', 's', 're', 've', 'll', 'd', 'm',
    'doesn', 'didn', 'won', 'wouldn', 'shouldn', 'couldn', 'hasn',
    'haven', 'hadn', 'isn', 'aren', 'wasn', 'weren', 'cant', 'cannot',

    // Survey-specific common words
    'yes', 'no', 'maybe', 'na', 'n/a', 'none', 'unknown', 'other',
    'prefer', 'answer', 'say', 'rather', 'somewhat', 'very'
  ];
}

// === CUSTOM WORDS MANAGEMENT ===

/**
 * Get user's custom removal words
 * @returns {string[]}
 */
function getCustomRemovalWords() {
  try {
    const userProps = PropertiesService.getUserProperties();
    const customWords = userProps.getProperty(WORDCLOUD_CUSTOM_WORDS_KEY);
    return customWords ? JSON.parse(customWords) : [];
  } catch (err) {
    Logger.log('Error getting custom words: ' + err.message);
    return [];
  }
}

/**
 * Set user's custom removal words
 * @param {string[]} words
 * @returns {boolean}
 */
function setCustomRemovalWords(words) {
  try {
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty(WORDCLOUD_CUSTOM_WORDS_KEY, JSON.stringify(words));
    return true;
  } catch (err) {
    Logger.log('Error setting custom words: ' + err.message);
    return false;
  }
}

// === TEXT PROCESSING ===

/**
 * Remove punctuation from text
 * @param {string} text
 * @returns {string}
 */
function removePunctuation(text) {
  if (!text || typeof text !== 'string') return '';

  return text
    .replace(/[^\w\s]|_/g, ' ')  // Replace punctuation with spaces
    .replace(/\s+/g, ' ')        // Collapse multiple spaces
    .trim();
}

/**
 * Remove stop words from text
 * @param {string} text
 * @param {boolean} includeCustomWords - Whether to also filter custom words
 * @returns {string}
 */
function removeStopWords(text, includeCustomWords = false) {
  if (!text || typeof text !== 'string') return '';

  const stopWords = getWordCloudStopWords();
  const allRemovalWords = new Set(stopWords.map(w => w.toLowerCase()));

  // Add custom words if requested
  if (includeCustomWords) {
    getCustomRemovalWords().forEach(w => allRemovalWords.add(w.toLowerCase()));
  }

  // Split, filter, and rejoin
  const words = text.toLowerCase().split(/\s+/);
  const filtered = words.filter(word => {
    if (!word || word.length === 0) return false;
    if (allRemovalWords.has(word)) return false;
    if (/^\d+$/.test(word)) return true;  // Keep numbers
    return word.length >= 2;  // Keep words 2+ characters
  });

  return filtered.join(' ');
}

/**
 * Clean text for word cloud generation
 * @param {string} text
 * @param {Object} options
 * @param {boolean} options.removePunctuation - Whether to remove punctuation
 * @param {boolean} options.removeStopWords - Whether to remove stop words
 * @param {boolean} options.removeCustomWords - Whether to remove custom words
 * @param {boolean} options.preserveSpaces - If false, convert spaces to commas
 * @returns {string}
 */
function cleanTextForWordCloud(text, options = {}) {
  if (!text || typeof text !== 'string') return '';

  let cleaned = text;

  // Remove punctuation first if requested
  if (options.removePunctuation) {
    cleaned = removePunctuation(cleaned);
  }

  // Remove stop words if requested
  if (options.removeStopWords) {
    cleaned = removeStopWords(cleaned, options.removeCustomWords || false);
  }

  cleaned = cleaned.trim();

  // Convert spaces to commas unless preserveSpaces is true
  if (cleaned && !options.preserveSpaces) {
    cleaned = cleaned.split(/\s+/).filter(w => w).join(', ');
  }

  return cleaned;
}

/**
 * Clean a batch of text values for word cloud generation
 * @param {string[]} values
 * @param {Object} options
 * @returns {string[]}
 */
function cleanTextForWordCloudBatch(values, options = {}) {
  if (!Array.isArray(values)) {
    throw new Error('values must be an array');
  }

  return values.map(value => cleanTextForWordCloud(value, options));
}

/**
 * Process text and return word frequencies
 * @param {string} text
 * @param {Object} options
 * @returns {Object} - { word: count, ... }
 */
function processTextForWordCloud(text, options = {}) {
  if (!text || typeof text !== 'string') return {};

  let processed = text.toLowerCase();

  // Remove punctuation if requested
  if (options.removePunctuation) {
    processed = removePunctuation(processed);
  }

  // Split into words
  let words = processed.split(/\s+/).filter(w => w.length > 0);

  // Remove stop words if requested
  if (options.removeStopWords) {
    const stopWords = new Set(getWordCloudStopWords());
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

// === CASE TRANSFORMATION ===

/**
 * Convert text to sentence case
 * @param {string} str
 * @returns {string}
 */
function toSentenceCase(str) {
  if (!str) return '';
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}

/**
 * Convert text to title/capital case
 * @param {string} str
 * @returns {string}
 */
function toCapitalCase(str) {
  if (!str) return '';
  return str.toLowerCase()
    .split(' ')
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
}
