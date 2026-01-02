/**
 * WordCloudTextRemovals.gs
 * Contains lists of common words to remove for word cloud generation
 * and functions to clean text for word cloud visualization
 */

/**
 * Get the list of common words to remove from word clouds.
 * You can customize this list by adding or removing words.
 * @returns {Array<string>} Array of lowercase words to filter out
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
    
    // Survey-specific common words (customize these for your needs)
    'yes', 'no', 'maybe', 'na', 'n/a', 'none', 'unknown', 'other',
    'prefer', 'answer', 'say', 'rather', 'somewhat', 'very'
  ];
}

/**
 * Clean a batch of text values for word cloud generation.
 * @param {Array<string>} values - Array of text values to clean
 * @param {Object} options - Cleaning options
 * @returns {Array<string>} Array of cleaned values in same order
 */
function cleanTextForWordCloudBatch(values, options) {
  if (!Array.isArray(values)) {
    throw new Error('values must be an array');
  }
  
  return values.map(value => {
    return cleanTextForWordCloud(value, options);
  });
}

/**
 * Remove punctuation from text.
 * Keeps letters, numbers, and spaces.
 * @param {string} text - The text to clean
 * @returns {string} Text with punctuation removed
 */
function removePunctuation(text) {
  if (!text || typeof text !== 'string') return '';
  
  // Replace all punctuation with spaces, then collapse multiple spaces
  return text
    .replace(/[^\w\s]|_/g, ' ')  // Replace punctuation with spaces
    .replace(/\s+/g, ' ')          // Collapse multiple spaces
    .trim();                       // Trim leading/trailing spaces
}

/**
 * Remove common stop words from text.
 * @param {string} text - The text to filter
 * @param {boolean} includeCustomWords - Whether to also filter custom words (currently unused)
 * @returns {string} Text with stop words removed
 */
function removeStopWords(text, includeCustomWords) {
  if (!text || typeof text !== 'string') return '';
  
  const stopWords = getWordCloudStopWords();
  // Note: includeCustomWords parameter exists for future extensibility
  // Custom words can be managed through Core.gs getCustomRemovalWords() if needed
  const allRemovalWords = new Set(stopWords.map(w => w.toLowerCase()));
  
  // Split into words, filter, and rejoin
  const words = text.toLowerCase().split(/\s+/);
  const filtered = words.filter(word => {
    // Remove if empty or in removal list
    if (!word || word.length === 0) return false;
    if (allRemovalWords.has(word)) return false;
    // Keep numbers
    if (/^\d+$/.test(word)) return true;
    // Keep words that are at least 2 characters
    return word.length >= 2;
  });
  
  return filtered.join(' ');
}

/**
 * Clean text for word cloud generation.
 * @param {string} text - The text to clean
 * @param {Object} options - Cleaning options
 * @param {boolean} options.removePunctuation - Whether to remove punctuation
 * @param {boolean} options.removeStopWords - Whether to remove stop words
 * @param {boolean} options.removeCustomWords - Whether to remove custom words
 * @param {string} options.caseTransform - Case transformation: 'none', 'lower', 'upper', 'sentence', 'capital'
 * @returns {string} Cleaned text with words comma-separated
 */
function cleanTextForWordCloud(text, options) {
  if (!text || typeof text !== 'string') return '';
  
  options = options || {};
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
  if (cleaned && !options.preserveSpaces) {
    cleaned = cleaned.split(/\s+/).filter(w => w).join(', ');
  }
  
  return cleaned;
}