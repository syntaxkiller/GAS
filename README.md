# Google Apps Script Survey Tool - Refactored

## Overview

This is a comprehensive refactoring of the survey data processing tool. The codebase has been reorganized for better maintainability, reduced duplication, and improved separation of concerns.

## Metrics

| Metric | Original | Refactored | Change |
|--------|----------|------------|--------|
| Total Lines | ~6,664 | ~6,080 | -9% |
| Files | 7 | 17 | +10 (smaller files) |
| Avg File Size | ~950 lines | ~358 lines | -62% |
| Max File Size | 2,731 lines | 681 lines | -75% |

## File Structure

```
├── Server (Google Apps Script .gs files)
│   ├── Constants.gs           (159 lines) - Centralized constants and configuration
│   ├── SpreadsheetDataManager.gs (555 lines) - All spreadsheet data access operations
│   ├── ConfigManager.gs       (477 lines) - Question configuration CRUD
│   ├── Core.gs                (527 lines) - Main app logic, menu, UI dialogs
│   ├── DataOperations.gs      (438 lines) - Faceting, mass editing, column combining
│   ├── ServerSync.gs          (320 lines) - Server upload/sync operations
│   ├── WordCloudProcessor.gs  (246 lines) - Text processing for word clouds
│   ├── DuplicateCleaner.gs    (233 lines) - Duplicate detection and cleanup
│   └── Utils.gs               (276 lines) - General utilities and helpers
│
├── UI (HTML files with CSS/JS)
│   ├── SharedStyles.html      (511 lines) - Common CSS styles
│   ├── SharedUtils.html       (296 lines) - Common JS utilities
│   ├── Sidebar.html           (382 lines) - Question linking sidebar template
│   ├── SidebarApp.html        (681 lines) - Sidebar Alpine.js logic
│   ├── QuestionConfig.html    (326 lines) - Config editor template
│   ├── QuestionConfigApp.html (341 lines) - Config editor Alpine.js logic
│   └── DuplicateCleanerUI.html (312 lines) - Duplicate manager UI
```

## Key Improvements

### 1. Eliminated Duplicate Code

**Before:** Duplicate `_processDataInColumns_` function defined twice in Core.gs
**After:** Single optimized version in SpreadsheetDataManager.gs

**Before:** Stop words defined in two places (WordCloudTextRemovals.gs and Core.gs)
**After:** Single source in WordCloudProcessor.gs with `getWordCloudStopWords()`

**Before:** Server proxy duplicated in Sidebar.html and QuestionConfig.html
**After:** Single definition in SharedUtils.html

**Before:** Common utilities (escapeHtml, toSentenceCase, etc.) copied between files
**After:** Single source in SharedUtils.html

### 2. Consolidated Data Access

**Before:** Direct sheet access scattered throughout code:
```javascript
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(year);
const ids = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
const colIndex = ids.findIndex(id => String(id).trim() === qId);
```

**After:** All access through SpreadsheetDataManager:
```javascript
const mgr = getDataManager();
const { sheet, colIndex, questionText } = mgr.findQuestionColumn(questionId);
```

### 3. Unified Config Operations

**Before:** Four similar config save functions with ~70% duplicate code:
- saveConfigToSheet()
- saveConfigToSheetBulk()
- saveMultipleConfigsToSheet()
- saveGroupConfigsToSheet()

**After:** Single ConfigManager class with clean methods:
```javascript
const mgr = new ConfigManager();
mgr.save(questionId, config);
mgr.saveMultiple(configsArray);
mgr.saveBulk(configObj);
mgr.saveGroups(groupConfigs);
```

### 4. Centralized Constants

**Before:** Magic numbers and strings scattered throughout:
```javascript
sheet.getRange(2, 1, lastRow - 1, 10);  // What's 10?
const CHUNK_SIZE = 5;                    // Why 5?
Utilities.sleep(100);                    // Why 100ms?
```

**After:** Named constants in Constants.gs:
```javascript
const CONFIG_COLUMN_COUNT = 10;
const RATE_LIMIT_BATCH_SIZE = 5;
const RATE_LIMIT_DELAY_MS = 100;
```

### 5. Modular HTML Structure

**Before:** Monolithic HTML files (QuestionConfig.html was 2,731 lines)

**After:** Separated concerns:
- Template files (*.html) - Just structure
- App files (*App.html) - Alpine.js logic
- Shared files - Common styles and utilities

Use the GAS include pattern:
```html
<?!= HtmlService.createHtmlOutputFromFile('SharedStyles').getContent(); ?>
<?!= HtmlService.createHtmlOutputFromFile('SharedUtils').getContent(); ?>
```

### 6. Consistent Error Handling

**Before:** Three different patterns (throw, return error objects, silent continuation)

**After:** Standardized approach:
- Server-side: Throw for exceptions, return `{success: boolean}` for expected results
- Client-side: Try/catch with toast notifications

### 7. Caching and Performance

**Before:** Repeated sheet lookups in loops

**After:** SpreadsheetDataManager with built-in caching:
```javascript
getSheet(name)           // Cached
getSheetQuestionIds(name) // Cached
findQuestionColumn(id)    // Cached
```

Plus singleton pattern:
```javascript
const mgr = getDataManager();  // Returns cached instance
```

## Usage

### Server-Side Patterns

```javascript
// Data access
const mgr = getDataManager();
const answers = mgr.getUniqueAnswers(questionId, { withCounts: true });
const stats = mgr.getNumericalStats(questionId);

// Config operations
const configMgr = new ConfigManager();
const allConfig = configMgr.loadAll();
configMgr.save(questionId, { questionType: 'Multiple Choice', chartType: 'bar' });

// Batch operations with rate limiting
mgr.batchReplaceValues(questionIds, replaceMap, { 
  batchSize: RATE_LIMIT_BATCH_SIZE,
  delayBetweenBatches: BATCH_WRITE_DELAY_MS
});
```

### Client-Side Patterns

```javascript
// Server calls (promisified)
const data = await server.getQuestionConfigData();

// Alpine.js app structure
function myApp() {
  return {
    // State
    items: [],
    isLoading: false,
    
    // Computed (use getters)
    get filteredItems() { ... },
    
    // Methods
    async initialize() { ... },
    handleError(e) { this.showToast(e.message, 'error'); }
  };
}
```

## Migration Notes

1. **Update include statements** in your HTML templates to use the new shared files
2. **Replace direct sheet access** with SpreadsheetDataManager methods
3. **Use ConfigManager** instead of individual config functions
4. **Import constants** from Constants.gs instead of hardcoding values

## Testing

After deployment, verify:
1. Menu items appear and open correct dialogs
2. Question linking works (save, edit, delete groups)
3. Configuration saves and loads correctly
4. Faceting and mass edit operations work
5. Server sync uploads data successfully
6. Duplicate cleaner finds and removes duplicates

## Future Improvements

1. Add JSDoc type annotations for better IDE support
2. Consider TypeScript with clasp for type safety
3. Add unit tests for server-side functions
4. Implement optimistic UI updates for better UX
5. Add error boundary pattern for graceful degradation
