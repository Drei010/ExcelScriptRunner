/**
 * ExcelScript Execution Engine
 * Processes ExcelScript commands on spreadsheet data
 */

/**
 * Parse Excel range notation (e.g., "A1:B10", "O1:O10")
 * @param {string} range - Excel range notation
 * @param {Array} sheet - Sheet data
 * @returns {Array} - Parsed range data
 */
const parseRange = (range, sheet) => {
  if (!range || !sheet) return [];
  
  // Handle single cell ranges like "A1"
  if (/^[A-Z]+\d+$/.test(range)) {
    const match = range.match(/^([A-Z]+)(\d+)$/);
    if (match) {
      const col = columnToIndex(match[1]);
      const row = parseInt(match[2]) - 1;
      return [[sheet[row] && sheet[row][col]]];
    }
  }
  
  // Handle column ranges like "O1:O10"
  if (/^[A-Z]+\d+:[A-Z]+\d+$/.test(range)) {
    const [start, end] = range.split(':');
    const startMatch = start.match(/^([A-Z]+)(\d+)$/);
    const endMatch = end.match(/^([A-Z]+)(\d+)$/);
    
    if (startMatch && endMatch) {
      const col = columnToIndex(startMatch[1]);
      const startRow = parseInt(startMatch[2]) - 1;
      const endRow = parseInt(endMatch[2]) - 1;
      
      const result = [];
      for (let i = startRow; i <= endRow; i++) {
        result.push([sheet[i] && sheet[i][col]]);
      }
      return result;
    }
  }
  
  return [];
};

/**
 * Convert Excel column letter to index (A=0, B=1, etc.)
 * @param {string} column - Column letter(s)
 * @returns {number} - Column index
 */
const columnToIndex = (column) => {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return result - 1;
};

/**
 * Convert column index to Excel letter (0=A, 1=B, etc.)
 * @param {number} index - Column index
 * @returns {string} - Column letter(s)
 */
const indexToColumn = (index) => {
  let result = '';
  while (index >= 0) {
    result = String.fromCharCode(65 + (index % 26)) + result;
    index = Math.floor(index / 26) - 1;
  }
  return result;
};

/**
 * Create a worksheet object with all ExcelScript methods
 * @param {Object} context - Execution context
 * @param {string} sheetName - Name of the sheet
 * @returns {Object} - Worksheet object
 */
const createWorksheetObject = (context, sheetName) => {
  const sheet = context.sheets[sheetName] || [];
  
  return {
    name: sheetName,
    getUsedRange: () => ({
      getRowCount: () => sheet.length,
      getColumnCount: () => sheet.length > 0 ? sheet[0].length : 0,
      getFormat: () => ({
        setHorizontalAlignment: (alignment) => {},
        setIndentLevel: (level) => {},
        setWrapText: (wrap) => {},
        setTextOrientation: (orientation) => {}
      })
    }),
    getRange: (range) => {
      const rangeData = parseRange(range, sheet);
      const rangeInfo = parseRangeInfo(range);
      
      return {
        getTexts: () => rangeData.map(row => row.map(cell => String(cell || ''))),
        getValues: () => rangeData,
        setValue: (value) => {
          if (rangeInfo.isSingleCell) {
            setCellValue(context, sheetName, rangeInfo.startRow, rangeInfo.startCol, value);
          }
        },
        setValues: (values) => {
          if (Array.isArray(values)) {
            setRangeValues(context, sheetName, rangeInfo, values);
          }
        },
        setNumberFormatLocal: (format) => {},
        getEntireRow: () => ({
          delete: (direction) => {
            if (rangeInfo.isSingleCell) {
              deleteRow(context, sheetName, rangeInfo.startRow);
            }
          }
        }),
        getSurroundingRegion: () => ({
          getLastRow: () => ({
            getRowIndex: () => sheet.length
          })
        }),
        removeDuplicates: (columns, hasHeaders) => {
          removeDuplicatesFromRange(context, sheetName, rangeInfo, columns, hasHeaders);
        },
        autoFill: (destinationRange, fillType) => {
          const destInfo = parseRangeInfo(destinationRange);
          autoFillRange(context, sheetName, rangeInfo, destInfo, fillType);
        },
        clear: (applyTo) => {
          clearRange(context, sheetName, rangeInfo, applyTo);
        }
      };
    },
    getAutoFilter: () => ({
      getRange: () => ({
        getSort: () => ({
          apply: (key, matchCase, hasHeaders) => {
            sortRange(context, sheetName, key, matchCase, hasHeaders);
          }
        })
      }),
      remove: () => {}
    }),
    addTable: (range, hasHeaders) => ({
      setPredefinedTableStyle: (style) => {}
    })
  };
};

/**
 * Parse range information for advanced operations
 * @param {string} range - Excel range notation
 * @returns {Object} - Range information
 */
const parseRangeInfo = (range) => {
  if (!range) return { isSingleCell: false };
  
  // Handle single cell ranges like "A1"
  if (/^[A-Z]+\d+$/.test(range)) {
    const match = range.match(/^([A-Z]+)(\d+)$/);
    if (match) {
      return {
        isSingleCell: true,
        startRow: parseInt(match[2]) - 1,
        endRow: parseInt(match[2]) - 1,
        startCol: columnToIndex(match[1]),
        endCol: columnToIndex(match[1])
      };
    }
  }
  
  // Handle column ranges like "A1:A10"
  if (/^[A-Z]+\d+:[A-Z]+\d+$/.test(range)) {
    const [start, end] = range.split(':');
    const startMatch = start.match(/^([A-Z]+)(\d+)$/);
    const endMatch = end.match(/^([A-Z]+)(\d+)$/);
    
    if (startMatch && endMatch) {
      return {
        isSingleCell: false,
        startRow: parseInt(startMatch[2]) - 1,
        endRow: parseInt(endMatch[2]) - 1,
        startCol: columnToIndex(startMatch[1]),
        endCol: columnToIndex(endMatch[1])
      };
    }
  }
  
  return { isSingleCell: false };
};

/**
 * Set cell value in context
 * @param {Object} context - Execution context
 * @param {string} sheetName - Sheet name
 * @param {number} row - Row index
 * @param {number} col - Column index
 * @param {*} value - Value to set
 */
const setCellValue = (context, sheetName, row, col, value) => {
  if (!context.sheets[sheetName]) {
    context.sheets[sheetName] = [];
  }
  if (!context.sheets[sheetName][row]) {
    context.sheets[sheetName][row] = [];
  }
  context.sheets[sheetName][row][col] = value;
};

/**
 * Set range values in context
 * @param {Object} context - Execution context
 * @param {string} sheetName - Sheet name
 * @param {Object} rangeInfo - Range information
 * @param {Array} values - Values to set
 */
const setRangeValues = (context, sheetName, rangeInfo, values) => {
  if (!context.sheets[sheetName]) {
    context.sheets[sheetName] = [];
  }
  
  values.forEach((row, rowIndex) => {
    const actualRow = rangeInfo.startRow + rowIndex;
    if (!context.sheets[sheetName][actualRow]) {
      context.sheets[sheetName][actualRow] = [];
    }
    
    if (Array.isArray(row)) {
      row.forEach((cell, colIndex) => {
        const actualCol = rangeInfo.startCol + colIndex;
        context.sheets[sheetName][actualRow][actualCol] = cell;
      });
    } else {
      context.sheets[sheetName][actualRow][rangeInfo.startCol] = row;
    }
  });
};

/**
 * Delete row from context
 * @param {Object} context - Execution context
 * @param {string} sheetName - Sheet name
 * @param {number} rowIndex - Row index to delete
 */
const deleteRow = (context, sheetName, rowIndex) => {
  if (context.sheets[sheetName] && context.sheets[sheetName][rowIndex]) {
    context.sheets[sheetName].splice(rowIndex, 1);
  }
};

/**
 * Remove duplicates from range
 * @param {Object} context - Execution context
 * @param {string} sheetName - Sheet name
 * @param {Object} rangeInfo - Range information
 * @param {Array} columns - Columns to check for duplicates
 * @param {boolean} hasHeaders - Whether range has headers
 */
const removeDuplicatesFromRange = (context, sheetName, rangeInfo, columns, hasHeaders) => {
  const sheet = context.sheets[sheetName];
  if (!sheet) return;
  
  const startRow = hasHeaders ? rangeInfo.startRow + 1 : rangeInfo.startRow;
  const endRow = rangeInfo.endRow;
  
  // Simple duplicate removal based on first column
  const seen = new Set();
  const newRows = [];
  
  for (let i = 0; i < sheet.length; i++) {
    if (i < startRow || i > endRow) {
      newRows.push(sheet[i]);
      continue;
    }
    
    const key = String(sheet[i][rangeInfo.startCol] || '');
    if (!seen.has(key)) {
      seen.add(key);
      newRows.push(sheet[i]);
    }
  }
  
  context.sheets[sheetName] = newRows;
};

/**
 * Auto fill range
 * @param {Object} context - Execution context
 * @param {string} sheetName - Sheet name
 * @param {Object} sourceInfo - Source range info
 * @param {Object} destInfo - Destination range info
 * @param {string} fillType - Fill type
 */
const autoFillRange = (context, sheetName, sourceInfo, destInfo, fillType) => {
  const sheet = context.sheets[sheetName];
  if (!sheet) return;
  
  // Simple copy fill for now
  for (let row = destInfo.startRow; row <= destInfo.endRow; row++) {
    if (!sheet[row]) sheet[row] = [];
    for (let col = destInfo.startCol; col <= destInfo.endCol; col++) {
      const sourceRow = sourceInfo.startRow + (row - destInfo.startRow) % (sourceInfo.endRow - sourceInfo.startRow + 1);
      const sourceCol = sourceInfo.startCol + (col - destInfo.startCol) % (sourceInfo.endCol - sourceInfo.startCol + 1);
      if (sheet[sourceRow] && sheet[sourceRow][sourceCol] !== undefined) {
        sheet[row][col] = sheet[sourceRow][sourceCol];
      }
    }
  }
};

/**
 * Clear range
 * @param {Object} context - Execution context
 * @param {string} sheetName - Sheet name
 * @param {Object} rangeInfo - Range information
 * @param {string} applyTo - What to clear
 */
const clearRange = (context, sheetName, rangeInfo, applyTo) => {
  const sheet = context.sheets[sheetName];
  if (!sheet) return;
  
  for (let row = rangeInfo.startRow; row <= rangeInfo.endRow; row++) {
    if (sheet[row]) {
      for (let col = rangeInfo.startCol; col <= rangeInfo.endCol; col++) {
        if (applyTo === 'contents' || applyTo === 'all') {
          sheet[row][col] = '';
        }
      }
    }
  }
};

/**
 * Sort range
 * @param {Object} context - Execution context
 * @param {string} sheetName - Sheet name
 * @param {Array} keys - Sort keys
 * @param {boolean} matchCase - Match case
 * @param {boolean} hasHeaders - Has headers
 */
const sortRange = (context, sheetName, keys, matchCase, hasHeaders) => {
  const sheet = context.sheets[sheetName];
  if (!sheet || sheet.length <= 1) return;
  
  const startRow = hasHeaders ? 1 : 0;
  const dataRows = sheet.slice(startRow);
  
  dataRows.sort((a, b) => {
    for (const key of keys) {
      const colIndex = key.key || key;
      const aVal = a[colIndex] || '';
      const bVal = b[colIndex] || '';
      
      let comparison = 0;
      if (typeof aVal === 'string' && typeof bVal === 'string') {
        comparison = matchCase ? aVal.localeCompare(bVal) : aVal.toLowerCase().localeCompare(bVal.toLowerCase());
      } else {
        comparison = aVal < bVal ? -1 : aVal > bVal ? 1 : 0;
      }
      
      if (comparison !== 0) {
        return key.ascending !== false ? comparison : -comparison;
      }
    }
    return 0;
  });
  
  // Replace the data rows
  sheet.splice(startRow, dataRows.length, ...dataRows);
};

/**
 * Execute ExcelScript on parsed data
 * @param {Object} data - Parsed Excel/CSV data
 * @param {string} script - ExcelScript code
 * @returns {Object} - Processed data
 */
export const executeScript = (data, script) => {
  try {
    // Create a safe execution context
    const context = {
      data: JSON.parse(JSON.stringify(data)), // Deep clone
      sheets: JSON.parse(JSON.stringify(data.sheets)), // Deep clone sheets too
      sheetNames: [...data.sheetNames], // Clone array
      console: {
        log: (...args) => console.log('[ExcelScript]', ...args)
      },
      
      // ExcelScript-like workbook object
      workbook: {
        getActiveWorksheet: () => createWorksheetObject(context, data.sheetNames[0] || 'Sheet1'),
        getWorksheet: (name) => {
          if (context.sheets[name]) {
            return createWorksheetObject(context, name);
          }
          return null;
        },
        addTable: (range, hasHeaders) => ({
          setPredefinedTableStyle: (style) => {}
        })
      }
    };
    
    // Add helper functions to context
    context.getCell = (sheetName, row, col) => {
      const sheet = context.sheets[sheetName];
      if (!sheet || !sheet[row] || sheet[row][col] === undefined) {
        return null;
      }
      return sheet[row][col];
    };
    
    context.setCell = (sheetName, row, col, value) => {
      if (!context.sheets[sheetName]) {
        context.sheets[sheetName] = [];
      }
      if (!context.sheets[sheetName][row]) {
        context.sheets[sheetName][row] = [];
      }
      context.sheets[sheetName][row][col] = value;
    };
    
    context.getRow = (sheetName, rowIndex) => {
      const sheet = context.sheets[sheetName];
      return sheet ? sheet[rowIndex] || [] : [];
    };
    
    context.setRow = (sheetName, rowIndex, rowData) => {
      if (!context.sheets[sheetName]) {
        context.sheets[sheetName] = [];
      }
      context.sheets[sheetName][rowIndex] = rowData;
    };
    
    context.getColumn = (sheetName, colIndex) => {
      const sheet = context.sheets[sheetName];
      if (!sheet) return [];
      return sheet.map(row => row[colIndex]).filter(val => val !== undefined);
    };
    
    context.setColumn = (sheetName, colIndex, columnData) => {
      if (!context.sheets[sheetName]) {
        context.sheets[sheetName] = [];
      }
      columnData.forEach((value, rowIndex) => {
        if (!context.sheets[sheetName][rowIndex]) {
          context.sheets[sheetName][rowIndex] = [];
        }
        context.sheets[sheetName][rowIndex][colIndex] = value;
      });
    };
    
    context.addSheet = (sheetName, data = []) => {
      context.sheets[sheetName] = data;
      context.sheetNames.push(sheetName);
    };
    
    context.removeSheet = (sheetName) => {
      delete context.sheets[sheetName];
      context.sheetNames = context.sheetNames.filter(name => name !== sheetName);
    };
    
    context.getSheetNames = () => context.sheetNames;
    
    context.getSheetData = (sheetName) => {
      return context.sheets[sheetName] || [];
    };
    
    context.setSheetData = (sheetName, data) => {
      context.sheets[sheetName] = data;
    };
    
    // Math functions
    context.sum = (...values) => values.reduce((a, b) => (a || 0) + (b || 0), 0);
    context.avg = (...values) => {
      const validValues = values.filter(v => typeof v === 'number' && !isNaN(v));
      return validValues.length > 0 ? validValues.reduce((a, b) => a + b, 0) / validValues.length : 0;
    };
    context.min = (...values) => Math.min(...values.filter(v => typeof v === 'number' && !isNaN(v)));
    context.max = (...values) => Math.max(...values.filter(v => typeof v === 'number' && !isNaN(v)));
    
    // String functions
    context.upper = (str) => String(str || '').toUpperCase();
    context.lower = (str) => String(str || '').toLowerCase();
    context.trim = (str) => String(str || '').trim();
    context.concat = (...values) => values.map(v => String(v || '')).join('');
    
    // Date functions
    context.now = () => new Date();
    context.today = () => new Date().toISOString().split('T')[0];
    
    // ExcelScript constants
    context.ExcelScript = {
      HorizontalAlignment: {
        general: 'general',
        left: 'left',
        center: 'center',
        right: 'right'
      },
      DeleteShiftDirection: {
        up: 'up',
        left: 'left'
      },
      AutoFillType: {
        fillDefault: 'fillDefault',
        fillCopy: 'fillCopy',
        fillSeries: 'fillSeries',
        fillFormats: 'fillFormats',
        fillValues: 'fillValues',
        fillDays: 'fillDays',
        fillWeekdays: 'fillWeekdays',
        fillMonths: 'fillMonths',
        fillYears: 'fillYears',
        fillLinearTrend: 'fillLinearTrend',
        fillGrowthTrend: 'fillGrowthTrend',
        fillFlashFill: 'fillFlashFill'
      },
      ClearApplyTo: {
        all: 'all',
        contents: 'contents',
        formats: 'formats',
        hyperlinks: 'hyperlinks'
      }
    };
    
    // Enhanced date functions
    context.Date = Date;
    context.newDate = () => new Date();
    
    // Execute the script
    const scriptFunction = new Function('context', `
      with (context) {
        ${script}
      }
    `);
    
    scriptFunction(context);
    
    // Update the data object with the modified sheets
    context.data.sheets = context.sheets;
    context.data.sheetNames = context.sheetNames;
    
    return context.data;
  } catch (error) {
    throw new Error(`Script execution error: ${error.message}`);
  }
};

/**
 * Validate ExcelScript syntax
 * @param {string} script - ExcelScript code
 * @returns {Object} - Validation result
 */
export const validateScript = (script) => {
  const errors = [];
  const warnings = [];
  
  if (!script || script.trim().length === 0) {
    errors.push('Script cannot be empty');
    return { isValid: false, errors, warnings };
  }
  
  // Check for potentially dangerous operations
  const dangerousPatterns = [
    /eval\s*\(/,
    /Function\s*\(/,
    /setTimeout\s*\(/,
    /setInterval\s*\(/,
    /document\./,
    /window\./,
    /require\s*\(/,
    /import\s+/,
    /export\s+/
  ];
  
  dangerousPatterns.forEach(pattern => {
    if (pattern.test(script)) {
      errors.push(`Potentially dangerous operation detected: ${pattern.source}`);
    }
  });
  
  // Check for common syntax errors
  const openBraces = (script.match(/\{/g) || []).length;
  const closeBraces = (script.match(/\}/g) || []).length;
  if (openBraces !== closeBraces) {
    errors.push('Mismatched braces in script');
  }
  
  const openParens = (script.match(/\(/g) || []).length;
  const closeParens = (script.match(/\)/g) || []).length;
  if (openParens !== closeParens) {
    errors.push('Mismatched parentheses in script');
  }
  
  return {
    isValid: errors.length === 0,
    errors,
    warnings
  };
};
