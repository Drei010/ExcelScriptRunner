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
      sheets: data.sheets,
      sheetNames: data.sheetNames,
      console: {
        log: (...args) => console.log('[ExcelScript]', ...args)
      },
      
      // ExcelScript-like workbook object
      workbook: {
        getActiveWorksheet: () => ({
          name: data.sheetNames[0] || 'Sheet1',
          getUsedRange: () => ({
            getRowCount: () => {
              const sheet = context.sheets[data.sheetNames[0]] || [];
              return sheet.length;
            },
            getColumnCount: () => {
              const sheet = context.sheets[data.sheetNames[0]] || [];
              return sheet.length > 0 ? sheet[0].length : 0;
            },
            getFormat: () => ({
              setHorizontalAlignment: (alignment) => {},
              setIndentLevel: (level) => {},
              setWrapText: (wrap) => {},
              setTextOrientation: (orientation) => {}
            })
          }),
          getRange: (range) => ({
            getTexts: () => {
              const sheet = context.sheets[data.sheetNames[0]] || [];
              const rangeData = parseRange(range, sheet);
              return rangeData.map(row => row.map(cell => String(cell || '')));
            },
            getValues: () => {
              const sheet = context.sheets[data.sheetNames[0]] || [];
              return parseRange(range, sheet);
            },
            setValue: (value) => {
              const sheet = context.sheets[data.sheetNames[0]] || [];
              const rangeData = parseRange(range, sheet);
              // Implementation for setting values
            },
            setNumberFormatLocal: (format) => {},
            getEntireRow: () => ({
              delete: (direction) => {
                // Implementation for row deletion
              }
            }),
            getSurroundingRegion: () => ({
              getLastRow: () => ({
                getRowIndex: () => {
                  const sheet = context.sheets[data.sheetNames[0]] || [];
                  return sheet.length;
                }
              })
            })
          }),
          getAutoFilter: () => null,
          addTable: (range, hasHeaders) => ({
            setPredefinedTableStyle: (style) => {}
          })
        }),
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
