/**
 * ExcelScript to JavaScript Converter
 * Converts Microsoft ExcelScript syntax to JavaScript-based ExcelScript
 */

/**
 * Convert ExcelScript syntax to JavaScript-based ExcelScript
 * @param {string} excelScript - Original ExcelScript code
 * @returns {string} - Converted JavaScript code
 */
export const convertExcelScript = (excelScript) => {
  let convertedScript = excelScript;
  
  // Remove function signature and convert to our syntax
  convertedScript = convertedScript.replace(
    /function\s+main\s*\(\s*workbook:\s*ExcelScript\.Workbook\s*\)\s*{/,
    '// Converted ExcelScript\nfunction main() {'
  );
  
  // Convert workbook.getActiveWorksheet() to getSheetNames()[0]
  convertedScript = convertedScript.replace(
    /workbook\.getActiveWorksheet\(\)/g,
    'workbook.getActiveWorksheet()'
  );
  
  // Convert workbook.getWorksheet() calls
  convertedScript = convertedScript.replace(
    /workbook\.getWorksheet\(([^)]+)\)/g,
    'workbook.getWorksheet($1)'
  );
  
  // Convert selectedSheet to activeSheet
  convertedScript = convertedScript.replace(
    /selectedSheet/g,
    'activeSheet'
  );
  
  // Convert range operations to our syntax
  convertedScript = convertedScript.replace(
    /activeSheet\.getRange\(([^)]+)\)\.getTexts\(\)/g,
    'getRangeTexts($1)'
  );
  
  convertedScript = convertedScript.replace(
    /activeSheet\.getRange\(([^)]+)\)\.getValues\(\)/g,
    'getRangeValues($1)'
  );
  
  convertedScript = convertedScript.replace(
    /activeSheet\.getRange\(([^)]+)\)\.setValue\(([^)]+)\)/g,
    'setRangeValue($1, $2)'
  );
  
  convertedScript = convertedScript.replace(
    /activeSheet\.getRange\(([^)]+)\)\.setNumberFormatLocal\(([^)]+)\)/g,
    'setRangeFormat($1, $2)'
  );
  
  // Convert row deletion
  convertedScript = convertedScript.replace(
    /activeSheet\.getRange\(([^)]+)\)\.getEntireRow\(\)\.delete\(([^)]+)\)/g,
    'deleteRow($1)'
  );
  
  // Convert table creation
  convertedScript = convertedScript.replace(
    /workbook\.addTable\(([^,]+),\s*([^)]+)\)/g,
    'createTable($1, $2)'
  );
  
  convertedScript = convertedScript.replace(
    /newTable\.setPredefinedTableStyle\(([^)]+)\)/g,
    'setTableStyle($1)'
  );
  
  // Convert autofilter operations
  convertedScript = convertedScript.replace(
    /activeSheet\.getAutoFilter\(\)\?\.remove\(\)/g,
    'removeAutoFilter()'
  );
  
  convertedScript = convertedScript.replace(
    /if\s*\(\s*activeSheet\.getAutoFilter\(\)\s*!==\s*null\s*\)\s*{[\s\S]*?activeSheet\.getAutoFilter\(\)\.remove\(\);[\s\S]*?}/g,
    'removeAutoFilter()'
  );
  
  // Convert setValues operations
  convertedScript = convertedScript.replace(
    /(\w+)\.getRange\(([^)]+)\)\.setValues\(([^)]+)\)/g,
    '$1.getRange($2).setValues($3)'
  );
  
  // Convert removeDuplicates operations
  convertedScript = convertedScript.replace(
    /(\w+)\.getRange\(([^)]+)\)\.removeDuplicates\(([^)]+)\)/g,
    '$1.getRange($2).removeDuplicates($3)'
  );
  
  // Convert autoFill operations
  convertedScript = convertedScript.replace(
    /(\w+)\.getRange\(([^)]+)\)\.autoFill\(([^)]+)\)/g,
    '$1.getRange($2).autoFill($3)'
  );
  
  // Convert clear operations
  convertedScript = convertedScript.replace(
    /(\w+)\.getRange\(([^)]+)\)\.clear\(([^)]+)\)/g,
    '$1.getRange($2).clear($3)'
  );
  
  // Convert sort operations
  convertedScript = convertedScript.replace(
    /(\w+)\.getAutoFilter\(\)\.getRange\(\)\.getSort\(\)\.apply\(([^)]+)\)/g,
    '$1.getAutoFilter().getRange().getSort().apply($2)'
  );
  
  // Add helper functions at the beginning
  const helperFunctions = `
// Helper functions for ExcelScript compatibility
function getActiveWorksheet() {
  const sheetName = getSheetNames()[0];
  return {
    name: sheetName,
    getUsedRange: () => ({
      getRowCount: () => getSheetData(sheetName).length,
      getColumnCount: () => {
        const data = getSheetData(sheetName);
        return data.length > 0 ? data[0].length : 0;
      },
      getFormat: () => ({
        setHorizontalAlignment: (alignment) => {},
        setIndentLevel: (level) => {},
        setWrapText: (wrap) => {},
        setTextOrientation: (orientation) => {}
      })
    }),
    getRange: (range) => ({
      getTexts: () => getRangeTexts(range),
      getValues: () => getRangeValues(range),
      setValue: (value) => setRangeValue(range, value),
      setNumberFormatLocal: (format) => setRangeFormat(range, format),
      getEntireRow: () => ({
        delete: (direction) => deleteRow(range)
      }),
      getSurroundingRegion: () => ({
        getLastRow: () => ({
          getRowIndex: () => getSheetData(sheetName).length
        })
      })
    }),
    getAutoFilter: () => null,
    addTable: (range, hasHeaders) => createTable(range, hasHeaders)
  };
}

function getRangeTexts(range) {
  const sheetName = getSheetNames()[0];
  const data = getSheetData(sheetName);
  return parseRange(range, data).map(row => row.map(cell => String(cell || '')));
}

function getRangeValues(range) {
  const sheetName = getSheetNames()[0];
  const data = getSheetData(sheetName);
  return parseRange(range, data);
}

function setRangeValue(range, value) {
  const sheetName = getSheetNames()[0];
  const data = getSheetData(sheetName);
  const rangeData = parseRange(range, data);
  // Implementation for setting values
  console.log('Setting range', range, 'to value', value);
}

function setRangeFormat(range, format) {
  console.log('Setting range', range, 'format to', format);
}

function deleteRow(range) {
  console.log('Deleting row at range', range);
}

function createTable(range, hasHeaders) {
  console.log('Creating table for range', range, 'with headers:', hasHeaders);
  return {
    setPredefinedTableStyle: (style) => {
      console.log('Setting table style to', style);
    }
  };
}

function setTableStyle(style) {
  console.log('Setting table style to', style);
}

function removeAutoFilter() {
  console.log('Removing autofilter');
}

function parseRange(range, sheet) {
  if (!range || !sheet) return [];
  
  // Handle single cell ranges like "A1"
  if (/^[A-Z]+\\d+$/.test(range)) {
    const match = range.match(/^([A-Z]+)(\\d+)$/);
    if (match) {
      const col = columnToIndex(match[1]);
      const row = parseInt(match[2]) - 1;
      return [[sheet[row] && sheet[row][col]]];
    }
  }
  
  // Handle column ranges like "O1:O10"
  if (/^[A-Z]+\\d+:[A-Z]+\\d+$/.test(range)) {
    const [start, end] = range.split(':');
    const startMatch = start.match(/^([A-Z]+)(\\d+)$/);
    const endMatch = end.match(/^([A-Z]+)(\\d+)$/);
    
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
}

function columnToIndex(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return result - 1;
}

// ExcelScript constants
const ExcelScript = {
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

`;

  return helperFunctions + convertedScript;
};

/**
 * Check if script contains ExcelScript syntax
 * @param {string} script - Script to check
 * @returns {boolean} - True if contains ExcelScript syntax
 */
export const isExcelScript = (script) => {
  return script.includes('function main(workbook: ExcelScript.Workbook)') ||
         script.includes('workbook.getActiveWorksheet()') ||
         script.includes('ExcelScript.');
};
