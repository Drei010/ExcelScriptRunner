import React, { useState } from 'react';
import { validateScript } from '../utils/scriptEngine';
import { convertExcelScript, isExcelScript } from '../utils/scriptConverter';

const ScriptInput = ({ onScriptChange, script, error }) => {
  const [validationResult, setValidationResult] = useState(null);
  const [isExcelScriptSyntax, setIsExcelScriptSyntax] = useState(false);
  const [convertedScript, setConvertedScript] = useState('');

  const handleScriptChange = (e) => {
    const newScript = e.target.value;
    onScriptChange(newScript);
    
    // Check if it's ExcelScript syntax
    const isExcel = isExcelScript(newScript);
    setIsExcelScriptSyntax(isExcel);
    
    if (isExcel) {
      try {
        const converted = convertExcelScript(newScript);
        setConvertedScript(converted);
        const result = validateScript(converted);
        setValidationResult(result);
      } catch (error) {
        setValidationResult({
          isValid: false,
          errors: [`Conversion error: ${error.message}`],
          warnings: []
        });
      }
    } else {
      setConvertedScript('');
      // Validate script in real-time
      if (newScript.trim()) {
        const result = validateScript(newScript);
        setValidationResult(result);
      } else {
        setValidationResult(null);
      }
    }
  };

  const exampleScript = `// Example ExcelScript - Process sales data
// Get the first sheet
const sheetName = getSheetNames()[0];
const data = getSheetData(sheetName);

// Add a new column for total sales
data.forEach((row, index) => {
  if (index === 0) {
    // Header row
    row.push('Total Sales');
  } else {
    // Data rows - sum columns 2 and 3 (assuming price and quantity)
    const price = parseFloat(row[2]) || 0;
    const quantity = parseFloat(row[3]) || 0;
    const total = price * quantity;
    row.push(total);
  }
});

// Update the sheet with modified data
setSheetData(sheetName, data);

console.log('Processing complete!');`;

  const excelScriptExample = `function main(workbook: ExcelScript.Workbook) {
    const selectedSheet = workbook.getActiveWorksheet();
    
    // Get current date
    const currentDate = new Date();
    const month = currentDate.getMonth() + 1;
    const day = currentDate.getDate();
    const year = currentDate.getFullYear();
    const hours = currentDate.getHours();
    const minutes = currentDate.getMinutes();
    const seconds = currentDate.getSeconds();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    
    const dateTimeString = \`\${month}/\${day}/\${year} \${hours % 12 || 12}:\${minutes.toString().padStart(2, '0')}:\${seconds.toString().padStart(2, '0')} \${ampm}\`;
    
    // Get the entire used range
    const usedRange = selectedSheet.getUsedRange();
    if (!usedRange) {
        console.log("No data found in the sheet.");
        return;
    }
    
    // Update column T with current date/time
    const lastRowA = selectedSheet.getRange("A1").getSurroundingRegion().getLastRow().getRowIndex();
    const targetRange = selectedSheet.getRange(\`T2:T\${lastRowA + 2}\`);
    targetRange.setValue(dateTimeString);
}`;

  return (
    <div className="section">
      <h2>📝 ExcelScript Input</h2>
      
      <div className="script-input-content">
        <textarea
          className="script-textarea"
          placeholder="Enter your ExcelScript here...\n\nExample:\n// Get data from first sheet\nconst sheetName = getSheetNames()[0];\nconst data = getSheetData(sheetName);\n\n// Process the data\ndata.forEach((row, index) => {\n  if (index > 0) { // Skip header\n    // Your processing logic here\n  }\n});\n\n// Update the sheet\nsetSheetData(sheetName, data);"
          value={script}
          onChange={handleScriptChange}
          aria-label="ExcelScript Input Text Area"
        />
        
        {isExcelScriptSyntax && (
          <div className="info-box">
            <strong>🔄 ExcelScript Detected!</strong>
            <p>
              Your script uses Microsoft ExcelScript syntax. It will be automatically converted to our JavaScript-based ExcelScript.
            </p>
          </div>
        )}

        {validationResult && !validationResult.isValid && (
          <div className="error-message">
            <strong>⚠️ Script Validation Errors:</strong>
            <ul style={{ marginTop: '10px', marginLeft: '20px' }}>
              {validationResult.errors.map((error, index) => (
                <li key={index}>{error}</li>
              ))}
            </ul>
          </div>
        )}
        
        {validationResult && validationResult.isValid && (
          <div className="success-message">
            ✅ Script syntax is valid
          </div>
        )}
        
        {validationResult && validationResult.warnings.length > 0 && (
          <div className="warning-message">
            <strong>⚠️ Warnings:</strong>
            <ul style={{ marginTop: '10px', marginLeft: '20px' }}>
              {validationResult.warnings.map((warning, index) => (
                <li key={index}>{warning}</li>
              ))}
            </ul>
          </div>
        )}
        
        <div className="example-script">
          <h4>💡 Example Scripts</h4>
          <p>Here are sample scripts to get you started:</p>
          
          <div style={{ marginBottom: '20px' }}>
            <h5 style={{ color: 'var(--text-color-dark)', marginBottom: '10px' }}>JavaScript-based ExcelScript:</h5>
            <pre>{exampleScript}</pre>
          </div>
          
          <div>
            <h5 style={{ color: 'var(--text-color-dark)', marginBottom: '10px' }}>Microsoft ExcelScript (Auto-converted):</h5>
            <pre>{excelScriptExample}</pre>
          </div>
        </div>
        
        {error && (
          <div className="error-message">
            ❌ {error}
          </div>
        )}
      </div>
    </div>
  );
};

export default ScriptInput;
