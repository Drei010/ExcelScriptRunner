import React, { useState } from 'react';
import FileUpload from './components/FileUpload';
import ScriptInput from './components/ScriptInput';
import ProcessButton from './components/ProcessButton';
import DataPreview from './components/DataPreview';
import DownloadButton from './components/DownloadButton';
import { parseFile } from './utils/excelParser';

function App() {
  const [selectedFile, setSelectedFile] = useState(null);
  const [originalData, setOriginalData] = useState(null);
  const [originalFileName, setOriginalFileName] = useState('');
  const [script, setScript] = useState('');
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const [processedData, setProcessedData] = useState(null);
  const [processedFileName, setProcessedFileName] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);

  const handleFileSelect = async (file, errorMessage) => {
    setSelectedFile(file);
    setError(errorMessage || '');
    setSuccess('');
    setProcessedData(null);
    setProcessedFileName('');
    setOriginalData(null);
    setOriginalFileName('');
    if (file && !errorMessage) {
      try {
        setIsProcessing(true);
        const parsed = await parseFile(file);
        setOriginalData(parsed);
        setOriginalFileName(file.name);
      } catch (e) {
        setOriginalData(null);
        setOriginalFileName('');
        setError('Failed to parse file for preview: ' + e.message);
      } finally {
        setIsProcessing(false);
      }
    }
  };

  const handleScriptChange = (newScript) => {
    setScript(newScript);
    setError('');
    setSuccess('');
    setProcessedData(null);
    setProcessedFileName('');
  };

  const handleError = (errorMessage) => {
    setError(errorMessage);
    setSuccess('');
    setProcessedData(null);
    setProcessedFileName('');
  };

  const handleSuccess = (successMessage) => {
    setSuccess(successMessage);
    setError('');
  };

  const handleProcessedData = (data, fileName) => {
    setProcessedData(data);
    setProcessedFileName(fileName);
  };

  return (
    <div className="app">
      <header className="header">
        <h1>ExcelScript Runner</h1>
        <p>Process Excel and CSV files with custom scripts - No server required!</p>
      </header>

      <div className="container">
        <FileUpload
          onFileSelect={handleFileSelect}
          selectedFile={selectedFile}
          error={error}
        />
        
        <ScriptInput
          onScriptChange={handleScriptChange}
          script={script}
          error={error}
        />
      </div>

      {error && (
        <div className="error-message">
          ❌ {error}
        </div>
      )}

      {success && (
        <div className="success-message">
          ✅ {success}
        </div>
      )}

      <ProcessButton
        selectedFile={selectedFile}
        script={script}
        onError={handleError}
        onSuccess={handleSuccess}
        onProcessedData={handleProcessedData}
      />

      <div className="results-container">
        <DataPreview
          label="Before Processing"
          data={originalData}
          error={null}
          isLoading={isProcessing && !processedData}
        />
        <DataPreview
          label="After Processing"
          data={processedData}
          error={error && processedData == null ? error : null}
          isLoading={isProcessing && !!processedData}
        />
      </div>

      <DownloadButton
        processedData={processedData}
        fileName={processedFileName}
        disabled={!processedData}
      />

      <div style={{ 
        marginTop: '50px', 
        padding: '20px', 
        backgroundColor: '#f8f9fa', 
        borderRadius: '8px',
        borderLeft: '4px solid #17a2b8'
      }}>
        <h3 style={{ color: '#2c3e50', marginBottom: '15px' }}>📚 How to Use</h3>
        <ol style={{ color: '#6c757d', lineHeight: '1.6' }}>
          <li><strong>Upload a file:</strong> Click or drag an Excel (.xlsx, .xls) or CSV (.csv) file</li>
          <li><strong>Write your script:</strong> Enter ExcelScript commands to process your data</li>
          <li><strong>Process and download:</strong> Click the button to execute your script and download the result</li>
        </ol>
        
        <h4 style={{ color: '#2c3e50', marginTop: '20px', marginBottom: '10px' }}>🔧 Available Functions</h4>
        <ul style={{ color: '#6c757d', lineHeight: '1.6' }}>
          <li><code>getSheetNames()</code> - Get all sheet names</li>
          <li><code>getSheetData(sheetName)</code> - Get data from a specific sheet</li>
          <li><code>setSheetData(sheetName, data)</code> - Update sheet data</li>
          <li><code>getCell(sheetName, row, col)</code> - Get a specific cell value</li>
          <li><code>setCell(sheetName, row, col, value)</code> - Set a specific cell value</li>
          <li><code>getRow(sheetName, rowIndex)</code> - Get an entire row</li>
          <li><code>setRow(sheetName, rowIndex, rowData)</code> - Set an entire row</li>
          <li><code>getColumn(sheetName, colIndex)</code> - Get an entire column</li>
          <li><code>setColumn(sheetName, colIndex, columnData)</code> - Set an entire column</li>
          <li><code>addSheet(sheetName, data)</code> - Add a new sheet</li>
          <li><code>removeSheet(sheetName)</code> - Remove a sheet</li>
          <li><code>sum(...values)</code> - Sum numbers</li>
          <li><code>avg(...values)</code> - Calculate average</li>
          <li><code>min(...values)</code> - Find minimum value</li>
          <li><code>max(...values)</code> - Find maximum value</li>
          <li><code>upper(str)</code> - Convert to uppercase</li>
          <li><code>lower(str)</code> - Convert to lowercase</li>
          <li><code>trim(str)</code> - Remove whitespace</li>
          <li><code>concat(...values)</code> - Concatenate strings</li>
          <li><code>now()</code> - Get current date/time</li>
          <li><code>today()</code> - Get today's date</li>
        </ul>
      </div>
    </div>
  );
}

export default App;
