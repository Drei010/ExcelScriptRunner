import React, { useState } from 'react';
import { parseFile } from '../utils/excelParser';
import { executeScript } from '../utils/scriptEngine';
import { convertExcelScript, isExcelScript } from '../utils/scriptConverter';

const ProcessButton = ({ selectedFile, script, onError, onSuccess, onProcessedData }) => {
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState('');

  const handleProcess = async () => {
    if (!selectedFile) {
      onError('Please select a file to process');
      return;
    }

    if (!script || script.trim().length === 0) {
      onError('Please enter an ExcelScript');
      return;
    }

    setIsProcessing(true);
    setProgress('Reading file...');

    try {
      // Parse the uploaded file
      setProgress('Parsing file...');
      const parsedData = await parseFile(selectedFile);
      
      // Execute the ExcelScript
      setProgress('Executing script...');
      
      // Convert ExcelScript syntax if needed
      let scriptToExecute = script;
      if (isExcelScript(script)) {
        setProgress('Converting ExcelScript syntax...');
        scriptToExecute = convertExcelScript(script);
      }
      
      const processedData = executeScript(parsedData, scriptToExecute);
      
      setProgress('Complete!');
      onSuccess('File processed successfully! Check the preview below.');
      onProcessedData(processedData, selectedFile.name);
      
    } catch (error) {
      onError(`Processing failed: ${error.message}`);
      onProcessedData(null, null);
    } finally {
      setIsProcessing(false);
      setProgress('');
    }
  };

  const isDisabled = !selectedFile || !script || !script.trim() || isProcessing;

  return (
    <div style={{ textAlign: 'center', marginTop: '30px' }}>
      <button
        className="process-button"
        onClick={handleProcess}
        disabled={isDisabled}
      >
        {isProcessing ? (
          <>
            <span className="loading"></span>
            Processing...
          </>
        ) : (
          '🚀 Process File'
        )}
      </button>
      
      {isProcessing && progress && (
        <div style={{ 
          marginTop: '15px', 
          color: '#3498db', 
          fontWeight: '500' 
        }}>
          {progress}
        </div>
      )}
      
      {isDisabled && !isProcessing && (
        <div style={{ 
          marginTop: '15px', 
          color: '#95a5a6', 
          fontSize: '0.9rem' 
        }}>
          {!selectedFile && !script ? 'Please upload a file and enter a script' :
           !selectedFile ? 'Please upload a file' :
           !script ? 'Please enter an ExcelScript' : ''}
        </div>
      )}
    </div>
  );
};

export default ProcessButton;
