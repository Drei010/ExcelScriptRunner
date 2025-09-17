import React, { useState, useRef } from 'react';
import { validateFileType } from '../utils/excelParser';

const FileUpload = ({ onFileSelect, selectedFile, error }) => {
  const [isDragOver, setIsDragOver] = useState(false);
  const fileInputRef = useRef(null);

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    setIsDragOver(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragOver(false);
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleFileSelect(files[0]);
    }
  };

  const handleFileSelect = (file) => {
    if (validateFileType(file)) {
      onFileSelect(file);
    } else {
      onFileSelect(null, 'Unsupported file type. Please upload an Excel (.xlsx, .xls) or CSV (.csv) file.');
    }
  };

  const handleInputChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      handleFileSelect(file);
    }
  };

  const handleClick = () => {
    fileInputRef.current.click();
  };

  const formatFileSize = (bytes) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  return (
    <div className="section">
      <h2>📁 File Upload</h2>
      
      <div
        className={`file-upload-area ${isDragOver ? 'dragover' : ''}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={handleClick}
      >
        <div className="file-upload-icon">📄</div>
        <div className="file-upload-text">
          {selectedFile ? 'Click or drag to change file' : 'Click or drag to upload file'}
        </div>
        <div className="file-upload-subtext">
          Supports Excel (.xlsx, .xls) and CSV (.csv) files
        </div>
      </div>
      
      <input
        ref={fileInputRef}
        type="file"
        className="file-input"
        accept=".xlsx,.xls,.csv"
        onChange={handleInputChange}
      />
      
      {selectedFile && (
        <div className="file-info">
          <h4>📋 Selected File</h4>
          <p><strong>Name:</strong> {selectedFile.name}</p>
          <p><strong>Size:</strong> {formatFileSize(selectedFile.size)}</p>
          <p><strong>Type:</strong> {selectedFile.type || 'Unknown'}</p>
        </div>
      )}
      
      {error && (
        <div className="error-message">
          ❌ {error}
        </div>
      )}
    </div>
  );
};

export default FileUpload;
