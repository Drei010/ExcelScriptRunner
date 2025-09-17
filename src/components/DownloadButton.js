import React, { useState } from 'react';
import { downloadExcel } from '../utils/excelParser';

const DownloadButton = ({ processedData, fileName, disabled }) => {
  const [downloadFormat, setDownloadFormat] = useState('xlsx');
  const [isDownloading, setIsDownloading] = useState(false);

  const handleDownload = async () => {
    if (!processedData || disabled) return;

    setIsDownloading(true);
    
    try {
      if (downloadFormat === 'xlsx') {
        // Use existing Excel download function
        const outputFileName = `processed_${fileName.replace(/\.[^/.]+$/, '')}.xlsx`;
        downloadExcel(processedData, outputFileName);
      } else if (downloadFormat === 'csv') {
        // Create CSV download
        const csvContent = generateCSV(processedData);
        downloadCSV(csvContent, `processed_${fileName.replace(/\.[^/.]+$/, '')}.csv`);
      }
    } catch (error) {
      console.error('Download error:', error);
      alert(`Download failed: ${error.message}`);
    } finally {
      setIsDownloading(false);
    }
  };

  const generateCSV = (data) => {
    const sheetNames = Object.keys(data.sheets);
    if (sheetNames.length === 0) return '';

    // For CSV, we'll use the first sheet or combine all sheets
    const firstSheet = data.sheets[sheetNames[0]];
    if (!firstSheet || firstSheet.length === 0) return '';

    return firstSheet.map(row => 
      row.map(cell => {
        // Escape CSV values
        const cellValue = cell !== null && cell !== undefined ? String(cell) : '';
        if (cellValue.includes(',') || cellValue.includes('"') || cellValue.includes('\n')) {
          return `"${cellValue.replace(/"/g, '""')}"`;
        }
        return cellValue;
      }).join(',')
    ).join('\n');
  };

  const downloadCSV = (csvContent, fileName) => {
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = window.URL.createObjectURL(blob);
    
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    window.URL.revokeObjectURL(url);
  };

  return (
    <div className="section">
      <h2>💾 Download Processed File</h2>
      
      <div style={{ 
        display: 'flex', 
        alignItems: 'center', 
        gap: '15px',
        flexWrap: 'wrap'
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <label style={{ 
            fontWeight: '500', 
            color: '#2c3e50',
            fontSize: '1rem'
          }}>
            Format:
          </label>
          <select
            value={downloadFormat}
            onChange={(e) => setDownloadFormat(e.target.value)}
            style={{
              padding: '8px 12px',
              border: '2px solid #e1e8ed',
              borderRadius: '6px',
              fontSize: '14px',
              backgroundColor: 'white',
              color: '#2c3e50',
              cursor: 'pointer'
            }}
            disabled={disabled}
          >
            <option value="xlsx">Excel (.xlsx)</option>
            <option value="csv">CSV (.csv)</option>
          </select>
        </div>

        <button
          onClick={handleDownload}
          disabled={disabled || isDownloading}
          style={{
            background: disabled ? '#bdc3c7' : 'linear-gradient(135deg, #27ae60, #2ecc71)',
            color: 'white',
            border: 'none',
            padding: '12px 24px',
            fontSize: '1rem',
            fontWeight: '600',
            borderRadius: '8px',
            cursor: disabled ? 'not-allowed' : 'pointer',
            transition: 'all 0.3s ease',
            boxShadow: disabled ? 'none' : '0 4px 15px rgba(39, 174, 96, 0.3)',
            display: 'flex',
            alignItems: 'center',
            gap: '8px'
          }}
          onMouseEnter={(e) => {
            if (!disabled) {
              e.target.style.transform = 'translateY(-2px)';
              e.target.style.boxShadow = '0 6px 20px rgba(39, 174, 96, 0.4)';
            }
          }}
          onMouseLeave={(e) => {
            if (!disabled) {
              e.target.style.transform = 'translateY(0)';
              e.target.style.boxShadow = '0 4px 15px rgba(39, 174, 96, 0.3)';
            }
          }}
        >
          {isDownloading ? (
            <>
              <span className="loading"></span>
              Downloading...
            </>
          ) : (
            <>
              📥 Download Processed File
            </>
          )}
        </button>
      </div>

      {disabled && (
        <div style={{ 
          marginTop: '15px', 
          color: '#95a5a6', 
          fontSize: '0.9rem',
          fontStyle: 'italic'
        }}>
          Process a file first to enable download
        </div>
      )}

      <div style={{ 
        marginTop: '20px',
        padding: '15px',
        backgroundColor: '#f8f9fa',
        borderRadius: '8px',
        borderLeft: '4px solid #17a2b8'
      }}>
        <h4 style={{ color: '#2c3e50', marginBottom: '10px' }}>📋 Download Options</h4>
        <ul style={{ color: '#6c757d', lineHeight: '1.6', margin: 0 }}>
          <li><strong>Excel (.xlsx):</strong> Preserves formatting, multiple sheets, and data types</li>
          <li><strong>CSV (.csv):</strong> Plain text format, compatible with most applications</li>
        </ul>
      </div>
    </div>
  );
};

export default DownloadButton;
