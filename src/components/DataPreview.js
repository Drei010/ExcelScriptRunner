import React from 'react';

const DataPreview = ({ label, data, error, isLoading }) => {
  const renderTable = (sheetName, sheetData) => {
    if (!sheetData || sheetData.length === 0) {
      return (
        <div style={{ 
          padding: '20px', 
          textAlign: 'center', 
          color: '#7f8c8d',
          fontStyle: 'italic'
        }}>
          No data available
        </div>
      );
    }

    const maxRows = 20;
    const maxCols = 5;
    const displayData = sheetData.slice(0, maxRows);
    const isPartial = sheetData.length > maxRows;
    const hasManyColumns = displayData[0] && displayData[0].length > maxCols;

    return (
      <div style={{ marginBottom: '30px' }}>
        <h4 style={{ 
          color: '#2c3e50', 
          marginBottom: '15px',
          borderBottom: '2px solid #3498db',
          paddingBottom: '5px'
        }}>
          📊 Sheet: {sheetName}
          {(isPartial || hasManyColumns) && (
            <span style={{ 
              fontSize: '0.8em', 
              color: '#7f8c8d', 
              fontWeight: 'normal',
              marginLeft: '10px'
            }}>
              {isPartial && `(Showing first ${maxRows} rows of ${sheetData.length} total)`}
              {hasManyColumns && ` (Showing first ${maxCols} columns of ${displayData[0].length} total)`}
            </span>
          )}
        </h4>
        
        <div style={{ 
          overflow: 'auto', 
          border: '1px solid #e1e8ed', 
          borderRadius: '8px',
          maxHeight: '400px',
          maxWidth: '100%'
        }}>
          <table style={{ 
            width: '100%', 
            borderCollapse: 'collapse',
            fontSize: '14px'
          }}>
            <thead>
              <tr style={{ backgroundColor: '#f8f9fa' }}>
                {displayData[0] && displayData[0].slice(0, maxCols).map((header, index) => (
                  <th key={index} style={{ 
                    padding: '12px 8px', 
                    textAlign: 'left',
                    borderBottom: '2px solid #dee2e6',
                    fontWeight: '600',
                    color: '#2c3e50',
                    minWidth: '120px',
                    whiteSpace: 'nowrap',
                    overflow: 'hidden',
                    textOverflow: 'ellipsis'
                  }}>
                    {header || `Column ${index + 1}`}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {displayData.slice(1).map((row, rowIndex) => (
                <tr key={rowIndex} style={{ 
                  borderBottom: '1px solid #e1e8ed',
                  backgroundColor: rowIndex % 2 === 0 ? '#fff' : '#f8f9fa'
                }}>
                  {row.slice(0, maxCols).map((cell, cellIndex) => (
                    <td key={cellIndex} style={{ 
                      padding: '10px 8px',
                      borderRight: '1px solid #e1e8ed',
                      color: '#495057',
                      minWidth: '120px',
                      maxWidth: '200px',
                      wordWrap: 'break-word',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis'
                    }} title={cell !== null && cell !== undefined ? String(cell) : ''}>
                      {cell !== null && cell !== undefined ? String(cell) : ''}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        
        {(isPartial || hasManyColumns) && (
          <div style={{ 
            marginTop: '10px', 
            padding: '10px',
            backgroundColor: '#fff3cd',
            border: '1px solid #ffeaa7',
            borderRadius: '4px',
            color: '#856404',
            fontSize: '0.9em'
          }}>
            {isPartial && `⚠️ This is a partial preview. The downloaded file will contain all ${sheetData.length} rows.`}
            {hasManyColumns && (
              <div style={{ marginTop: isPartial ? '5px' : '0' }}>
                📊 This table has {displayData[0].length} columns. Only the first {maxCols} columns are shown in the preview. The downloaded file will contain all columns.
              </div>
            )}
          </div>
        )}
      </div>
    );
  };

  if (isLoading) {
    return (
      <div className="section">
        <h2>{label}</h2>
        <div style={{ 
          textAlign: 'center', 
          padding: '40px',
          color: '#7f8c8d'
        }}>
          <div className="loading" style={{ margin: '0 auto 15px' }}></div>
          Generating preview...
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="section">
        <h2>{label}</h2>
        <div className="error-message">
          ❌ Preview Error: {error}
        </div>
      </div>
    );
  }

  if (!data || !data.sheets) {
    return (
      <div className="section">
        <h2>{label}</h2>
        <div style={{ 
          textAlign: 'center', 
          padding: '40px',
          color: '#7f8c8d',
          fontStyle: 'italic'
        }}>
          No data to preview
        </div>
      </div>
    );
  }

  const sheetNames = Object.keys(data.sheets);

  return (
    <div className="section">
      <h2>{label}</h2>
      {sheetNames.length === 0 ? (
        <div style={{ 
          textAlign: 'center', 
          padding: '40px',
          color: '#7f8c8d',
          fontStyle: 'italic'
        }}>
          No data sheets found
        </div>
      ) : (
        <div>
          {sheetNames.map(sheetName => (
            <div key={sheetName}>
              {renderTable(sheetName, data.sheets[sheetName])}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default DataPreview;
