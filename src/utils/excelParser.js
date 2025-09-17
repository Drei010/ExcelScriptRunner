import * as XLSX from 'xlsx';

/**
 * Parse an Excel or CSV file and return the data
 * @param {File} file - The uploaded file
 * @returns {Promise<Object>} - Parsed data with sheets and metadata
 */
export const parseFile = async (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        
        const result = {
          fileName: file.name,
          fileType: file.type,
          fileSize: file.size,
          sheets: {},
          sheetNames: workbook.SheetNames
        };
        
        // Parse each sheet
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          result.sheets[sheetName] = jsonData;
        });
        
        resolve(result);
      } catch (error) {
        reject(new Error(`Failed to parse file: ${error.message}`));
      }
    };
    
    reader.onerror = () => {
      reject(new Error('Failed to read file'));
    };
    
    reader.readAsBinaryString(file);
  });
};

/**
 * Convert processed data back to Excel format and trigger download
 * @param {Object} data - Processed data object
 * @param {string} fileName - Output file name
 */
export const downloadExcel = (data, fileName = 'processed_data.xlsx') => {
  try {
    const workbook = XLSX.utils.book_new();
    
    // Add each sheet to the workbook
    Object.keys(data.sheets).forEach(sheetName => {
      const worksheet = XLSX.utils.aoa_to_sheet(data.sheets[sheetName]);
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    });
    
    // Generate Excel file
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    
    // Create blob and trigger download
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    window.URL.revokeObjectURL(url);
  } catch (error) {
    throw new Error(`Failed to generate Excel file: ${error.message}`);
  }
};

/**
 * Validate file type
 * @param {File} file - The file to validate
 * @returns {boolean} - Whether the file type is supported
 */
export const validateFileType = (file) => {
  const supportedTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
    'application/vnd.ms-excel', // .xls
    'text/csv', // .csv
    'application/csv' // .csv alternative
  ];
  
  const supportedExtensions = ['.xlsx', '.xls', '.csv'];
  const fileName = file.name.toLowerCase();
  
  return supportedTypes.includes(file.type) || 
         supportedExtensions.some(ext => fileName.endsWith(ext));
};
