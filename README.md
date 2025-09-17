# ExcelScript Runner

![ExcelScript Runner Screenshot](screenshot.png)

A static React web application that processes Excel and CSV files using custom ExcelScript logic. It runs entirely in the browser, requiring no backend server, making it ideal for static hosting solutions like GitHub Pages.

## Table of Contents

- [Features](#features)
- [Getting Started](#getting-started)
- [Deployment to GitHub Pages](#deployment-to-github-pages)
- [Usage](#usage)
- [ExcelScript Functions](#excelscript-functions)
- [Supported ExcelScript Syntax](#supported-excelscript-syntax)
- [Technical Details](#technical-details)
- [Security](#security)
- [Browser Support](#browser-support)
- [License](#license)

## Features

- 📁 **File Upload**: Supports Excel (.xlsx, .xls) and CSV (.csv) files with drag-and-drop functionality and file validation.
- 📝 **ExcelScript Input**: A dedicated text area for writing or pasting custom ExcelScript logic. Includes real-time syntax validation, automatic conversion for Microsoft ExcelScript, and helpful example scripts.
- 📊 **Dual Preview**: Visually compare your data *before* and *after* processing with side-by-side table previews. Supports partial previews for large files.
- 💾 **Flexible Download**: Download your processed file in either Excel (.xlsx) or CSV (.csv) format.
- ✅ **Robust Error Handling**: Clear messages for missing files/scripts, unsupported formats, and script execution errors.
- 🎨 **Modern & Responsive UI**: A clean, intuitive design that adapts seamlessly to desktop, tablet, and mobile devices, ensuring optimal readability and usability.

## Getting Started

### Prerequisites

- Node.js (version 14 or higher)
- npm or yarn
- Git (for cloning and GitHub Pages deployment)

### Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/<YOUR_GITHUB_USERNAME>/excelscript-runner.git
    cd excelscript-runner
    ```
    *Remember to replace `<YOUR_GITHUB_USERNAME>` with your actual GitHub username.*

2.  **Install dependencies:**
    ```bash
    npm install
    ```

### Running Locally

To start the development server:

```bash
npm start
```

This will open the application in your browser at [http://localhost:3000](http://localhost:3000).

### Building for Production

To create a production-ready build:

```bash
npm run build
```

This command compiles the React application into static files and places them in the `build` folder.

## Deployment to GitHub Pages

This application is configured for easy deployment to GitHub Pages using the `gh-pages` package.

1.  **Ensure `homepage` is set in `package.json`:**
    Open `package.json` and verify that the `homepage` field is correctly set to your GitHub Pages URL:
    ```json
    "homepage": "http://<YOUR_GITHUB_USERNAME>.github.io/excelscript-runner",
    ```
    *Replace `<YOUR_GITHUB_USERNAME>` with your actual GitHub username.*

2.  **Deploy the application:**
    ```bash
    npm run deploy
    ```

    This command will:
    -   Run `npm run build` to create a production build.
    -   Push the contents of the `build` directory to the `gh-pages` branch of your repository.

    After deployment, your application will be accessible at the `homepage` URL you specified (e.g., `http://<YOUR_GITHUB_USERNAME>.github.io/excelscript-runner`). It might take a few minutes for GitHub Pages to update.

## Usage

1.  **Upload a File**: Click or drag an Excel (.xlsx, .xls) or CSV (.csv) file to the left-side upload area. A preview of your original data will appear.
2.  **Write ExcelScript**: Enter your processing logic in the script input area. You can use either JavaScript-based ExcelScript or Microsoft ExcelScript syntax (which will be auto-converted).
3.  **Process**: Click the "Process File" button. The script will execute, and a preview of the processed data will appear on the right side.
4.  **Compare & Download**: Review the "Before Processing" and "After Processing" previews. Select your desired download format (Excel or CSV) and click "Download Processed File" to get the full dataset.

## ExcelScript Functions

The application provides a rich set of built-in functions for data manipulation, compatible with both JavaScript and auto-converted Microsoft ExcelScript:

### Data Access
-   `getSheetNames()`: Returns an array of sheet names in the workbook.
-   `getSheetData(sheetName)`: Retrieves all data from a specified sheet as a 2D array.
-   `setSheetData(sheetName, data)`: Overwrites a sheet's data with a new 2D array.
-   `getCell(sheetName, row, col)`: Gets the value of a specific cell (0-indexed row/column).
-   `setCell(sheetName, row, col, value)`: Sets the value of a specific cell.

### Row/Column Operations
-   `getRow(sheetName, rowIndex)`: Retrieves a specific row's data.
-   `setRow(sheetName, rowIndex, rowData)`: Sets a specific row's data.
-   `getColumn(sheetName, colIndex)`: Retrieves a specific column's data.
-   `setColumn(sheetName, colIndex, columnData)`: Sets a specific column's data.

### Sheet Management
-   `addSheet(sheetName, data = [])`: Adds a new sheet with optional initial data.
-   `removeSheet(sheetName)`: Deletes a specified sheet.

### Math Functions
-   `sum(...values)`: Calculates the sum of all provided numeric values.
-   `avg(...values)`: Calculates the average of all provided numeric values.
-   `min(...values)`: Finds the minimum value among provided numeric values.
-   `max(...values)`: Finds the maximum value among provided numeric values.

### String Functions
-   `upper(str)`: Converts a string to uppercase.
-   `lower(str)`: Converts a string to lowercase.
-   `trim(str)`: Removes whitespace from both ends of a string.
-   `concat(...values)`: Concatenates multiple values into a single string.

### Date Functions
-   `now()`: Returns the current `Date` object.
-   `today()`: Returns today's date in `YYYY-MM-DD` format.
-   `newDate()`: Creates a new `Date` object (compatible with Microsoft ExcelScript's `new Date()`).

## Supported ExcelScript Syntax

The application features an auto-conversion engine that allows you to write scripts using syntax very similar to Microsoft ExcelScript (Office Scripts). This includes:

-   **Function Signature**: `function main(workbook: ExcelScript.Workbook) { ... }`
-   **Workbook/Worksheet Access**: `workbook.getActiveWorksheet()`, `selectedSheet` variables.
-   **Range References**: `selectedSheet.getRange("A1")`, `selectedSheet.getRange("B:B")`.
-   **Range Methods**: `getTexts()`, `getValues()`, `setValue()`, `setNumberFormatLocal()`.
-   **Row/Column Manipulation**: `getEntireRow().delete()`.
-   **Table Management**: `workbook.addTable()`, `newTable.setPredefinedTableStyle()`.
-   **Constants**: `ExcelScript.HorizontalAlignment`, `ExcelScript.DeleteShiftDirection`.

### Example Microsoft ExcelScript (Auto-converted)

```typescript
function main(workbook: ExcelScript.Workbook) {
    const selectedSheet = workbook.getActiveWorksheet();

    // Remove existing autofilter if present
    selectedSheet.getAutoFilter()?.remove();

    // Get current date
    const currentDate = new Date();
    const month = currentDate.getMonth() + 1;
    const day = currentDate.getDate();
    const year = currentDate.getFullYear();
    const hours = currentDate.getHours();
    const minutes = currentDate.getMinutes();
    const seconds = currentDate.getSeconds();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const currentMonthYear = monthNames[currentDate.getMonth()] + "-" + currentDate.getFullYear().toString().slice(-2);

    const dateTimeString = `${month}/${day}/${year} ${hours % 12 || 12}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')} ${ampm}`;

    // Get the entire used range at once
    const usedRange = selectedSheet.getUsedRange();
    if (!usedRange) {
        console.log("No data found in the sheet.");
        return;
    }

    // Format the entire range at once
    const format = usedRange.getFormat();
    format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.general);
    format.setIndentLevel(0);
    format.setWrapText(false);
    format.setTextOrientation(0);

    // Process rows to delete in batch
    const rowCount = usedRange.getRowCount();
    const columnOValues = selectedSheet.getRange(`O1:O${rowCount}`).getTexts();
    const rowsToDelete: number[] = [];

    for (let i = 1; i < rowCount; i++) {
        const cellValue = columnOValues[i][0];
        if (cellValue !== "" && cellValue !== "Closed" && cellValue.substring(3, 9) !== currentMonthYear) {
            rowsToDelete.push(i + 1);
        }
    }

    // Delete rows in reverse order
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        selectedSheet.getRange(`O${rowsToDelete[i]}`).getEntireRow().delete(ExcelScript.DeleteShiftDirection.up);
    }

    // Format columns and update dates
    selectedSheet.getRange("M:M").setNumberFormatLocal("m/d/yyyy");
    selectedSheet.getRange("T:T").setNumberFormatLocal("m/d/yyyy");

    // Process date conversion in batch
    const lastRowA = selectedSheet.getRange("A1").getSurroundingRegion().getLastRow().getRowIndex();
    const columnAValues = selectedSheet.getRange(`A1:A${lastRowA}`).getValues();
    let lastRow = columnAValues.filter(row => String(row[0]).startsWith("INC") || String(row[0]).startsWith("RIT")).length;

    if (lastRow === 0) {
        console.log("No data found in the sheet.");
        return;
    }

    // Update column T with current date/time in batch
    const targetRange = selectedSheet.getRange(`T2:T${lastRow + 2}`);
    targetRange.setValue(dateTimeString);

    // Remove autofilter before creating table
    if (selectedSheet.getAutoFilter() !== null) {
        selectedSheet.getAutoFilter().remove();
    }

    // Create table with the final range
    const newTable = workbook.addTable(selectedSheet.getRange(`A1:T${lastRow + 2}`), true);
    newTable.setPredefinedTableStyle("TableStyleLight1");
}
```

## Technical Details

-   **Frontend**: React 18 (CRA)
-   **File Processing**: SheetJS (xlsx library)
-   **Script Execution**: Custom JavaScript interpreter with a sandboxed context, enhanced for Microsoft ExcelScript compatibility.
-   **Styling**: Custom CSS with a modern design system and full responsiveness.
-   **Build & Deployment**: `react-scripts` for building and `gh-pages` for seamless GitHub Pages deployment.

## Security

The ExcelScript execution engine runs in a sandboxed browser environment that:
-   Prevents direct access to sensitive browser APIs (e.g., `document`, `window`).
-   Blocks potentially dangerous JavaScript operations (`eval`, `Function` constructor, `setTimeout`, `setInterval`).
-   Validates script syntax before execution to catch common issues.
-   Uses deep cloning for spreadsheet data to prevent unintended modifications to the original dataset.

## Browser Support

Optimized and tested for modern web browsers:
-   Google Chrome (recommended)
-   Mozilla Firefox
-   Apple Safari
-   Microsoft Edge

## License

This project is open source and available under the MIT License.
