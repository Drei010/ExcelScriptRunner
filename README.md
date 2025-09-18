npm# ExcelScript Runner

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

### Advanced ExcelScript Features
-   `workbook.getWorksheet(name)`: Get a specific worksheet by name.
-   `workbook.getActiveWorksheet()`: Get the currently active worksheet.
-   `sheet.getRange(range).getValues()`: Get values from a range as a 2D array.
-   `sheet.getRange(range).setValues(values)`: Set multiple values at once using a 2D array.
-   `sheet.getRange(range).removeDuplicates(columns, hasHeaders)`: Remove duplicate rows based on specified columns.
-   `sheet.getRange(range).autoFill(destinationRange, fillType)`: Auto-fill data patterns to a destination range.
-   `sheet.getRange(range).clear(applyTo)`: Clear range contents, formats, or both.
-   `sheet.getAutoFilter().getRange().getSort().apply(keys, matchCase, hasHeaders)`: Sort data using autofilter functionality.

### ExcelScript Constants
-   `ExcelScript.AutoFillType`: Constants for auto-fill operations (fillDefault, fillCopy, fillSeries, etc.).
-   `ExcelScript.ClearApplyTo`: Constants for clear operations (all, contents, formats, hyperlinks).
-   `ExcelScript.HorizontalAlignment`: Text alignment constants (general, left, center, right).
-   `ExcelScript.DeleteShiftDirection`: Row/column deletion direction (up, left).

## Supported ExcelScript Syntax

The application features an auto-conversion engine that allows you to write scripts using syntax very similar to Microsoft ExcelScript (Office Scripts). This includes:

-   **Function Signature**: `function main(workbook: ExcelScript.Workbook) { ... }`
-   **Workbook/Worksheet Access**: `workbook.getActiveWorksheet()`, `selectedSheet` variables.
-   **Range References**: `selectedSheet.getRange("A1")`, `selectedSheet.getRange("B:B")`.
-   **Range Methods**: `getTexts()`, `getValues()`, `setValue()`, `setValues()`, `setNumberFormatLocal()`.
-   **Advanced Operations**: `removeDuplicates()`, `autoFill()`, `clear()`.
-   **Sorting**: `getAutoFilter().getRange().getSort().apply()`.
-   **Row/Column Manipulation**: `getEntireRow().delete()`.
-   **Table Management**: `workbook.addTable()`, `newTable.setPredefinedTableStyle()`.
-   **Constants**: `ExcelScript.HorizontalAlignment`, `ExcelScript.DeleteShiftDirection`, `ExcelScript.AutoFillType`, `ExcelScript.ClearApplyTo`.

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

### Large Script Support

The application now supports complex ExcelScripts with multiple functions and advanced operations. Here's an example of a large script that demonstrates the full capabilities:

```typescript
function main(workbook: ExcelScript.Workbook) {
    // Get worksheets with early exit if not found
    const ticketExportSheet = workbook.getWorksheet("Ticket Export");
    const ticketDetailsSheet = workbook.getWorksheet("Ticket Details");
    
    if (!ticketExportSheet || !ticketDetailsSheet) return;
    
    // Process Ticket Export sheet
    processTicketExportSheet(ticketExportSheet);
    
    // Compare tickets between sheets
    compareAndUpdateTickets(ticketExportSheet, ticketDetailsSheet);
    
    // Populate the rows of Ticket Details
    populateTicketDetailsRows(ticketDetailsSheet);
    
    // Update internal status of Ticket Details
    updateInternalStatusRows(ticketDetailsSheet);
}

function processTicketExportSheet(sheet: ExcelScript.Worksheet) {
    const lastRowExportSheet = sheet.getRange("A1").getSurroundingRegion().getLastRow().getRowIndex();
    
    // Custom sort on range spanned by auto filter (descending on column S)
    sheet.getAutoFilter().getRange().getSort().apply([{ key: 19, ascending: false }], false, true);
    
    // Remove duplicates from the entire data range (A2:T...)
    sheet.getRange(`A2:T${lastRowExportSheet + 1}`).removeDuplicates([0], true);
}

function compareAndUpdateTickets(exportSheet: ExcelScript.Worksheet, detailsSheet: ExcelScript.Worksheet) {
    const lastRowExportSheet = exportSheet.getRange("A1").getSurroundingRegion().getLastRow().getRowIndex() + 1;
    const lastRowDetailsSheet = detailsSheet.getRange("C1").getSurroundingRegion().getLastRow().getRowIndex() + 2;
    const exportTickets = getValidTickets(exportSheet, `A2:A${lastRowExportSheet}`);
    const detailsTickets = getValidTickets(detailsSheet, `C3:C${lastRowDetailsSheet}`);
    
    if (exportTickets.length === 0) return;
    
    // Find missing tickets using Set for O(1) lookups
    const detailsSet = new Set(detailsTickets);
    const missingTickets = exportTickets.filter(ticket => !detailsSet.has(ticket));
    
    if (missingTickets.length === 0) return;
    
    // Calculate target range for adding missing tickets
    const startRow = lastRowDetailsSheet;
    const availableRows = 5000 - startRow; // Assuming worksheet has 5000 rows max
    const ticketsToAdd = missingTickets.slice(0, availableRows);
    
    if (ticketsToAdd.length > 0) {
        const targetRange = detailsSheet.getRange(`C${startRow}:C${startRow + ticketsToAdd.length - 1}`);
        targetRange.setValues(ticketsToAdd.map(ticket => [ticket]));
    }
}

function getValidTickets(sheet: ExcelScript.Worksheet, rangeAddress: string): string[] {
    const values = sheet.getRange(rangeAddress).getValues();
    return values
        .filter(row => {
            const value = row[0];
            if (!value) return false;
            const strValue = String(value).toUpperCase();
            return strValue.startsWith("RIT") || strValue.startsWith("INC");
        })
        .map(row => String(row[0]));
}

function populateTicketDetailsRows(sheet: ExcelScript.Worksheet) {
    const startRow = sheet.getRange("C1").getSurroundingRegion().getLastRow().getRowIndex() + 1;
    const columnDValues = sheet.getRange(`A1:A${startRow}`).getValues();
    const endRow = columnDValues.filter(row => row[0] !== "").length;
    if (startRow == endRow) return;
    
    // Perform autofill for specified column ranges
    autoFillRange(sheet, "A", "B", startRow, endRow);
    autoFillRange(sheet, "D", "L", startRow, endRow);
    autoFillRange(sheet, "U", "AA", startRow, endRow);
    autoFillRange(sheet, "AD", "AD", startRow, endRow);
    autoFillRange(sheet, "AF", "AH", startRow, endRow);
    autoFillRange(sheet, "AJ", "AM", startRow, endRow);
    autoFillRange(sheet, "AP", "AP", startRow, endRow);
    
    // Copy values from Z to AB where AB is empty
    copyZtoABIfEmpty(sheet, startRow, endRow);
}

function autoFillRange(sheet: ExcelScript.Worksheet, startCol: string, endCol: string, startRow: number, endRow: number) {
    const rangeToAutoFill = sheet.getRange(`${startCol}${endRow}:${endCol}${endRow}`);
    const fillToRange = sheet.getRange(`${startCol}${endRow}:${endCol}${endRow + (startRow - endRow)}`);
    rangeToAutoFill.autoFill(fillToRange, ExcelScript.AutoFillType.fillDefault);
}

function copyZtoABIfEmpty(sheet: ExcelScript.Worksheet, startRow: number, endRow: number) {
    // Get the range for columns Z and AB for the specified rows
    const zRange = sheet.getRange(`Z${startRow}:Z${endRow}`);
    const abRange = sheet.getRange(`AB${startRow}:AB${endRow}`);
    
    // Get the values from the ranges
    const zValues = zRange.getValues();
    const abValues = abRange.getValues();
    
    // Loop through the values and update column AB if it's empty
    for (let i = 0; i < zValues.length; i++) {
        if (!abValues[i][0]) {
            abValues[i][0] = zValues[i][0];
        }
    }
    abRange.setValues(abValues);
}

function updateInternalStatusRows(sheet: ExcelScript.Worksheet) {
    const endRow = sheet.getRange("C1").getSurroundingRegion().getLastRow().getRowIndex();
    
    // Define the range for processing (last 1000 rows)
    const startRow = Math.max(1, endRow - 2000); // Ensure startRow doesn't go below 1
    const rowCount = endRow - startRow + 1;
    
    // Define status mapping
    const statusMap: Record<string, string> = {
        "Closed Complete": "Closed",
        "Closed": "Closed",
        "Closed Incomplete": "Closed",
        "Closed Skipped": "Closed",
        "Resolved": "Closed",
        "Pending": "On Hold",
        "On Hold": "On Hold",
        "Cancelled": "Cancelled",
        "In Progress": "In Progress",
        "Work in Progress": "In Progress"
    };
    
    // Read column L, N, M, and O values in bulk
    const columnLValues = sheet.getRange(`L${startRow}:L${endRow}`).getValues() as string[][];
    const columnNValues = sheet.getRange(`N${startRow}:N${endRow}`).getValues() as string[][];
    const columnMValues = sheet.getRange(`M${startRow}:M${endRow}`).getValues() as string[][];
    const columnOValues = sheet.getRange(`O${startRow}:O${endRow}`).getValues() as string[][];
    
    // Process and update values for columns M, N, and O
    for (let i = 0; i < rowCount; i++) {
        const lValue = columnLValues[i][0];
        const nValue = columnNValues[i][0];
        const resolvedValue = "Resolved";
        
        // Skip rows where column N is not blank, keep the existing value in column M
        if (nValue.trim() !== "") {
            continue; // Skip this row
        } else {
            // If column N is blank, map the status from column L to column M
            columnMValues[i][0] = statusMap[lValue] || ""; // Update column M with mapped value
        }
        
        // Additional condition: If M is "Closed" and N is blank, insert "Resolved" into N and O
        if ((columnMValues[i][0] === "Closed" || columnMValues[i][0] === "Cancelled") && nValue.trim() === "") {
            columnNValues[i][0] = resolvedValue;
            columnOValues[i][0] = resolvedValue;
        }
    }
    
    // Write the updated column M, N, and O values back in bulk
    sheet.getRange(`M${startRow}:M${endRow}`).setValues(columnMValues);
    sheet.getRange(`N${startRow}:N${endRow}`).setValues(columnNValues);
    sheet.getRange(`O${startRow}:O${endRow}`).setValues(columnOValues);
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
