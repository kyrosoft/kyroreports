# Kyro Reports

A client-side web application for creating Excel templates and compiling multiple reports into a single consolidated document. No backend server required – all processing happens in your browser.

## Features

- **Template Creator**: Define custom column names and generate formatted .xlsx templates with styled headers
- **Report Compiler**: Upload a template and multiple reports (via ZIP), validate them, and compile into a single Excel file
- **Validation**: Automatic validation ensures all reports match the template structure (sheet names and columns)
- **Progress Tracking**: Visual progress bar during compilation
- **Toast Notifications**: Real-time feedback for success and error states
- **Auto-numbering**: Compiled reports include sequential row numbers for easy reference

## How It Works

Kyro Reports runs entirely in the browser using:
- `xlsx-js-style` – Excel file generation and manipulation
- `JSZip` – ZIP archive handling for bulk report uploads

### Template Creator

1. Enter column names separated by commas (e.g., `Name, Email, Phone, Department`)
2. Preview the table structure in real-time
3. Click **DOWNLOAD TEMPLATE XLSX** to generate a formatted Excel file
4. Distribute the template to users for data collection

### Report Compiler

1. Upload the original template .xlsx file
2. Upload a ZIP archive containing multiple filled report files (.xlsx)
3. The app validates each report against the template:
   - Checks sheet names match
   - Verifies column headers are identical
4. Click **DOWNLOAD COMPILED XLSX** to merge all reports into one file with:
   - Original headers with "No." as the first column
   - Sequential row numbers for each data row
   - Styled header row (bold, centered)

## Installation

### Using npm

```bash
# Install dependencies
npm install

# Run development server
npm run dev

# Build for production
npm run build
```

### Standalone Usage

Simply open `index.html` in a modern web browser. No build process required.

## Browser Compatibility

Requires a modern browser with ES6+ support:
- Chrome/Edge 90+
- Firefox 88+
- Safari 14+

## License

See [LICENSE](LICENSE) file for details.
