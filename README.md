# SheetExporter

A Google Apps Script library for exporting Google Sheets to PDF, Excel, CSV, and other formats.

## Features

- Export to PDF, Excel (XLSX), CSV, TSV, ODS, and ZIP
- Fluent chainable API
- Control page size, orientation, margins, and scaling
- Export specific sheets or ranges
- Preset configurations for common use cases

## Quick Start

```javascript
function exportToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const blob = new SheetExporter(ss)
    .setFormat('pdf')
    .setOrientation('landscape')
    .setPageSize('letter')
    .setFileName('Weekly Report')
    .exportAsBlob();

  DriveApp.createFile(blob);
}
```

## Documentation

For the complete 34-page guide with 25+ examples, best practices, and advanced usage:

**https://spreadsheet.dev/sheet-exporter**

## License

MIT License - Copyright (c) 2025 Spreadsheet Dev
