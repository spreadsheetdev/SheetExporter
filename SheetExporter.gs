/**
 * SheetExporter - A utility class for exporting Google Sheets spreadsheets to various formats.
 * 
 * Version 1.0
 *
 * Get the complete guide with examples and best practices at:
 * https://spreadsheet.dev/sheet-exporter 
 * 
 * Loved this script? Learn how to build complete, automated reporting systems
 * by joining the free newsletter at https://spreadsheet.dev/subscribe
 * 
 * MIT License
 * 
 * Copyright (c) 2025 Spreadsheet Dev
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */
class SheetExporter {
  
  /**
   * Creates a new SheetExporter instance
   * @param {Spreadsheet} spreadsheet - The Google Sheets Spreadsheet object to export
   * @throws {Error} If spreadsheet is not provided or invalid
   */
  constructor(spreadsheet) {
    if (!spreadsheet || typeof spreadsheet.getId !== 'function') {
      throw new Error('Valid Spreadsheet object is required');
    }
    
    this.spreadsheet = spreadsheet;
    this.params = {};
    this.fileName = 'export';
    
    // Internal properties for deferred timestamp calculation
    this._timestampDate = null;
    this._timezone = null;
    
    // Valid values for validation
    this.validFormats = ['pdf', 'csv', 'xls', 'xlsx', 'tsv', 'ods', 'zip'];
    this.validPageSizes = ['letter', 'tabloid', 'legal', 'statement', 'executive', 
                           'folio', 'a3', 'a4', 'a5', 'b4', 'b5'];
    this.validScales = [1, 2, 3, 4]; // 1=normal, 2=fit width, 3=fit height, 4=fit page
  }
  
  /**
   * Converts a JavaScript Date object to the Google Sheets "1900" serial number format.
   * @private
   * @param {Date} date The JavaScript Date object.
   * @returns {number} The spreadsheet serial number timestamp.
   */
  _jsDateToSpreadsheetTimestamp(date) {
    // (date.valueOf() / 86400000) is equivalent to (date.valueOf() / (24 * 60 * 60 * 1000))
    return (date.valueOf() / 86400000) + 25569;
  }
  
  setFormat(format) {
    if (!this.validFormats.includes(format)) {
      throw new Error(`Invalid format: ${format}. Valid formats: ${this.validFormats.join(', ')}`);
    }
    this.params.format = format;
    return this;
  }
  
  setOrientation(orientation) {
    if (orientation !== 'portrait' && orientation !== 'landscape') {
      throw new Error('Orientation must be "portrait" or "landscape"');
    }
    this.params.portrait = orientation === 'portrait' ? 'true' : 'false';
    return this;
  }
  
  setPageSize(size) {
    if (!this.validPageSizes.includes(size)) {
      throw new Error(`Invalid page size: ${size}. Valid sizes: ${this.validPageSizes.join(', ')}`);
    }
    this.params.size = size;
    return this;
  }
  
  setScale(scale) {
    const scaleMap = {
      'normal': 1,
      'fitWidth': 2,
      'fitHeight': 3,
      'fitPage': 4
    };
    
    let scaleValue = scale;
    if (typeof scale === 'string' && scaleMap[scale]) {
      scaleValue = scaleMap[scale];
    }
    
    if (!this.validScales.includes(Number(scaleValue))) {
      throw new Error('Scale must be 1 (normal), 2 (fit width), 3 (fit height), or 4 (fit page)');
    }
    this.params.scale = scaleValue;
    return this;
  }
  
  setSheetNames(val) {
    this.params.sheetnames = val ? 'true' : 'false';
    return this;
  }
  
  setSheetId(gid) {
    const numericGid = Number(gid);
    if (isNaN(numericGid)) {
      throw new Error('Sheet ID (gid) must be a number');
    }
    this.params.gid = numericGid;
    return this;
  }
  
  setSheetByName(sheetName) {
    try {
      const sheet = this.spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        throw new Error(`Sheet "${sheetName}" not found in spreadsheet`);
      }
      this.params.gid = sheet.getSheetId();
      return this;
    } catch (e) {
      throw new Error(`Error accessing sheet: ${e.message}`);
    }
  }
  
  setNotes(val) {
    this.params.printnotes = val ? 'true' : 'false';
    return this;
  }
  
  setTitle(val) {
    this.params.title = val ? 'true' : 'false';
    return this;
  }
  
  setGridlines(val) {
    this.params.gridlines = val ? 'true' : 'false';
    return this;
  }
  
  setPageNumbers(val) {
    if (val) {
      this.params.pagenum = 'CENTER';
    } else {
      delete this.params.pagenum;
    }
    return this;
  }
  
  setRepeatFrozenRows(repeat) {
    this.params.fzr = repeat ? 'true' : 'false';
    return this;
  }
  
  setRepeatFrozenColumns(repeat) {
    this.params.fzc = repeat ? 'true' : 'false';
    return this;
  }
  
  setRange(a1Notation) {
    if (!this.params.gid && this.params.gid !== 0) {
      throw new Error('You must set a sheet ID using setSheetId() or setSheetByName() before setting a range');
    }
    
    try {
      const sheets = this.spreadsheet.getSheets();
      const sheet = sheets.find(s => s.getSheetId() === this.params.gid);
      
      if (!sheet) {
        throw new Error('Sheet not found for the given gid');
      }
      
      const range = sheet.getRange(a1Notation);
      const startRow = range.getRow();
      const startCol = range.getColumn();
      const numRows = range.getNumRows();
      const numCols = range.getNumColumns();
      
      this.params.r1 = startRow - 1;
      this.params.r2 = startRow + numRows - 1;
      this.params.c1 = startCol - 1;
      this.params.c2 = startCol + numCols - 1;
      
      return this;
    } catch (e) {
      throw new Error(`Invalid range: ${e.message}`);
    }
  }
  
  setMargins(left, right, top, bottom) {
    const margins = [left, right, top, bottom];
    if (margins.some(m => typeof m !== 'number' || m < 0)) {
      throw new Error('All margins must be non-negative numbers');
    }
    
    this.params.left_margin = left;
    this.params.right_margin = right;
    this.params.top_margin = top;
    this.params.bottom_margin = bottom;
    
    return this;
  }
  
  /**
   * Sets whether to print the date in the export footer.
   * If a timestamp is not explicitly set, the current date will be used.
   * @param {boolean} print - True to print the date, false to hide it.
   * @returns {SheetExporter} The instance for chaining.
   */
  setPrintDate(print) {
    this.params.printdate = print ? 'true' : 'false';
    return this;
  }

  /**
   * Sets whether to print the time in the export footer.
   * If a timestamp is not explicitly set, the current time will be used.
   * @param {boolean} print - True to print the time, false to hide it.
   * @returns {SheetExporter} The instance for chaining.
   */
  setPrintTime(print) {
    this.params.printtime = print ? 'true' : 'false';
    return this;
  }

  /**
   * Sets the timezone to use for rendering the date and time in the footer.
   * This applies to both explicitly set timestamps and the default current time.
   * If not set, the script's default timezone will be used.
   * @param {string} timezone - An IANA timezone name (e.g., "America/New_York").
   * @returns {SheetExporter} The instance for chaining.
   */
  setTimezone(timezone) {
    if (typeof timezone !== 'string' || timezone.length === 0) {
      throw new Error('Timezone must be a valid non-empty string.');
    }
    this._timezone = timezone;
    return this;
  }

  /**
   * Sets the date/time to be printed in the footer. The actual numeric timestamp
   * is calculated just before export, using the specified timezone.
   * If not called, the current date/time will be used when printdate or printtime is enabled.
   * @param {Date} date - The JavaScript Date object to use for the timestamp.
   * @returns {SheetExporter} The instance for chaining.
   */
  setTimestamp(date) {
    if (!(date instanceof Date) || isNaN(date)) {
      throw new Error('A valid JavaScript Date object is required for the timestamp.');
    }
    this._timestampDate = date;
    return this;
  }

  setFileName(fileName) {
    if (!fileName || typeof fileName !== 'string') {
      throw new Error('File name must be a non-empty string');
    }
    this.fileName = fileName;
    return this;
  }
  
  validate() {
    const rangeParams = ['r1', 'r2', 'c1', 'c2'];
    const hasRangeParams = rangeParams.some(p => this.params[p] !== undefined);
    
    if (hasRangeParams) {
      if (!rangeParams.every(p => this.params[p] !== undefined)) {
        throw new Error('When using range parameters, all of r1, r2, c1, and c2 must be specified');
      }
      if (!this.params.gid && this.params.gid !== 0) {
        throw new Error('Range parameters require a sheet ID (gid) to be specified');
      }
    }
    
    const marginParams = ['left_margin', 'right_margin', 'top_margin', 'bottom_margin'];
    const hasMarginParams = marginParams.some(p => this.params[p] !== undefined);
    
    if (hasMarginParams) {
      if (!marginParams.every(p => this.params[p] !== undefined)) {
        throw new Error('When using margins, all four margins must be specified');
      }
    }
  }
  
  buildUrl() {
    this.validate();

    // Process and set the final timestamp parameter just before building the URL
    // If timestamp not explicitly set, use current date/time as default
    if (this.params.printdate === 'true' || this.params.printtime === 'true') {
      const dateToUse = this._timestampDate || new Date();
      const timezone = this._timezone || Session.getScriptTimeZone();
      
      // Format the date into a "wall clock" string for the target timezone
      const dateString = Utilities.formatDate(dateToUse, timezone, "yyyy-MM-dd'T'HH:mm:ss");
      
      // Create a new Date object interpreting that wall clock string as if it were UTC
      const adjustedDate = new Date(dateString + 'Z');
      
      // Convert to the spreadsheet serial number and set the parameter
      this.params.timestamp = this._jsDateToSpreadsheetTimestamp(adjustedDate);
    }
    
    const spreadsheetId = this.spreadsheet.getId();
    const baseUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export`;
    const queryParams = Object.entries(this.params)
      .map(([key, value]) => `${key}=${value}`)
      .join('&');
    
    return queryParams ? `${baseUrl}?${queryParams}` : baseUrl;
  }
  
  exportAsBlob() {
    const url = this.buildUrl();
    
    try {
      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: {
          Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        },
      });
      
      if (response.getResponseCode() !== 200) {
        throw new Error(`Export failed with status ${response.getResponseCode()}: ${response.getContentText()}`);
      }
      
      const blob = response.getBlob();
      blob.setName(this._getFileNameWithExtension());
      
      return blob;
    } catch (e) {
      throw new Error(`Failed to export spreadsheet: ${e.message}`);
    }
  }
  
  _getFileNameWithExtension() {
    const format = this.params.format || 'xlsx'; // Default to xlsx if not set
    const extension = format === 'zip' ? 'zip' : (this.params.format || 'xlsx');
    
    if (this.fileName.endsWith(`.${extension}`)) {
      return this.fileName;
    }
    
    return `${this.fileName}.${extension}`;
  }
  
  usePreset(preset) {
    switch (preset) {
      case 'pdfReport':
        return this
          .setFormat('pdf')
          .setOrientation('portrait')
          .setPageSize('letter')
          .setGridlines(false)
          .setTitle(true)
          .setPageNumbers(true);
      
      case 'pdfLandscape':
        return this
          .setFormat('pdf')
          .setOrientation('landscape')
          .setPageSize('letter')
          .setScale('fitPage');
      
      case 'csvData':
        return this.setFormat('csv');
      
      case 'excelBackup':
        return this.setFormat('xlsx');
      
      default:
        throw new Error(`Unknown preset: ${preset}. Available presets: pdfReport, pdfLandscape, csvData, excelBackup`);
    }
  }
}
