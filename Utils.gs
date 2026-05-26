/**
 * @fileoverview Utility functions for CEMS with Execution Caching
 */

// In-memory cache for a single script execution to prevent repetitive Google Sheets API calls
const EXECUTION_CACHE = {
  sheets: {},
  headers: {},
  data: {}
};

function getSheetByName(name) {
  if (EXECUTION_CACHE.sheets[name]) return EXECUTION_CACHE.sheets[name];
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error(`Sheet '${name}' not found. Check spreadsheet setup.`);
  }
  
  EXECUTION_CACHE.sheets[name] = sheet;
  return sheet;
}

function getColumnIndex(sheet, headerName) {
  const sheetName = sheet.getName();
  
  if (!EXECUTION_CACHE.headers[sheetName]) {
    const lastCol = sheet.getLastColumn() || 1;
    EXECUTION_CACHE.headers[sheetName] = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  }
  
  const index = EXECUTION_CACHE.headers[sheetName].indexOf(headerName);
  if (index === -1) {
    throw new Error(`Header '${headerName}' not found in sheet '${sheetName}'.`);
  }
  return index;
}

function getSheetData(sheet) {
  const sheetName = sheet.getName();
  
  if (EXECUTION_CACHE.data[sheetName]) return EXECUTION_CACHE.data[sheetName];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    EXECUTION_CACHE.data[sheetName] = [];
    return [];
  }
  
  // Cache the entire data range minus the header row
  EXECUTION_CACHE.data[sheetName] = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return EXECUTION_CACHE.data[sheetName];
}

function getRowsForDate(sheet, dateString) {
  const data = getSheetData(sheet);
  
  return data.filter(row => {
    const colA = row[0];
    const rowDateStr = (colA instanceof Date) ? getDateString(colA) : String(colA);
    return rowDateStr === dateString;
  });
}

function findRowByColumnValue(sheet, headerName, value) {
  const colIndex = getColumnIndex(sheet, headerName);
  const data = getSheetData(sheet);
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][colIndex] === value) {
      return i + 2; // +2 offset (0-based index + skipping row 1 header)
    }
  }
  return -1;
}

function appendRow(sheet, rowArray) {
  sheet.appendRow(rowArray);
  // Invalidate cache for this sheet so any subsequent reads in the same run see the new data
  EXECUTION_CACHE.data[sheet.getName()] = null; 
}

function formatTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
}

function parseTimeString(timeStr, referenceDate) {
  const parts = timeStr.split(':');
  const hours = parseInt(parts[0], 10);
  const minutes = parseInt(parts[1], 10);
  const newDate = new Date(referenceDate.getTime());
  newDate.setHours(hours, minutes, 0, 0);
  return newDate;
}

function minutesBetween(startDate, endDate) {
  const diffMs = Math.abs(endDate.getTime() - startDate.getTime());
  return Math.floor(diffMs / (1000 * 60));
}

function getDateString(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function timeFallsWithin(timeStr, windowStart, windowEnd) {
  const timeToMins = (t) => {
    if (!t) return 0;
    const [h, m] = t.split(':').map(Number);
    return h * 60 + m;
  };
  
  const target = timeToMins(timeStr);
  const start = timeToMins(windowStart);
  const end = timeToMins(windowEnd);

  if (start <= end) return target >= start && target <= end;
  return target >= start || target <= end; // Overnight window
}
