/**
 * gas-incremental-excel-import
 *
 * Purpose
 * - Convert Excel files placed in a Drive folder to Google Sheets and
 *   append new rows into a target Google Sheet while ensuring exact overlap
 *   with previously imported rows.
 *
 * Assumptions
 * - Drive Advanced Service (Drive API) is enabled for this project.
 * - The script has permission to read/write Drive and Spreadsheets.
 * - Input files are Excel workbooks. Rows have reverse order (newest first)
 *   and might have abstract headers. Conversion will produce a Google
 *   Spreadsheet whose first sheet contains the data to import.
 * - Filenames contain a parseable date for ordering (e.g. YYYY-MM-DD or YYYYMMDD).
 * - The target spreadsheet and converted spreadsheet may have different
 *   timezones; the converted sheet is re-set to the target sheet timezone
 *   before reading date values.
 * - Import behavior: the import process searches the converted sheet for
 *   the first exact match of the current target's last row. Imported rows
 *   start at that matched row (exclusive) and continue up to the top of
 *   the converted sheet; if no exact match is found the file is skipped
 *   and an error is reported to avoid introducing gaps in the target data.

 * Limitations
 * - Only the first worksheet of the converted spreadsheet is used.
 * - Matching relies on exact row equality after normalization. Date
 *   normalization/parse is best-effort; ambiguous strings may not parse.
 * - Large spreadsheets may hit Apps Script execution time or memory limits.
 * - Moving processed files requires a valid `PROCESSED_FOLDER_ID` and
 *   appropriate Drive permissions.
 *
 * Implemented flows
 * 1) Scan files in `FOLDER_ID` and parse a date from the filename.
 * 2) Sort candidate files by parsed date (oldest first) and skip files
 *    without a parseable filename date.
 * 3) For each file:
 *    - Convert the Excel file to a Google Spreadsheet (`Drive.Files.create`).
 *    - Set the converted spreadsheet timezone to match the target sheet.
 *    - Read the first sheet's data (first sheet only).
 *    - Determine the first-column value type from the last non-empty
 *      source row (used to skip header/text rows at top of the file).
 *    - Find an exact match for the last row present in the target sheet.
 *      If no exact match is found, the file is rejected to avoid gaps.
 *    - Import all earlier rows (skipping rows whose first-column type
 *      doesn't match the detected type) and append them to the target.
 *    - Delete the temporary converted spreadsheet and move the original
 *      file into `PROCESSED_FOLDER_ID` (if set).
 *
 * Debugging
 * - Set `DEBUG = true` to enable `Logger.log` debug output.
 */

// STUB VALUES - replace these with real IDs before running in production
const FOLDER_ID = 'REPLACE_ME_FOLDER_ID';
const TARGET_SHEET_ID = 'REPLACE_ME_TARGET_SHEET_ID';
const TARGET_SHEET_NAME = 'Sheet1';
const PROCESSED_FOLDER_ID = 'REPLACE_ME_PROCESSED_FOLDER_ID';
const DEBUG = false;

function debugLog(message) {
  if (DEBUG) {
    Logger.log(message);
  }
}

function main() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  const targetSs = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const targetSheet = targetSs.getSheetByName(TARGET_SHEET_NAME) || targetSs.getSheets()[0];
  const fileEntries = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    const fileDate = parseDateFromFilename(name);
    fileEntries.push({ file, name, fileDate });
  }

  fileEntries.sort((a, b) => {
    if (!a.fileDate && !b.fileDate) return 0;
    if (!a.fileDate) return 1;
    if (!b.fileDate) return -1;
    return a.fileDate - b.fileDate;
  });

  for (const entry of fileEntries) {
    const { file, name, fileDate: entryFileDate } = entry;
    debugLog('Checking file: ' + name);

    if (!entryFileDate) {
      debugLog('Skipping ' + name + ': no parseable date in filename');
      continue;
    }

    const lastTargetRowValues = getLastTargetRowValues(targetSheet);

    try {
      processSingleFile(file, name, targetSheet, lastTargetRowValues);
    } catch (e) {
      Logger.log('Import error for ' + name + ': ' + e);
    }
  }
}

function getLastDateInSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return new Date(0);
  const val = sheet.getRange(lastRow, 1).getValue();
  if (val instanceof Date) return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  const parsed = new Date(val);
  if (!isNaN(parsed)) return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  return new Date(0);
}

function parseDateValue(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === 'number') {
    return new Date(value);
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return null;

    const direct = new Date(trimmed);
    if (!isNaN(direct)) {
      return new Date(direct.getFullYear(), direct.getMonth(), direct.getDate());
    }

    let m = trimmed.match(/(\d{4})[-_\.]?(\d{2})[-_\.]?(\d{2})/);
    if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));

    m = trimmed.match(/(\d{8})/);
    if (m) {
      const s = m[1];
      return new Date(parseInt(s.substring(0, 4), 10), parseInt(s.substring(4, 6), 10) - 1, parseInt(s.substring(6, 8), 10));
    }
  }

  return null;
}

function parseDateFromFilename(name) {
  // Try patterns: YYYY-MM-DD or YYYYMMDD or similar numeric forms
  return parseDateValue(name);
}

function getValueType(value) {
  if (value instanceof Date) return 'date';
  if (value === null || value === undefined) return 'null';
  return typeof value;
}

function getLastTargetRowValues(sheet) {
  const lastRow = sheet.getLastRow();
  return lastRow > 0
    ? sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0]
    : null;
}

function processSingleFile(file, name, targetSheet, lastTargetRowValues) {
  const resource = { title: name, mimeType: 'application/vnd.google-apps.spreadsheet' };
  const converted = Drive.Files.create(resource, file.getBlob(), { convert: true });

  const tempSs = SpreadsheetApp.openById(converted.id);
  const targetTimeZone = targetSheet.getParent().getSpreadsheetTimeZone();
  if (typeof tempSs.setSpreadsheetTimeZone === 'function') {
    tempSs.setSpreadsheetTimeZone(targetTimeZone);
    debugLog('Converted sheet timezone set to ' + targetTimeZone);
  }

  const srcSheet = tempSs.getSheets()[0];
  const data = srcSheet.getDataRange().getValues();

  let matchIndex = -1;
  let firstColType = null;
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    if (!row || !row.length) continue;

    if (firstColType === null) {
      firstColType = getValueType(row[0]);
      debugLog('First non-empty source row type is ' + firstColType + ' for file ' + name);
    }

    if (lastTargetRowValues && rowsEqual(lastTargetRowValues, row)) {
      matchIndex = i;
      break;
    }
  }

  if (lastTargetRowValues && matchIndex === -1) {
    throw new Error('Potential hole in data: no exact match for existing last row');
  }

  const rowsToImport = [];
  let startIndex = matchIndex >= 0 ? matchIndex - 1 : data.length - 1;

  for (let i = startIndex; i >= 0; i--) {
    const row = data[i];
    if (!row || !row.length) continue;

    const currentType = getValueType(row[0]);
    if (firstColType === null) {
      firstColType = currentType;
      rowsToImport.push(row);
      continue;
    }

    if (currentType !== firstColType) {
      debugLog('Skipping row ' + (i + 1) + ' from ' + name + ': first column type "' + currentType + '" does not match "' + firstColType + '"');
      debugLog('Row values: ' + JSON.stringify(row));
      continue;
    }

    rowsToImport.push(row);
  }

  if (rowsToImport.length) {
    const startRow = targetSheet.getLastRow() + 1 || 1;
    targetSheet.getRange(startRow, 1, rowsToImport.length, rowsToImport[0].length).setValues(rowsToImport);
    debugLog('Imported ' + rowsToImport.length + ' rows from ' + name);
  } else {
    debugLog('No rows to import from ' + name);
  }

  Drive.Files.remove(converted.id); // remove temp converted file

  if (PROCESSED_FOLDER_ID) {
    try {
      // pass null for mediaData so the API treats the 4th arg as optionalArgs
      Drive.Files.update({}, file.getId(), null, { addParents: PROCESSED_FOLDER_ID, removeParents: FOLDER_ID });
      debugLog('Moved ' + name + ' to processed folder');
    } catch (e) {
      debugLog('Failed to move ' + name + ' to processed folder: ' + (e && e.message ? e.message : e));
    }
  }
}

function normalizeValue(value) {
  if (value instanceof Date) {
    return value.getTime();
  }
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function rowsEqual(rowA, rowB) {
  const length = Math.max(rowA.length, rowB.length);
  for (let i = 0; i < length; i++) {
    if (normalizeValue(rowA[i]) !== normalizeValue(rowB[i])) {
      return false;
    }
  }
  return true;
}
