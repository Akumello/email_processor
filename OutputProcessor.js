/** Output column order — must match the Output sheet headers. */
const OUTPUT_HEADERS = ['Category', 'Label', 'Importance', 'Read Status'];

/** Map from JSON key → Output header for flexible key matching. */
const JSON_KEY_MAP = {
  category: 'Category',
  label: 'Label',
  importance: 'Importance',
  readStatus: 'Read Status'
};

/**
 * Processes unprocessed rows in the Raw Emails sheet.
 * Reads the AI JSON from column C, appends parsed values to the Output sheet,
 * and marks each processed row with TRUE in column D (Processed).
 * @returns {{rowsAdded: number, errors: number, skipped: number}} Result summary for the client
 */
function processOutputJson() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(SHEET_NAMES.RAW_EMAILS);
  if (!rawSheet) {
    Logger.log('[processOutputJson] ERROR: Sheet "%s" not found.', SHEET_NAMES.RAW_EMAILS);
    throw new Error('Sheet "' + SHEET_NAMES.RAW_EMAILS + '" not found. Run a scan first.');
  }

  const lastRow = rawSheet.getLastRow();
  Logger.log('[processOutputJson] Starting. Raw Emails lastRow=%s', lastRow);

  if (lastRow <= 1) {
    Logger.log('[processOutputJson] No data rows found. Exiting early.');
    return { rowsAdded: 0, errors: 0, skipped: 0 };
  }

  // Read columns C (AI) and D (Processed) for all data rows
  const dataRange = rawSheet.getRange(2, 3, lastRow - 1, 2).getValues();
  Logger.log('[processOutputJson] Read %s data rows (cols C-D).', dataRange.length);
  const outputSheet = ensureOutputSheet_(ss);

  const newRows = [];
  const processedRowIndices = []; // 0-based indices into dataRange
  let errorCount = 0;
  let skippedCount = 0;

  for (let i = 0; i < dataRange.length; i++) {
    const sheetRow = i + 2; // 1-based sheet row number
    const aiValue = dataRange[i][0];
    const processed = dataRange[i][1];

    // Skip already-processed rows
    if (processed === true || String(processed).trim().toUpperCase() === 'TRUE') {
      skippedCount++;
      Logger.log('[processOutputJson] Row %s: SKIPPED (already processed).', sheetRow);
      continue;
    }

    // Skip empty AI cells
    if (!aiValue || String(aiValue).trim() === '') {
      Logger.log('[processOutputJson] Row %s: SKIPPED (AI cell empty).', sheetRow);
      continue;
    }

    Logger.log('[processOutputJson] Row %s: Parsing AI JSON (%s chars)...', sheetRow, String(aiValue).length);
    const parsed = cleanAndParseJson_(String(aiValue));
    if (parsed === null) {
      errorCount++;
      Logger.log('[processOutputJson] Row %s: ERROR — failed to parse JSON. Raw value: %s',
        sheetRow, String(aiValue).substring(0, 200));
      continue;
    }

    Logger.log('[processOutputJson] Row %s: Parsed OK → category=%s, label=%s, importance=%s, readStatus=%s',
      sheetRow, parsed.category, parsed.label, parsed.importance, parsed.readStatus);

    // Build a row matching OUTPUT_HEADERS order
    const row = OUTPUT_HEADERS.map(header => {
      // Find the JSON key that maps to this header
      const jsonKey = Object.keys(JSON_KEY_MAP).find(k => JSON_KEY_MAP[k] === header);
      if (!jsonKey) return '';
      const val = parsed[jsonKey];
      if (val === null || val === undefined) return '';
      if (typeof val === 'object') return JSON.stringify(val);
      return val;
    });

    newRows.push(row);
    processedRowIndices.push(i);
  }

  Logger.log('[processOutputJson] Loop done. newRows=%s, errors=%s, skipped=%s',
    newRows.length, errorCount, skippedCount);

  // Append new rows to Output sheet in one batch
  if (newRows.length > 0) {
    const startRow = outputSheet.getLastRow() + 1;
    Logger.log('[processOutputJson] Writing %s rows to Output sheet starting at row %s.', newRows.length, startRow);
    outputSheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);

    // Mark processed rows in column D (Processed) of Raw Emails
    for (const idx of processedRowIndices) {
      rawSheet.getRange(idx + 2, 4).setValue(true); // idx+2 because row 1 is header
    }
    Logger.log('[processOutputJson] Marked %s rows as processed in column D.', processedRowIndices.length);

    SpreadsheetApp.flush();
    Logger.log('[processOutputJson] Flushed changes to spreadsheet.');
  } else {
    Logger.log('[processOutputJson] No new rows to write.');
  }

  const result = {
    rowsAdded: newRows.length,
    errors: errorCount,
    skipped: skippedCount
  };
  Logger.log('[processOutputJson] Done. Result: %s', JSON.stringify(result));
  return result;
}

/**
 * Cleans a raw string by trimming non-JSON characters from the start and end,
 * then parses it as JSON. Trims until '{' is found at the start and '}' at the end.
 * @param {string} raw - The raw string potentially containing JSON
 * @returns {Object|null} Parsed JSON object, or null if parsing fails
 * @private
 */
function cleanAndParseJson_(raw) {
  let str = raw.trim();

  // Find the first '{' and trim everything before it
  const openIdx = str.indexOf('{');
  if (openIdx === -1) return null;
  str = str.substring(openIdx);

  // Find the last '}' and trim everything after it
  const closeIdx = str.lastIndexOf('}');
  if (closeIdx === -1) return null;
  str = str.substring(0, closeIdx + 1);

  try {
    return JSON.parse(str);
  } catch (e) {
    return null;
  }
}

/**
 * Ensures the Output sheet exists. Creates it if missing.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The active spreadsheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Output sheet
 * @private
 */
/**
 * Ensures the Output sheet exists with the correct headers.
 * Creates it with headers if missing; adds headers if the sheet is empty.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The active spreadsheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Output sheet
 * @private
 */
function ensureOutputSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.OUTPUT);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.OUTPUT);
  }

  // Add headers if the sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(OUTPUT_HEADERS);
    sheet.getRange('1:1').setFontWeight('bold');
  }

  return sheet;
}
