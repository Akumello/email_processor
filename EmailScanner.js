/**
 * Scans Gmail inbox for new emails and appends them to the Raw Emails sheet.
 * Called by the client-side button and optionally by a time-driven trigger.
 * @returns {{newCount: number, totalProcessed: number}} Result summary for the client
 */
function scanForNewEmails() {
  const sheet = ensureSheet_();
  const processedIds = getProcessedMessageIds_(sheet);
  const query = buildSearchQuery_();
  const threads = GmailApp.search(query, 0, EMAIL_CONFIG.maxResults);

  const newRows = [];

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const message of messages) {
      const messageId = message.getId();
      if (processedIds.has(messageId)) continue;

      const subject = message.getSubject() || '(no subject)';
      const body = message.getPlainBody() || '';
      const date = message.getDate();
      const sender = message.getFrom() || '(unknown sender)';

      const contents = [
        'From: ' + sender,
        'Date: ' + Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
        'Subject: ' + subject,
        '',
        body
      ].join('\n');

      newRows.push([messageId, contents]);
      processedIds.add(messageId);
    }
  }

  if (newRows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, 2).setValues(newRows);
    SpreadsheetApp.flush();
  }

  return {
    newCount: newRows.length,
    totalProcessed: processedIds.size
  };
}

/**
 * Reads all processed message IDs from column A of the Raw Emails sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Raw Emails sheet
 * @returns {Set<string>} Set of previously processed Gmail message IDs
 * @private
 */
function getProcessedMessageIds_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return new Set();

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return new Set(ids.map(row => row[0]).filter(Boolean));
}

/**
 * Builds a Gmail search query string from EMAIL_CONFIG.
 * Always includes the cutoff date filter.
 * @returns {string} Gmail search query
 * @private
 */
function buildSearchQuery_() {
  const parts = [EMAIL_CONFIG.query, 'after:' + EMAIL_CONFIG.cutoffDate];

  if (EMAIL_CONFIG.label) {
    parts.push('label:' + EMAIL_CONFIG.label);
  }
  if (EMAIL_CONFIG.from) {
    parts.push('from:' + EMAIL_CONFIG.from);
  }
  if (EMAIL_CONFIG.subject) {
    parts.push('subject:(' + EMAIL_CONFIG.subject + ')');
  }

  return parts.join(' ');
}

/**
 * Ensures the Raw Emails sheet exists with the correct headers.
 * Creates it if missing.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Raw Emails sheet
 * @private
 */
function ensureSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.RAW_EMAILS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.RAW_EMAILS);
    sheet.appendRow(['Message ID', 'Contents', 'AI', 'Processed']);
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 600);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 100);
    sheet.getRange('1:1').setFontWeight('bold');
  }

  return sheet;
}
