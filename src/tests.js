/**
 * Manual test harness — invoke each function from the Apps Script editor's
 * function dropdown. Output goes to Logger / Console. Assertions throw on
 * failure so the editor reports a red error.
 */

function _assert(cond, msg) {
  if (!cond) throw new Error("ASSERT FAILED: " + msg);
  Logger.log("OK: " + msg);
}

function test_schema_notifications_sheet_exists() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nSheet = ss.getSheetByName('Notifications');
  _assert(nSheet !== null, "Notifications sheet exists");
  const headers = nSheet.getRange(1, 1, 1, 14).getValues()[0];
  _assert(headers[0] === 'Notif ID', "Header col 0 is 'Notif ID'");
  _assert(headers[2] === 'Recipient Email', "Header col 2 is 'Recipient Email'");
  _assert(headers[11] === 'Read', "Header col 11 is 'Read'");
  _assert(headers[13] === 'Email Status', "Header col 13 is 'Email Status'");
  _assert(nSheet.getFrozenRows() === 1, "Frozen rows = 1");
}

function test_schema_requests_user_email_column() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName('Requests');
  const headers = rSheet.getRange(1, 1, 1, rSheet.getLastColumn()).getValues()[0];
  _assert(headers.length >= 14, "Requests has at least 14 columns");
  _assert(headers[13] === 'User Email', "Requests col 13 (14th) is 'User Email'");
}
