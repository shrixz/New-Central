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

function test_validateUserProfile_rejects_missing() {
  let threw = false;
  try { validateUserProfile(null); } catch (e) { threw = e.message.indexOf("Authentication required") !== -1; }
  _assert(threw, "Throws 'Authentication required' for null profile");
}

function test_validateUserProfile_rejects_unknown_email() {
  let threw = false;
  try { validateUserProfile({ email: 'not-a-real-user@nowhere.test' }); }
  catch (e) { threw = e.message.indexOf("not recognized") !== -1; }
  _assert(threw, "Throws 'User account not recognized' for unknown email");
}

function test_validateUserProfile_returns_sheet_values_not_client_values() {
  initializeSheets();
  const result = validateUserProfile({
    email: 'admin@test.com',
    fullName: 'CLIENT-SUPPLIED-NAME-IGNORE-ME',
    role: 'team leader'  // client claims team leader but sheet says admin
  });
  _assert(result.role === 'admin', "Returns role from sheet (admin), not from client (team leader)");
  _assert(result.fullName === 'Admin User', "Returns fullName from sheet, not from client");
  _assert(result.email === 'admin@test.com', "Returns email from sheet");
}

function test_processBulkTransaction_rejects_unknown_email() {
  initializeSheets();
  const result = processBulkTransaction({
    email: 'ghost@nowhere.test',
    user: 'Ghost',
    role: 'admin',
    action: 'PURCHASE_LOG',
    location: 'NCR Hub',
    poNumber: 'PO-10001',
    items: [{ code: 'ITM-001', name: 'Test', uom: 'pc', qty: 1, wbs: '' }]
  });
  _assert(result.success === false, "processBulkTransaction returns success=false for unknown email");
  _assert(result.error.indexOf("not recognized") !== -1, "Error mentions 'not recognized'");
}

function test_notify_writes_one_row_per_recipient() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nSheet = ss.getSheetByName('Notifications');
  const startRows = nSheet.getLastRow();

  const recipients = [
    { email: 'wh@test.com', name: 'Alice Warehouseman', role: 'warehouseman' },
    { email: 'tl1@test.com', name: 'Bob TeamLeader', role: 'team leader' }
  ];
  const sender = { email: 'admin@test.com', name: 'Admin User', role: 'admin' };

  const result = notify(recipients, 'DR_CREATE', sender, 'Test message body', 'DOC-TEST-001');

  _assert(result.inserted === 2, "Inserted 2 rows");
  _assert(nSheet.getLastRow() === startRows + 2, "Sheet has 2 more rows");

  const newRows = nSheet.getRange(startRows + 1, 1, 2, 14).getValues();
  _assert(newRows[0][2] === 'wh@test.com', "Row 1 recipient email");
  _assert(newRows[0][8] === 'DR_CREATE', "Row 1 action");
  _assert(newRows[0][11] === false, "Row 1 Read = false");
  _assert(newRows[0][13].toString().indexOf('sent') === 0 || newRows[0][13].toString().indexOf('failed') === 0,
    "Row 1 Email Status is sent or failed (test inbox may not accept mail)");
}

function test_notify_handles_empty_email_gracefully() {
  initializeSheets();
  const recipients = [{ email: '', name: 'No Email User', role: 'admin' }];
  const sender = { email: 'admin@test.com', name: 'Admin User', role: 'admin' };
  const result = notify(recipients, 'DR_CREATE', sender, 'Test', 'DOC-X');
  _assert(result.inserted === 1, "Still writes a row");
  _assert(result.emailsSent === 0, "Skips empty-email send");
}
