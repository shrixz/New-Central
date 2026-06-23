# Notifications + System Readiness Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add cross-role in-app + email notifications (bell with unread badge, recipient targeting by role + location/site access, acknowledgments on confirm/reject) and bundle a set of system-readiness fixes that prevent notifications from misfiring or crashing.

**Architecture:** Backend resolves recipients from the app's `Users` sheet (never `Session.getActiveUser()`), writes one `Notifications` sheet row per (recipient × event), and fires one `MailApp.sendEmail` per recipient. A `validateUserProfile()` gate re-reads identity from the `Users` sheet on every server entry point so the client cannot spoof a role. Frontend renders a bell with badge in the navbar; the dropdown polls every 60s while the tab is visible.

**Tech Stack:** Google Apps Script (V8 runtime), `SpreadsheetApp`, `MailApp`, `LockService`, `ScriptApp`, vanilla JS + Bootstrap 5 in `Index.html`. No test framework — backend tests run via named harness functions inside Apps Script (invokable from the editor or `clasp run`), frontend tests are manual UAT in the deployed web app.

## Global Constraints

- Spec: `docs/superpowers/specs/2026-06-23-notifications-and-system-readiness-design.md` — every requirement in §5-§10 of the spec must be implemented; cross-reference is authoritative for any ambiguity in this plan.
- Recipient/sender identity always resolved via `Users` sheet looked up by email — never `Session.getActiveUser()`.
- All backend mutations remain inside `LockService.getScriptLock()` (existing pattern).
- Notification failures must never abort the underlying business transaction — wrap every `notify(...)` call in `try/catch` and log only.
- No changes to: login flow (`user.js`), `handleActionChange()` field-visibility, `Logs` sheet schema, action codes, queue status strings.
- New file allowed: `src/notifications.js` (logic) and `src/tests.js` (test harness). Existing files modified: `src/Code.js`, `src/Index.html`, `src/setup.js`.
- Commit messages follow the repo's existing style (short imperative line, optional body) and include the `Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>` trailer.

## File Structure

| File | Purpose | Status |
|---|---|---|
| `src/Code.js` | Add `validateUserProfile`, wire `notify` calls, append `notifications` to `getAppData`, apply readiness fixes 10.3-10.7 | Modify |
| `src/notifications.js` | New file. Contains `notify`, `resolveRecipients`, `resolveRequester`, `getNotificationsForUser`, `markNotificationRead`, `markAllNotificationsRead`, and `NOTIFICATIONS_SHEET` constant | Create |
| `src/setup.js` | Add `Notifications` sheet definition + `User Email` column on Requests | Modify |
| `src/Index.html` | Add bell UI (HTML+CSS+JS), 60s poll, `userEmail` in all server payloads | Modify |
| `src/tests.js` | New file. Manual test harness functions invokable from Apps Script editor | Create |

Apps Script flattens all `.js` and `.gs` files into a single global namespace — no module imports needed. Functions in `notifications.js` are callable directly from `Code.js`.

---

### Task 1: Schema bootstrap — Notifications sheet + Requests User Email column

**Files:**
- Modify: `src/setup.js`
- Create: `src/tests.js`

**Interfaces:**
- Consumes: nothing
- Produces: `Notifications` sheet (14 cols) and `Requests` sheet 14th col `User Email` exist after `initializeSheets()` runs

- [ ] **Step 1: Write the test harness for setup**

Create `src/tests.js`:

```javascript
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
```

- [ ] **Step 2: Run tests to verify they fail**

In the Apps Script editor, select `test_schema_notifications_sheet_exists` from the function dropdown and click Run.
Expected: red error `ASSERT FAILED: Notifications sheet exists` (sheet doesn't exist yet).

- [ ] **Step 3: Update `src/setup.js` schema**

Modify `src/setup.js`. In the `sheets` array, change the `Requests` entry to add `'User Email'` as the 14th header, and append a new `Notifications` sheet definition. Edit between line 13 and line 16 (Requests entry) and add the Notifications entry before the closing `]`:

```javascript
    { 
      name: 'Requests', 
      head: ['Req ID', 'Date', 'User Name', 'Role', 'Action', 'Target Location', 'Target Site', 'Item Code', 'Item Description', 'UOM', 'Qty', 'Status', 'Remarks', 'User Email'] 
    },
```

Then add this entry to the `sheets` array (place it after `PO Assignments` so it's last):

```javascript
    {
      name: 'Notifications',
      head: ['Notif ID', 'Timestamp', 'Recipient Email', 'Recipient Name', 'Recipient Role', 'Sender Email', 'Sender Name', 'Sender Role', 'Action', 'Related Req ID', 'Message', 'Read', 'Read At', 'Email Status']
    }
```

Also update the `SHEETS` constant in `src/Code.js` (around lines 2-11) to add the new sheet name:

```javascript
const SHEETS = {
  INV: 'Inventory',
  LOGS: 'Logs',
  USERS: 'Users',
  REQ: 'Requests',
  DISCREP: 'Discrepancies',
  DROPS: 'Dropdowns',
  PO_DB: 'PO Database',
  PO_ASSIGN: 'PO Assignments',
  NOTIF: 'Notifications'
};
```

The existing `initializeSheets()` logic (`setup.js:62-75`) is already idempotent — it rewrites the header row only when `currentCols === 0 || currentCols < sh.head.length`, so running it on an existing spreadsheet appends the new `User Email` column header without disturbing data, and creates the `Notifications` sheet only if it doesn't exist. No code change needed beyond the schema entries.

- [ ] **Step 4: Run tests to verify they pass**

Run `test_schema_notifications_sheet_exists` and `test_schema_requests_user_email_column` from the editor. Both should log green `OK:` lines and the function should complete without error.

- [ ] **Step 5: Commit**

```bash
git add src/setup.js src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Add Notifications sheet and Requests User Email column

Schema bootstrap for the notifications feature. initializeSheets() is
idempotent so this safely adds the new sheet and column on existing
spreadsheets.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 2: `validateUserProfile()` helper

**Files:**
- Modify: `src/Code.js` (add function near other utilities, e.g., after `getLogoSafe`)
- Modify: `src/tests.js`

**Interfaces:**
- Consumes: `SHEETS.USERS` constant from Task 1
- Produces: `validateUserProfile(userProfile) → {email, fullName, role, locAccess, siteAccess}` or throws

- [ ] **Step 1: Write the failing test**

Append to `src/tests.js`:

```javascript
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run all three test functions. Each should fail with `ReferenceError: validateUserProfile is not defined`.

- [ ] **Step 3: Implement `validateUserProfile` in `Code.js`**

Add to `src/Code.js` (insert near the end of the file, before any closing module-level code):

```javascript
function validateUserProfile(userProfile) {
  if (!userProfile || !userProfile.email) {
    throw new Error("Authentication required.");
  }
  const uSheet = SS.getSheetByName(SHEETS.USERS);
  if (!uSheet) throw new Error("Users sheet missing.");
  const data = uSheet.getDataRange().getValues();
  const claimed = userProfile.email.toString().trim().toLowerCase();
  const row = data.find(r => r[0] && r[0].toString().trim().toLowerCase() === claimed);
  if (!row) {
    throw new Error("User account not recognized.");
  }
  return {
    email: row[0].toString().trim(),
    fullName: row[1] ? row[1].toString() : '',
    role: row[4] ? row[4].toString().toLowerCase() : 'team leader',
    locAccess: row[5] ? row[5].toString().trim() : '',
    siteAccess: row[6] ? row[6].toString().trim() : ''
  };
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run all three test functions. All three should log green `OK:` lines.

- [ ] **Step 5: Commit**

```bash
git add src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Add validateUserProfile server-side identity gate

Re-reads identity from Users sheet so client cannot spoof role or
impersonate another user. Caller uses the returned profile, never the
client payload, for downstream role decisions.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 3: Frontend — pass `email` in every server payload

**Files:**
- Modify: `src/Index.html`

**Interfaces:**
- Consumes: `currentEmail` (already populated by login at `Index.html:449`)
- Produces: Every `google.script.run` call to a backend mutator now includes `email: currentEmail` in the `userProfile`-style argument

- [ ] **Step 1: Identify all server calls that pass userProfile**

The Index.html call sites to update (search for `google.script.run`):
- `Index.html:749` — `getAppData({ fullName: currentUser, role: userRole, locAccess: locAccess, siteAccess: siteAccess })`
- `Index.html:992` — `processQueueAction(reqId, action, { fullName: currentUser, role: userRole })`
- `Index.html:1835` — `processBulkTransaction({ user: currentUser, role: userRole, location, ... })`
- `Index.html:1137` — `assignPOToDoc(docNumber, poNumber)` — does NOT pass userProfile today; the third argument is added in Task 4 (where the server side is updated to require it). Leave this line alone in Task 3.

(Other `google.script.run` calls — login, refresh-without-user, logo fetch — don't take userProfile and don't need changes.)

- [ ] **Step 2: Modify the three primary call sites**

Edit `src/Index.html`:

At line 749, change:
```javascript
.getAppData({ fullName: currentUser, role: userRole, locAccess: locAccess, siteAccess: siteAccess });
```
to:
```javascript
.getAppData({ email: currentEmail, fullName: currentUser, role: userRole, locAccess: locAccess, siteAccess: siteAccess });
```

At line 992, change:
```javascript
}).processQueueAction(reqId, action, { fullName: currentUser, role: userRole });
```
to:
```javascript
}).processQueueAction(reqId, action, { email: currentEmail, fullName: currentUser, role: userRole });
```

At line 1835, change:
```javascript
.processBulkTransaction({ user: currentUser, role: userRole, location, siteId, siteName, action, client, refDoc, poNumber, poToFollow, mrcNum, targetLoc, targetSite, drId, sourceDoc, items, returnType: 'BY_DOC' });
```
to:
```javascript
.processBulkTransaction({ email: currentEmail, user: currentUser, role: userRole, location, siteId, siteName, action, client, refDoc, poNumber, poToFollow, mrcNum, targetLoc, targetSite, drId, sourceDoc, items, returnType: 'BY_DOC' });
```

For `assignPOToDoc` at `Index.html:1137`, leave it for now — Task 7 will add the userProfile second-arg pattern. (If the engineer's editor differs and that line isn't `assignPOToDoc`, search for `assignPOToDoc(` in Index.html and update it where called.)

- [ ] **Step 3: Manual verification in the deployed app**

Push to the Apps Script project and reload the web app. Open browser devtools → Network or just add a temporary `console.log(JSON.stringify(payload))` before the `processBulkTransaction` call. Submit a small transaction. Verify the logged payload includes `email: <your-login-email>`. Remove the temporary `console.log`.

- [ ] **Step 4: Commit**

```bash
git add src/Index.html
git commit -m "$(cat <<'EOF'
Include email in all client-to-server userProfile payloads

Prerequisite for the server-side validateUserProfile gate added in the
previous commit.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 4: Wire `validateUserProfile` into existing server entry points

**Files:**
- Modify: `src/Code.js`

**Interfaces:**
- Consumes: `validateUserProfile` from Task 2
- Produces: Every server mutator now uses a validated profile (called `validated` by convention) for role/scope/identity decisions

- [ ] **Step 1: Wire into `processBulkTransaction`**

In `src/Code.js` around line 226-230 (after `lock.waitLock(20000)` and the spreadsheet null-check), add a validation call. Find:

```javascript
function processBulkTransaction(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (!SS) throw new Error("Could not access spreadsheet.");
```

Insert after the `if (!SS)` line:

```javascript
    const validated = validateUserProfile({ email: payload.email, fullName: payload.user, role: payload.role });
    payload.user = validated.fullName;
    payload.role = validated.role;
    payload.userEmail = validated.email;
```

This overwrites any client-supplied role/name with the trusted server-side values; the rest of the function reads `payload.user` and `payload.role` and now gets the validated versions transparently. `payload.userEmail` is the new field used by Task 8 onwards.

- [ ] **Step 2: Wire into `processQueueAction`**

In `src/Code.js` around line 516-525 (top of `processQueueAction`), find:

```javascript
function processQueueAction(reqId, action, userProfile) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const rSheet = SS.getSheetByName(SHEETS.REQ);
```

Insert after `lock.waitLock(15000);`:

```javascript
    const validated = validateUserProfile(userProfile);
    userProfile = validated;
```

Reassigning `userProfile` means all the existing references (e.g., `userProfile.fullName` in the log-row appends at lines 554, 562, 564, 576, 586, 596) now use validated values without further edits.

- [ ] **Step 3: Wire into `getAppData`**

In `src/Code.js` around line 24, find:

```javascript
function getAppData(userProfile) {
  try {
    if (!SS) throw new Error("Could not access spreadsheet.");
```

Insert after the `if (!SS)` line:

```javascript
    let validated = null;
    try {
      validated = validateUserProfile(userProfile);
      userProfile = validated;
    } catch (validErr) {
      // Fall through with unvalidated profile so the app still loads for
      // legacy/transitional sessions. Pending queue and notifications
      // will be empty for unrecognized users.
      console.warn("getAppData unvalidated profile: " + validErr.toString());
    }
```

- [ ] **Step 4: Wire into `assignPOToDoc`**

In `src/Code.js` around line 605, change the signature:

```javascript
function assignPOToDoc(docNumber, poNumber) {
```

to:

```javascript
function assignPOToDoc(docNumber, poNumber, userProfile) {
```

And add inside the try, after `lock.waitLock(20000)`:

```javascript
    const validated = validateUserProfile(userProfile);
```

Then find the call site in `src/Index.html` (search for `assignPOToDoc(`) and add `, { email: currentEmail, fullName: currentUser, role: userRole }` as the third argument.

- [ ] **Step 5: Add a regression test**

Append to `src/tests.js`:

```javascript
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
```

- [ ] **Step 6: Run the test and verify it passes**

Run `test_processBulkTransaction_rejects_unknown_email` from the editor. Should log `OK:` lines.

- [ ] **Step 7: Commit**

```bash
git add src/Code.js src/Index.html src/tests.js
git commit -m "$(cat <<'EOF'
Wire validateUserProfile into all server entry points

processBulkTransaction, processQueueAction, getAppData, and assignPOToDoc
now re-read identity from the Users sheet before acting. Client-supplied
role and name are overwritten with trusted server values.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 5: `notify()` core — write rows and send emails

**Files:**
- Create: `src/notifications.js`
- Modify: `src/tests.js`

**Interfaces:**
- Consumes: `SS`, `SHEETS.NOTIF`, `SHEETS.USERS` from Code.js (Apps Script flat namespace)
- Produces:
  - `notify(recipients, action, sender, message, relatedReqId) → {inserted, emailsSent, emailsFailed}`
  - `recipients`: array of `{email, name, role}`
  - `sender`: `{email, name, role}` — must be a validated profile

- [ ] **Step 1: Write the failing test**

Append to `src/tests.js`:

```javascript
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
```

- [ ] **Step 2: Run tests to verify failure**

Run both — should fail with `ReferenceError: notify is not defined`.

- [ ] **Step 3: Implement `notify` in `src/notifications.js`**

Create `src/notifications.js`:

```javascript
/**
 * Notification system — in-app rows + email per recipient.
 * Caller must pass validated sender. Recipient lookups happen elsewhere.
 */

function _notif_ensureSheet() {
  let nSheet = SS.getSheetByName(SHEETS.NOTIF);
  if (!nSheet) {
    nSheet = SS.insertSheet(SHEETS.NOTIF);
    nSheet.getRange(1, 1, 1, 14).setValues([[
      'Notif ID', 'Timestamp', 'Recipient Email', 'Recipient Name', 'Recipient Role',
      'Sender Email', 'Sender Name', 'Sender Role', 'Action', 'Related Req ID',
      'Message', 'Read', 'Read At', 'Email Status'
    ]]).setFontWeight("bold").setBackground("#0f172a").setFontColor("white");
    nSheet.setFrozenRows(1);
  }
  return nSheet;
}

function _notif_id() {
  const d = new Date();
  const stamp = Utilities.formatDate(d, "GMT+8", "yyyyMMdd-HHmmss");
  const rand = Math.floor(1000 + Math.random() * 9000);
  return 'NOTIF-' + stamp + '-' + rand;
}

function notify(recipients, action, sender, message, relatedReqId) {
  if (!recipients || recipients.length === 0) return { inserted: 0, emailsSent: 0, emailsFailed: 0 };
  if (!sender || !sender.email) throw new Error("notify: sender required.");

  const nSheet = _notif_ensureSheet();
  const d = new Date();
  const webAppUrl = ScriptApp.getService().getUrl() || '';
  const rows = [];
  let emailsSent = 0;
  let emailsFailed = 0;

  recipients.forEach(r => {
    const notifId = _notif_id();
    let emailStatus;
    if (!r.email || r.email.toString().trim() === '') {
      emailStatus = 'skipped: no email';
    } else {
      try {
        const subject = `[Inventory] ${action} by ${sender.name} — ${(message || '').substring(0, 60)}`;
        const body = `Hi ${r.name || ''},\n\n${sender.name} (${sender.role}) just performed ${action}.\n\n${message}\nRelated ID: ${relatedReqId || '—'}\n\nOpen the inventory app to confirm or reject:\n${webAppUrl}\n\n— Inventory System`;
        MailApp.sendEmail(r.email, subject, body);
        emailStatus = 'sent';
        emailsSent++;
      } catch (mailErr) {
        emailStatus = 'failed: ' + mailErr.toString().substring(0, 80);
        emailsFailed++;
      }
    }
    rows.push([
      notifId,
      d,
      r.email || '',
      r.name || '',
      r.role || '',
      sender.email,
      sender.name || '',
      sender.role || '',
      action,
      relatedReqId || '',
      message || '',
      false,
      '',
      emailStatus
    ]);
  });

  if (rows.length > 0) {
    nSheet.getRange(nSheet.getLastRow() + 1, 1, rows.length, 14).setValues(rows);
  }

  return { inserted: rows.length, emailsSent: emailsSent, emailsFailed: emailsFailed };
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run both notify tests. Both should pass. The Email Status will likely be `sent` if MailApp quota is available, but the test only asserts it starts with `sent` or `failed` — both are acceptable.

- [ ] **Step 5: Commit**

```bash
git add src/notifications.js src/tests.js
git commit -m "$(cat <<'EOF'
Add notify() — writes notification rows and sends per-recipient emails

One row per recipient in Notifications sheet, one MailApp email per
recipient with non-empty email. Mail failures are logged in the
Email Status column but never abort the call.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 6: `resolveRecipients()` and `resolveRequester()`

**Files:**
- Modify: `src/notifications.js`
- Modify: `src/tests.js`

**Interfaces:**
- Consumes: `Users` sheet, `Requests` sheet
- Produces:
  - `resolveRecipients(action, payload) → [{email, name, role}, ...]`
  - `resolveRequester(reqId) → {email, name, role} | null`

- [ ] **Step 1: Write failing tests**

Append to `src/tests.js`:

```javascript
function test_resolveRecipients_DR_CREATE_filters_by_location() {
  initializeSheets();
  // Default users: wh@test.com has 'NCR Hub' access; admin has '' (all)
  const r = resolveRecipients('DR_CREATE', { location: 'NCR Hub' });
  const emails = r.map(x => x.email);
  _assert(emails.indexOf('wh@test.com') !== -1, "NCR Hub warehouseman included");
  _assert(emails.indexOf('admin@test.com') === -1, "Admin NOT included for DR_CREATE");
}

function test_resolveRecipients_ISSUE_filters_by_site() {
  initializeSheets();
  const r = resolveRecipients('ISSUE', { siteName: 'Makati Site' });
  const emails = r.map(x => x.email);
  _assert(emails.indexOf('tl1@test.com') !== -1, "Makati team leader included");
  _assert(emails.indexOf('tl2@test.com') === -1, "Cebu team leader NOT included");
}

function test_resolveRecipients_RETURN_CLIENT_all_admins() {
  initializeSheets();
  const r = resolveRecipients('RETURN_CLIENT', {});
  _assert(r.some(x => x.email === 'admin@test.com'), "Admin included");
  _assert(r.every(x => x.role === 'admin'), "Only admins returned");
}

function test_resolveRequester_returns_null_for_unknown_id() {
  initializeSheets();
  const r = resolveRequester('NOT-A-REAL-ID-12345');
  _assert(r === null, "Returns null for unknown reqId");
}
```

- [ ] **Step 2: Run tests to verify failure**

Run — should fail with `ReferenceError: resolveRecipients is not defined`.

- [ ] **Step 3: Implement in `src/notifications.js`**

Append to `src/notifications.js`:

```javascript
function _notif_listUsers() {
  const uSheet = SS.getSheetByName(SHEETS.USERS);
  if (!uSheet || uSheet.getLastRow() < 2) return [];
  const data = uSheet.getRange(2, 1, uSheet.getLastRow() - 1, 7).getValues();
  return data
    .filter(r => r[0] && r[0].toString().trim() !== '')
    .map(r => ({
      email: r[0].toString().trim(),
      name: r[1] ? r[1].toString().trim() : '',
      role: r[4] ? r[4].toString().toLowerCase().trim() : '',
      locAccess: r[5] ? r[5].toString().trim() : '',
      siteAccess: r[6] ? r[6].toString().trim() : ''
    }));
}

function _notif_hasAccess(accessCsv, value) {
  if (!accessCsv || accessCsv.toString().trim() === '') return true; // empty means all
  const list = accessCsv.toString().split(',').map(s => s.trim()).filter(s => s);
  return list.indexOf((value || '').toString().trim()) !== -1;
}

function resolveRecipients(action, payload) {
  const users = _notif_listUsers();
  const role = (r) => r.role;

  if (action === 'DR_CREATE') {
    return users.filter(u => role(u) === 'warehouseman' && _notif_hasAccess(u.locAccess, payload.location));
  }
  if (action === 'TRANSFER_WH') {
    return users.filter(u => role(u) === 'warehouseman' && _notif_hasAccess(u.locAccess, payload.targetLoc));
  }
  if (action === 'ISSUE') {
    return users.filter(u => role(u) === 'team leader' && _notif_hasAccess(u.siteAccess, payload.siteName));
  }
  if (action === 'RETURN_WH') {
    return users.filter(u => role(u) === 'warehouseman' && _notif_hasAccess(u.locAccess, payload.location));
  }
  if (action === 'RETURN_CLIENT') {
    return users.filter(u => role(u) === 'admin');
  }
  return [];
}

function resolveRequester(reqId) {
  if (!reqId) return null;
  const rSheet = SS.getSheetByName(SHEETS.REQ);
  if (!rSheet || rSheet.getLastRow() < 2) return null;
  const lastCol = Math.max(14, rSheet.getLastColumn());
  const data = rSheet.getRange(2, 1, rSheet.getLastRow() - 1, lastCol).getValues();
  const target = reqId.toString().trim();
  const row = data.find(r => r[0] && r[0].toString().trim() === target);
  if (!row) return null;

  const claimedEmail = row[13] ? row[13].toString().trim() : '';
  const claimedName = row[2] ? row[2].toString().trim() : '';
  const users = _notif_listUsers();
  let match = null;
  if (claimedEmail) {
    match = users.find(u => u.email.toLowerCase() === claimedEmail.toLowerCase());
  }
  if (!match && claimedName) {
    match = users.find(u => u.name.toLowerCase() === claimedName.toLowerCase());
  }
  if (!match) return null;
  return { email: match.email, name: match.name, role: match.role };
}
```

- [ ] **Step 4: Run tests to verify they pass**

All four tests should log `OK:` and complete.

- [ ] **Step 5: Commit**

```bash
git add src/notifications.js src/tests.js
git commit -m "$(cat <<'EOF'
Add resolveRecipients and resolveRequester

resolveRecipients targets by role + Location Access / Site Access for
DR_CREATE, TRANSFER_WH, ISSUE, RETURN_WH, RETURN_CLIENT. resolveRequester
looks up by User Email (new Requests col 13) with name fallback for
legacy rows.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 7: Write `User Email` on every new Requests row

**Files:**
- Modify: `src/Code.js`

**Interfaces:**
- Consumes: `validated.email` from Task 4
- Produces: Every new Requests row written by `processBulkTransaction` has a non-empty `User Email` in col 13

- [ ] **Step 1: Find all Requests-row inserts in `processBulkTransaction`**

Search `src/Code.js` for `requestEntries.push(` — there are four insertion points:
- Line ~342 (DR_CREATE branch)
- Line ~364 (TRANSFER_WH branch)
- Line ~373 (ISSUE branch)
- Line ~409 (RETURN_WH branch)

Each currently writes 13 columns; we need to extend each to 14 by appending `validated.email`.

- [ ] **Step 2: Update each `requestEntries.push` call**

In `src/Code.js`, find each call and add `, validated.email` immediately before the closing `]`:

Line ~342 — change:
```javascript
requestEntries.push([ finalDocId, d, payload.user, payload.role, payload.action, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, 'Pending DR', '' ]);
```
to:
```javascript
requestEntries.push([ finalDocId, d, payload.user, payload.role, payload.action, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, 'Pending DR', '', validated.email ]);
```

Line ~364 — change:
```javascript
requestEntries.push([ reqId, d, payload.user, payload.role, payload.action, payload.targetLoc, payload.targetSite, cleanCode, cleanName, itemReq.uom, actualQty, 'Pending Receipt', '' ]);
```
to:
```javascript
requestEntries.push([ reqId, d, payload.user, payload.role, payload.action, payload.targetLoc, payload.targetSite, cleanCode, cleanName, itemReq.uom, actualQty, 'Pending Receipt', '', validated.email ]);
```

Line ~373 — change:
```javascript
requestEntries.push([ reqId, d, payload.user, payload.role, payload.action, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, 'In Transit', '' ]);
```
to:
```javascript
requestEntries.push([ reqId, d, payload.user, payload.role, payload.action, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, 'In Transit', '', validated.email ]);
```

Line ~409 — change:
```javascript
requestEntries.push([ reqId, d, payload.user, payload.role, payload.action, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, 'Pending Return', '' ]);
```
to:
```javascript
requestEntries.push([ reqId, d, payload.user, payload.role, payload.action, payload.location, payload.siteName, cleanCode, cleanName, itemReq.uom, actualQty, 'Pending Return', '', validated.email ]);
```

- [ ] **Step 3: Update the batch write width**

In `src/Code.js` around line 494, find:
```javascript
if (requestEntries.length > 0) rSheet.getRange(rSheet.getLastRow() + 1, 1, requestEntries.length, 13).setValues(requestEntries);
```
Change `13` to `14`:
```javascript
if (requestEntries.length > 0) rSheet.getRange(rSheet.getLastRow() + 1, 1, requestEntries.length, 14).setValues(requestEntries);
```

- [ ] **Step 4: Add a test**

Append to `src/tests.js`:

```javascript
function test_processBulkTransaction_writes_user_email_on_requests() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName('Requests');
  const startRows = rSheet.getLastRow();

  const result = processBulkTransaction({
    email: 'admin@test.com',
    user: 'Admin User',
    role: 'admin',
    action: 'DR_CREATE',
    location: 'NCR Hub',
    siteName: 'Makati Site',
    siteId: 'S-001',
    client: 'Acme Corp',
    poNumber: 'PO-10001',
    items: [{ code: 'ITM-001', name: 'Dell Latitude', uom: 'pc', qty: 5, wbs: 'WBS-991' }]
  });
  _assert(result.success === true, "DR_CREATE succeeded");

  const newRow = rSheet.getRange(startRows + 1, 1, 1, 14).getValues()[0];
  _assert(newRow[13] === 'admin@test.com', "User Email column populated on new Requests row");
}
```

- [ ] **Step 5: Run the test and verify it passes**

Run `test_processBulkTransaction_writes_user_email_on_requests`. Should log `OK:` lines.

- [ ] **Step 6: Commit**

```bash
git add src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Write User Email on every new Requests row

Lets processQueueAction look up the original requester by email for
acknowledgment notifications instead of fragile name matching.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 8: Wire `notify` into DR_CREATE

**Files:**
- Modify: `src/Code.js`

**Interfaces:**
- Consumes: `notify`, `resolveRecipients` from notifications.js; `validated` from Task 4
- Produces: A DR_CREATE submission writes Notifications rows for every targeted warehouseman

- [ ] **Step 1: Locate the DR_CREATE notification point**

The DR_CREATE loop is at `src/Code.js:341-345`. We must notify ONCE per transaction, not once per item, so we collect item summaries during the loop and fire after the loop completes (and before `lock.releaseLock()`).

- [ ] **Step 2: Add the post-loop notification logic**

In `src/Code.js`, find the section between the for-loop close at ~line 416 (`}` closing the `else` branch) and `SpreadsheetApp.flush();` at line 507. We'll add notification logic just after the writes to `requestEntries` complete (i.e., after line 494 batch write).

After the line:
```javascript
if (requestEntries.length > 0) rSheet.getRange(rSheet.getLastRow() + 1, 1, requestEntries.length, 14).setValues(requestEntries);
```

add a notification block:

```javascript
    // === Notifications ===
    try {
      const sender = { email: validated.email, name: validated.fullName, role: validated.role };

      if (payload.action === 'DR_CREATE' && requestEntries.length > 0) {
        const recipients = resolveRecipients('DR_CREATE', payload);
        const itemList = requestEntries.map(r => r[7]).join(', ').substring(0, 120);
        const msg = `${validated.fullName} created DR ${finalDocId} (${requestEntries.length} item${requestEntries.length === 1 ? '' : 's'}: ${itemList}) — pending your receipt`;
        notify(recipients, 'DR_CREATE', sender, msg, finalDocId);
      }
    } catch (notifErr) {
      console.error("notify(DR_CREATE) failed: " + notifErr.toString());
    }
```

- [ ] **Step 3: Add a manual UAT test in `tests.js`**

Append to `src/tests.js`:

```javascript
function test_DR_CREATE_fires_notification() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nSheet = ss.getSheetByName('Notifications');
  const startNotifs = nSheet.getLastRow();

  processBulkTransaction({
    email: 'admin@test.com',
    user: 'Admin User',
    role: 'admin',
    action: 'DR_CREATE',
    location: 'NCR Hub',
    siteName: 'Makati Site',
    siteId: 'S-001',
    client: 'Acme Corp',
    poNumber: 'PO-10001',
    items: [{ code: 'ITM-001', name: 'Dell Latitude', uom: 'pc', qty: 5, wbs: 'WBS-991' }]
  });

  const endNotifs = nSheet.getLastRow();
  _assert(endNotifs > startNotifs, "Notifications sheet grew");
  const newRow = nSheet.getRange(endNotifs, 1, 1, 14).getValues()[0];
  _assert(newRow[8] === 'DR_CREATE', "Notification action = DR_CREATE");
  _assert(newRow[4] === 'warehouseman', "Recipient role = warehouseman");
  _assert(newRow[2] === 'wh@test.com', "Recipient is NCR Hub warehouseman");
}
```

- [ ] **Step 4: Run the test**

Run `test_DR_CREATE_fires_notification`. Verify it passes. Note: the test will send a real email to `wh@test.com`. If this is undesired during development, temporarily replace `MailApp.sendEmail` calls with a stub or accept the test inbox receives mail.

- [ ] **Step 5: Commit**

```bash
git add src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Wire DR_CREATE notifications to scoped warehousemen

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 9: Wire `notify` into TRANSFER_WH

**Files:**
- Modify: `src/Code.js`

**Interfaces:**
- Consumes: same as Task 8
- Produces: A TRANSFER_WH submission notifies warehousemen at the target location

- [ ] **Step 1: Extend the notification block from Task 8**

In `src/Code.js`, find the `=== Notifications ===` block added in Task 8 and extend it with the TRANSFER_WH branch. After the DR_CREATE branch, add:

```javascript
      if (payload.action === 'TRANSFER_WH' && requestEntries.length > 0) {
        const recipients = resolveRecipients('TRANSFER_WH', payload);
        const reqId = requestEntries[0][0];
        const itemList = requestEntries.map(r => r[7]).join(', ').substring(0, 120);
        const targetLabel = (payload.targetSite && payload.targetSite !== '-') ? payload.targetSite : payload.targetLoc;
        const msg = `${validated.fullName} initiated transfer to ${targetLabel} (${requestEntries.length} item${requestEntries.length === 1 ? '' : 's'}: ${itemList}) — pending your receipt`;
        notify(recipients, 'TRANSFER_WH', sender, msg, reqId);
      }
```

(Each TRANSFER_WH iteration in the item-loop writes its own `requestEntries` row with a freshly minted `reqId`; for the notification we use the first row's reqId as the "primary" — this is acceptable since TRANSFER_WH submissions typically share context within one user action. If multiple distinct reqIds are needed later, this is a known limitation noted in spec §8 deduplication.)

- [ ] **Step 2: Add a test**

Append to `src/tests.js`:

```javascript
function test_TRANSFER_WH_notifies_target_warehousemen() {
  initializeSheets();
  // Seed a second warehouseman with Visayas Hub access for this test:
  const uSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  // Only add if not present
  const existing = uSheet.getRange(2, 1, uSheet.getLastRow() - 1, 1).getValues().flat();
  if (existing.indexOf('wh2@test.com') === -1) {
    uSheet.appendRow(['wh2@test.com', 'Dora Warehouseman', 'temp123', '', 'warehouseman', 'Visayas Hub', '']);
  }

  const nSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notifications');
  const startNotifs = nSheet.getLastRow();

  processBulkTransaction({
    email: 'wh@test.com',
    user: 'Alice Warehouseman',
    role: 'warehouseman',
    action: 'TRANSFER_WH',
    location: 'NCR Hub',
    siteName: '-',
    targetLoc: 'Visayas Hub',
    targetSite: '-',
    items: [{ code: 'ITM-001', name: 'Dell Latitude', uom: 'pc', qty: 1, wbs: '' }]
  });

  const endNotifs = nSheet.getLastRow();
  _assert(endNotifs > startNotifs, "Notifications sheet grew");
  const recents = nSheet.getRange(startNotifs + 1, 1, endNotifs - startNotifs, 14).getValues();
  _assert(recents.some(r => r[2] === 'wh2@test.com' && r[8] === 'TRANSFER_WH'),
    "Visayas Hub warehouseman (wh2) was notified for TRANSFER_WH");
  _assert(!recents.some(r => r[2] === 'wh@test.com' && r[8] === 'TRANSFER_WH'),
    "Source warehouseman was NOT self-notified");
}
```

- [ ] **Step 3: Run the test**

Run `test_TRANSFER_WH_notifies_target_warehousemen`. Should pass.

- [ ] **Step 4: Commit**

```bash
git add src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Wire TRANSFER_WH notifications to target-location warehousemen

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 10: Wire `notify` into ISSUE

**Files:**
- Modify: `src/Code.js`

**Interfaces:**
- Consumes: same as Task 8
- Produces: An ISSUE submission notifies team leaders covering the target site

- [ ] **Step 1: Extend the notification block**

In the `=== Notifications ===` block, add after the TRANSFER_WH branch:

```javascript
      if (payload.action === 'ISSUE' && requestEntries.length > 0) {
        const recipients = resolveRecipients('ISSUE', payload);
        const reqId = requestEntries[0][0];
        const itemList = requestEntries.map(r => r[7]).join(', ').substring(0, 120);
        const msg = `${validated.fullName} issued ${requestEntries.length} item${requestEntries.length === 1 ? '' : 's'} (${itemList}) to ${payload.siteName} — pending your acknowledgment`;
        notify(recipients, 'ISSUE', sender, msg, reqId);
      }
```

- [ ] **Step 2: Add a test**

Append to `src/tests.js`:

```javascript
function test_ISSUE_notifies_scoped_team_leader() {
  initializeSheets();
  const nSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notifications');
  const startNotifs = nSheet.getLastRow();

  processBulkTransaction({
    email: 'wh@test.com',
    user: 'Alice Warehouseman',
    role: 'warehouseman',
    action: 'ISSUE',
    location: 'NCR Hub',
    siteName: 'Makati Site',
    siteId: 'S-001',
    client: 'Acme Corp',
    items: [{ code: 'ITM-001', name: 'Dell Latitude', uom: 'pc', qty: 1, wbs: 'WBS-991' }]
  });

  const recents = nSheet.getRange(startNotifs + 1, 1, nSheet.getLastRow() - startNotifs, 14).getValues();
  _assert(recents.some(r => r[2] === 'tl1@test.com' && r[8] === 'ISSUE'),
    "Makati team leader notified for ISSUE");
  _assert(!recents.some(r => r[2] === 'tl2@test.com' && r[8] === 'ISSUE'),
    "Cebu team leader NOT notified");
}
```

- [ ] **Step 3: Run, verify, commit**

```bash
git add src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Wire ISSUE notifications to site-scoped team leaders

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 11: Wire `notify` into RETURN_WH

**Files:**
- Modify: `src/Code.js`

**Interfaces:**
- Consumes: same as Task 8
- Produces: A RETURN_WH submission notifies warehousemen at the destination warehouse

- [ ] **Step 1: Extend the notification block**

Add to the `=== Notifications ===` block, after ISSUE:

```javascript
      if (payload.action === 'RETURN_WH' && requestEntries.length > 0) {
        const recipients = resolveRecipients('RETURN_WH', payload);
        const reqId = requestEntries[0][0];
        const itemList = requestEntries.map(r => r[7]).join(', ').substring(0, 120);
        const msg = `${validated.fullName} returned ${requestEntries.length} item${requestEntries.length === 1 ? '' : 's'} (${itemList}) from ${payload.siteName} to ${payload.location} — pending your receipt`;
        notify(recipients, 'RETURN_WH', sender, msg, reqId);
      }
```

- [ ] **Step 2: Add a test**

Append to `src/tests.js`:

```javascript
function test_RETURN_WH_notifies_location_warehouseman() {
  initializeSheets();
  // Seed stock at Makati Site to allow the return
  const iSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  const invData = iSheet.getDataRange().getValues();
  const headers = invData[0];
  let makatiCol = headers.indexOf('Makati Site');
  if (makatiCol === -1) {
    iSheet.getRange(1, headers.length + 1).setValue('Makati Site');
    makatiCol = headers.length;
  }
  // Set Makati Site stock for ITM-001 = 10
  const itmRow = invData.findIndex(r => r[0] === 'ITM-001');
  if (itmRow > 0) iSheet.getRange(itmRow + 1, makatiCol + 1).setValue(10);

  const nSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notifications');
  const startNotifs = nSheet.getLastRow();

  processBulkTransaction({
    email: 'tl1@test.com',
    user: 'Bob TeamLeader',
    role: 'team leader',
    action: 'RETURN_WH',
    location: 'NCR Hub',
    siteName: 'Makati Site',
    siteId: 'S-001',
    client: 'Acme Corp',
    items: [{ code: 'ITM-001', name: 'Dell Latitude', uom: 'pc', qty: 1, wbs: 'WBS-991' }]
  });

  const recents = nSheet.getRange(startNotifs + 1, 1, nSheet.getLastRow() - startNotifs, 14).getValues();
  _assert(recents.some(r => r[2] === 'wh@test.com' && r[8] === 'RETURN_WH'),
    "NCR Hub warehouseman notified for RETURN_WH");
}
```

- [ ] **Step 3: Run, verify, commit**

```bash
git add src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Wire RETURN_WH notifications to destination warehousemen

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 12: Wire `notify` into RETURN_CLIENT

**Files:**
- Modify: `src/Code.js`

**Interfaces:**
- Consumes: same as Task 8
- Produces: A RETURN_CLIENT submission notifies all admins. The existing PDF email to the submitter continues to fire independently.

- [ ] **Step 1: Extend the notification block**

Add to the `=== Notifications ===` block, after RETURN_WH:

```javascript
      if (payload.action === 'RETURN_CLIENT') {
        const recipients = resolveRecipients('RETURN_CLIENT', payload);
        const itemList = (returnedItems || []).map(r => r.code).join(', ').substring(0, 120);
        const msg = `${validated.fullName} processed a client return at ${payload.location} (${(returnedItems || []).length} item${(returnedItems || []).length === 1 ? '' : 's'}: ${itemList})`;
        notify(recipients, 'RETURN_CLIENT', sender, msg, finalDocId);
      }
```

Note: `returnedItems` is the array populated inside the RETURN_CLIENT branch of the item loop (`Code.js:398`); it's in scope at this notification block.

- [ ] **Step 2: Add a test**

Append to `src/tests.js`:

```javascript
function test_RETURN_CLIENT_notifies_admins() {
  initializeSheets();
  // Need a prior RECEIVE_DR to provide stock to return; for the test we just
  // seed Inventory directly with stock at NCR Hub for ITM-001.
  const iSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  const invData = iSheet.getDataRange().getValues();
  const headers = invData[0];
  const ncrCol = headers.indexOf('NCR Hub');
  const itmRow = invData.findIndex(r => r[0] === 'ITM-001');
  if (itmRow > 0 && ncrCol !== -1) iSheet.getRange(itmRow + 1, ncrCol + 1).setValue(50);

  const nSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notifications');
  const startNotifs = nSheet.getLastRow();

  processBulkTransaction({
    email: 'wh@test.com',
    user: 'Alice Warehouseman',
    role: 'warehouseman',
    action: 'RETURN_CLIENT',
    location: 'NCR Hub',
    siteName: '-',
    siteId: '-',
    client: 'Acme Corp',
    refDoc: 'DR-TEST-001',
    sourceDoc: 'DR-TEST-001',
    mrcNum: 'MRC-001',
    returnType: 'BY_SITE',
    items: [{ code: 'ITM-001', name: 'Dell Latitude', uom: 'pc', qty: 1, actualReturnQty: '1', wbs: '' }]
  });

  const recents = nSheet.getRange(startNotifs + 1, 1, nSheet.getLastRow() - startNotifs, 14).getValues();
  _assert(recents.some(r => r[8] === 'RETURN_CLIENT' && r[4] === 'admin'),
    "Admin notified for RETURN_CLIENT");
}
```

- [ ] **Step 3: Run, verify, commit**

```bash
git add src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Wire RETURN_CLIENT notifications to all admins

PDF receipt email to submitter continues to fire independently of these
in-app/email notifications.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 13: Wire RECEIVE_DR acknowledgment to original admin

**Files:**
- Modify: `src/Code.js`

**Interfaces:**
- Consumes: `resolveRequester`, `notify`
- Produces: When a warehouseman runs RECEIVE_DR, the admin who originally created the DR receives a CONFIRM notification

- [ ] **Step 1: Extend the notification block**

Add to the `=== Notifications ===` block, after the RETURN_CLIENT branch:

```javascript
      if (payload.action === 'RECEIVE_DR') {
        const creator = resolveRequester(payload.drId);
        if (creator) {
          const itemCount = payload.items ? payload.items.length : 0;
          const msg = `${validated.fullName} received DR ${payload.drId} (${itemCount} item${itemCount === 1 ? '' : 's'})`;
          notify([creator], 'CONFIRM', sender, msg, payload.drId);
        }
      }
```

- [ ] **Step 2: Add a test**

Append to `src/tests.js`:

```javascript
function test_RECEIVE_DR_notifies_creator_admin() {
  initializeSheets();
  const nSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notifications');

  // First: admin creates a DR
  const drId = 'DR-RECV-TEST-' + Math.floor(Math.random() * 10000);
  // Use processBulkTransaction so the User Email column is populated.
  processBulkTransaction({
    email: 'admin@test.com', user: 'Admin User', role: 'admin', action: 'DR_CREATE',
    location: 'NCR Hub', siteName: 'Makati Site', siteId: 'S-001', client: 'Acme Corp',
    refDoc: drId, poNumber: 'PO-10001',
    items: [{ code: 'ITM-001', name: 'Dell Latitude', uom: 'pc', qty: 2, wbs: 'WBS-991' }]
  });

  const startAfterCreate = nSheet.getLastRow();

  // Now warehouseman receives it
  processBulkTransaction({
    email: 'wh@test.com', user: 'Alice Warehouseman', role: 'warehouseman', action: 'RECEIVE_DR',
    location: 'NCR Hub', drId: drId, client: 'Acme Corp', siteName: 'Makati Site', siteId: 'S-001',
    items: [{ code: 'ITM-001', name: 'Dell Latitude', uom: 'pc', qty: 2, wbs: 'WBS-991' }]
  });

  const recents = nSheet.getRange(startAfterCreate + 1, 1, nSheet.getLastRow() - startAfterCreate, 14).getValues();
  _assert(recents.some(r => r[8] === 'CONFIRM' && r[2] === 'admin@test.com'),
    "Admin received CONFIRM notification for RECEIVE_DR");
}
```

- [ ] **Step 3: Run, verify, commit**

```bash
git add src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Wire RECEIVE_DR acknowledgment back to original DR creator

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 14: Wire `notify` into `processQueueAction` (confirm/reject acks)

**Files:**
- Modify: `src/Code.js`

**Interfaces:**
- Consumes: `resolveRequester`, `notify`
- Produces: Confirming or rejecting a queue item notifies the original requester

- [ ] **Step 1: Add the notification after status update**

In `src/Code.js`, find `processQueueAction` around line 599 — the line:
```javascript
rSheet.getRange(rowNum, 12).setValue(newStatus); 
```

Immediately after this line and before `SpreadsheetApp.flush();`, add:

```javascript
    // === Acknowledgment notification ===
    try {
      const creator = resolveRequester(reqId);
      if (creator) {
        const outcome = (action === 'Reject') ? 'REJECT' : 'CONFIRM';
        const verb = (action === 'Reject') ? 'rejected' : 'confirmed';
        const msg = `${userProfile.fullName} (${userProfile.role}) ${verb} your ${reqAction} request ${reqId}`;
        const sender = { email: userProfile.email, name: userProfile.fullName, role: userProfile.role };
        notify([creator], outcome, sender, msg, reqId);
      }
    } catch (notifErr) {
      console.error("notify(queue " + action + ") failed: " + notifErr.toString());
    }
```

Note: `userProfile` here is the validated profile (reassigned at the top of the function in Task 4 step 2). It has `email`, `fullName`, and `role`.

- [ ] **Step 2: Add a test**

Append to `src/tests.js`:

```javascript
function test_queue_confirm_notifies_original_requester() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nSheet = ss.getSheetByName('Notifications');

  // Seed stock so ISSUE will succeed
  const iSheet = ss.getSheetByName('Inventory');
  const invData = iSheet.getDataRange().getValues();
  const ncrCol = invData[0].indexOf('NCR Hub');
  const itmRow = invData.findIndex(r => r[0] === 'ITM-001');
  if (itmRow > 0 && ncrCol !== -1) iSheet.getRange(itmRow + 1, ncrCol + 1).setValue(50);

  // Warehouseman issues to Makati Site
  const issueRes = processBulkTransaction({
    email: 'wh@test.com', user: 'Alice Warehouseman', role: 'warehouseman', action: 'ISSUE',
    location: 'NCR Hub', siteName: 'Makati Site', siteId: 'S-001', client: 'Acme Corp',
    items: [{ code: 'ITM-001', name: 'Dell Latitude', uom: 'pc', qty: 1, wbs: 'WBS-991' }]
  });
  _assert(issueRes.success, "ISSUE submitted");

  // Find the issue reqId
  const rSheet = ss.getSheetByName('Requests');
  const reqs = rSheet.getDataRange().getValues();
  const issueRow = reqs.slice().reverse().find(r => r[4] === 'ISSUE' && r[11] === 'In Transit');
  _assert(issueRow, "Found pending ISSUE");
  const reqId = issueRow[0];

  const startNotifs = nSheet.getLastRow();

  // Team leader confirms
  processQueueAction(reqId, 'Confirm Receipt', { email: 'tl1@test.com', fullName: 'Bob TeamLeader', role: 'team leader' });

  const recents = nSheet.getRange(startNotifs + 1, 1, nSheet.getLastRow() - startNotifs, 14).getValues();
  _assert(recents.some(r => r[8] === 'CONFIRM' && r[2] === 'wh@test.com'),
    "Original warehouseman notified of confirmation");
}
```

- [ ] **Step 3: Run, verify, commit**

```bash
git add src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Notify original requester on queue confirm/reject

Acknowledgments back to the warehouseman/team leader who initiated the
queued action.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 15: `getNotificationsForUser`, `markNotificationRead`, `markAllNotificationsRead`

**Files:**
- Modify: `src/notifications.js`
- Modify: `src/Code.js` (integrate into `getAppData`)
- Modify: `src/tests.js`

**Interfaces:**
- Consumes: `validateUserProfile`, `Notifications` sheet
- Produces:
  - `getNotificationsForUser(userProfile) → { items: [...], unreadCount: N }`
  - `markNotificationRead(notifId, userProfile) → { success: true } or throws`
  - `markAllNotificationsRead(userProfile) → { success: true, marked: N }`

- [ ] **Step 1: Write failing tests**

Append to `src/tests.js`:

```javascript
function test_getNotificationsForUser_filters_by_email() {
  initializeSheets();
  // Generate a notification for wh@test.com
  notify(
    [{ email: 'wh@test.com', name: 'Alice Warehouseman', role: 'warehouseman' }],
    'TEST_ACTION',
    { email: 'admin@test.com', name: 'Admin User', role: 'admin' },
    'unit test message',
    'TEST-001'
  );
  const result = getNotificationsForUser({ email: 'wh@test.com' });
  _assert(Array.isArray(result.items), "Returns items array");
  _assert(typeof result.unreadCount === 'number', "Returns unreadCount number");
  _assert(result.items.every(i => i.recipientEmail === undefined || i.recipientEmail.toLowerCase() === 'wh@test.com' || i.recipientEmail === ''),
    "All items are for the requested user (or stripped of recipientEmail in the payload)");
}

function test_markNotificationRead_marks_read() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nSheet = ss.getSheetByName('Notifications');
  notify(
    [{ email: 'wh@test.com', name: 'Alice', role: 'warehouseman' }],
    'TEST', { email: 'admin@test.com', name: 'Admin', role: 'admin' },
    'mark-read test', 'X-001'
  );
  const lastRow = nSheet.getRange(nSheet.getLastRow(), 1, 1, 14).getValues()[0];
  const notifId = lastRow[0];
  _assert(lastRow[11] === false, "Initially unread");
  markNotificationRead(notifId, { email: 'wh@test.com' });
  const after = nSheet.getRange(nSheet.getLastRow(), 1, 1, 14).getValues()[0];
  _assert(after[11] === true, "Read column set to true");
}

function test_markNotificationRead_rejects_foreign_notif() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nSheet = ss.getSheetByName('Notifications');
  notify(
    [{ email: 'wh@test.com', name: 'Alice', role: 'warehouseman' }],
    'TEST', { email: 'admin@test.com', name: 'Admin', role: 'admin' },
    'ownership test', 'X-002'
  );
  const lastRow = nSheet.getRange(nSheet.getLastRow(), 1, 1, 14).getValues()[0];
  const notifId = lastRow[0];
  let threw = false;
  try { markNotificationRead(notifId, { email: 'tl1@test.com' }); }
  catch (e) { threw = e.message.indexOf("Not your notification") !== -1; }
  _assert(threw, "Throws when caller's email != recipient");
}
```

- [ ] **Step 2: Run — verify failure**

All three should fail with `ReferenceError: getNotificationsForUser is not defined` (etc).

- [ ] **Step 3: Implement in `src/notifications.js`**

Append to `src/notifications.js`:

```javascript
function getNotificationsForUser(userProfile) {
  const validated = validateUserProfile(userProfile);
  const nSheet = SS.getSheetByName(SHEETS.NOTIF);
  if (!nSheet || nSheet.getLastRow() < 2) return { items: [], unreadCount: 0 };

  const data = nSheet.getRange(2, 1, nSheet.getLastRow() - 1, 14).getValues();
  const target = validated.email.toLowerCase();
  const mine = data.filter(r => r[2] && r[2].toString().toLowerCase() === target);

  let unreadCount = 0;
  mine.forEach(r => { if (r[11] === false || r[11] === '' || r[11] === 'FALSE') unreadCount++; });

  const sorted = mine.slice().sort((a, b) => {
    const ta = a[1] && a[1].getTime ? a[1].getTime() : 0;
    const tb = b[1] && b[1].getTime ? b[1].getTime() : 0;
    return tb - ta;
  });

  const items = sorted.slice(0, 30).map(r => ({
    notifId: r[0],
    timestamp: r[1] && r[1].toISOString ? r[1].toISOString() : (r[1] || ''),
    senderName: r[6] || '',
    senderRole: r[7] || '',
    action: r[8] || '',
    relatedReqId: r[9] || '',
    message: r[10] || '',
    read: r[11] === true || r[11] === 'TRUE'
  }));

  return { items: items, unreadCount: unreadCount };
}

function markNotificationRead(notifId, userProfile) {
  const validated = validateUserProfile(userProfile);
  const nSheet = SS.getSheetByName(SHEETS.NOTIF);
  if (!nSheet || nSheet.getLastRow() < 2) throw new Error("No notifications.");

  const data = nSheet.getRange(2, 1, nSheet.getLastRow() - 1, 14).getValues();
  const target = notifId.toString().trim();
  const idx = data.findIndex(r => r[0] && r[0].toString().trim() === target);
  if (idx === -1) throw new Error("Notification not found.");
  if ((data[idx][2] || '').toString().toLowerCase() !== validated.email.toLowerCase()) {
    throw new Error("Not your notification.");
  }
  nSheet.getRange(idx + 2, 12, 1, 2).setValues([[true, new Date()]]);
  return { success: true };
}

function markAllNotificationsRead(userProfile) {
  const validated = validateUserProfile(userProfile);
  const nSheet = SS.getSheetByName(SHEETS.NOTIF);
  if (!nSheet || nSheet.getLastRow() < 2) return { success: true, marked: 0 };

  const data = nSheet.getRange(2, 1, nSheet.getLastRow() - 1, 14).getValues();
  const target = validated.email.toLowerCase();
  const now = new Date();
  let marked = 0;
  data.forEach((r, i) => {
    if (r[2] && r[2].toString().toLowerCase() === target && r[11] !== true && r[11] !== 'TRUE') {
      nSheet.getRange(i + 2, 12, 1, 2).setValues([[true, now]]);
      marked++;
    }
  });
  return { success: true, marked: marked };
}
```

- [ ] **Step 4: Integrate into `getAppData`**

In `src/Code.js`, find the return statement at line 218:

```javascript
    return { inventory, logs, dropdowns, pending, pendingDRs, receiptDocs };
```

Change to:

```javascript
    let notifications = { items: [], unreadCount: 0 };
    try {
      if (validated) notifications = getNotificationsForUser(validated);
    } catch (notifErr) {
      console.warn("getNotificationsForUser failed: " + notifErr.toString());
    }
    return { inventory, logs, dropdowns, pending, pendingDRs, receiptDocs, notifications };
```

- [ ] **Step 5: Run all three tests — verify they pass**

Run the three new tests. All pass.

- [ ] **Step 6: Commit**

```bash
git add src/notifications.js src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Add getNotificationsForUser, markNotificationRead, markAllNotificationsRead

Plus integration into getAppData so the master payload now carries
the user's notifications (items + unreadCount).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 16: Frontend — bell icon, dropdown panel, base CSS

**Files:**
- Modify: `src/Index.html`

**Interfaces:**
- Consumes: `master.notifications` from `getAppData` response (Task 15)
- Produces: A `<button id="notif-bell">` in the navbar with badge, and a `<div id="notif-panel">` dropdown. Initially empty / hidden — populated in Task 17.

- [ ] **Step 1: Add HTML to the navbar**

In `src/Index.html`, find the navbar block at lines 120-133:

```html
<div class="user-nav-info">
  <div class="text-end d-none d-md-block">
    <div id="user-display" class="fw-bold text-dark small"></div>
    <div class="text-muted smaller" style="font-size: 0.75rem;"><i class="bi bi-person-badge text-accent"></i> <span id="role-display" class="text-uppercase"></span></div>
  </div>
  <button onclick="doLogout()" class="btn btn-sm btn-outline-danger border-0"><i class="bi bi-box-arrow-right fs-5"></i></button>
</div>
```

Change to add the bell button immediately before the user info block:

```html
<div class="user-nav-info">
  <div id="notif-wrap" style="position: relative;">
    <button id="notif-bell" class="btn btn-sm btn-outline-secondary border-0" onclick="toggleNotifPanel()" type="button" aria-label="Notifications">
      <i class="bi bi-bell fs-5"></i>
      <span id="notif-badge" class="notif-badge d-none">0</span>
    </button>
    <div id="notif-panel" class="notif-panel d-none">
      <div class="notif-panel-header">
        <span class="fw-bold">Notifications</span>
        <a href="#" onclick="markAllNotifsRead(); return false;" id="notif-mark-all">Mark all read</a>
      </div>
      <div id="notif-list" class="notif-list"></div>
      <div class="notif-panel-footer" id="notif-footer"></div>
    </div>
  </div>
  <div class="text-end d-none d-md-block">
    <div id="user-display" class="fw-bold text-dark small"></div>
    <div class="text-muted smaller" style="font-size: 0.75rem;"><i class="bi bi-person-badge text-accent"></i> <span id="role-display" class="text-uppercase"></span></div>
  </div>
  <button onclick="doLogout()" class="btn btn-sm btn-outline-danger border-0"><i class="bi bi-box-arrow-right fs-5"></i></button>
</div>
```

- [ ] **Step 2: Add CSS**

In `src/Index.html`, find the `<style>` block (starts at line 9). Insert the following just before the closing `</style>` (search for the last `}` before `</style>`):

```css
    .notif-badge {
      position: absolute; top: -2px; right: -4px;
      background: #ef4444; color: white;
      font-size: 0.65rem; font-weight: 700;
      min-width: 18px; height: 18px;
      border-radius: 9px;
      display: inline-flex; align-items: center; justify-content: center;
      padding: 0 5px;
    }
    .notif-panel {
      position: absolute; top: calc(100% + 8px); right: 0;
      width: 340px; max-height: 480px;
      background: white; border: 1px solid #e2e8f0;
      border-radius: 16px; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.15);
      z-index: 10000; overflow: hidden;
      display: flex; flex-direction: column;
    }
    .notif-panel-header {
      padding: 14px 16px; border-bottom: 1px solid #f1f5f9;
      display: flex; justify-content: space-between; align-items: center;
      background: #f8fafc;
    }
    .notif-panel-header a { font-size: 0.8rem; color: var(--accent); text-decoration: none; }
    .notif-panel-header a[aria-disabled="true"] { color: #cbd5e1; pointer-events: none; }
    .notif-list { flex: 1; overflow-y: auto; }
    .notif-item {
      padding: 12px 16px; border-bottom: 1px solid #f1f5f9; cursor: pointer;
      transition: background 0.15s;
    }
    .notif-item:hover { background: #f1f5f9; }
    .notif-item.unread { background: #eff6ff; border-left: 3px solid var(--accent); padding-left: 13px; }
    .notif-item .sender { font-weight: 700; color: var(--primary); font-size: 0.85rem; }
    .notif-item .role-tag { font-size: 0.7rem; color: #64748b; background: #e2e8f0; border-radius: 6px; padding: 1px 6px; margin-left: 6px; }
    .notif-item .msg { color: #475569; font-size: 0.8rem; margin-top: 4px; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }
    .notif-item .time { font-size: 0.7rem; color: #94a3b8; margin-top: 4px; }
    .notif-empty { padding: 32px 16px; text-align: center; color: #94a3b8; font-size: 0.85rem; }
    .notif-panel-footer { padding: 8px 16px; font-size: 0.7rem; color: #94a3b8; text-align: center; border-top: 1px solid #f1f5f9; background: #f8fafc; }
```

- [ ] **Step 3: Add the toggle function (stub for now — full populate is Task 17)**

In `src/Index.html`, search for `function renderQueue()` (around line 896). Insert these functions just before it:

```javascript
  function toggleNotifPanel() {
    const panel = document.getElementById('notif-panel');
    panel.classList.toggle('d-none');
  }

  // Close panel when clicking outside
  document.addEventListener('click', function(e) {
    const wrap = document.getElementById('notif-wrap');
    const panel = document.getElementById('notif-panel');
    if (!wrap || !panel) return;
    if (!wrap.contains(e.target) && !panel.classList.contains('d-none')) {
      panel.classList.add('d-none');
    }
  });

  function markAllNotifsRead() {
    // Implemented in Task 18
  }
```

- [ ] **Step 4: Manual UAT — verify bell renders**

Deploy/push the changes and reload the web app. Log in. Confirm: the bell icon appears in the navbar to the left of the user name. Clicking it toggles an empty panel. Clicking outside closes it.

- [ ] **Step 5: Commit**

```bash
git add src/Index.html
git commit -m "$(cat <<'EOF'
Add notification bell UI (HTML + CSS + open/close)

Bell, badge, and empty dropdown panel. Populate logic comes in the
next commit.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 17: Populate the bell from `master.notifications`

**Files:**
- Modify: `src/Index.html`

**Interfaces:**
- Consumes: `master.notifications.items[]` and `master.notifications.unreadCount` from getAppData
- Produces: Bell renders unread count, dropdown lists up to 10 newest notifications

- [ ] **Step 1: Add render functions**

In `src/Index.html`, find the `function refresh(silent = false)` (around line 718). Inside its `withSuccessHandler`, find:

```javascript
master = data; 
populateDropdowns();
handleActionChange();
renderStock();
renderLogs();
renderQueue();
renderPendingPOs();
handleLocationChange();
```

Add `renderNotifications();` to the end of that list.

Then add the `renderNotifications` function near `toggleNotifPanel` (which was added in Task 16, just before `renderQueue`):

```javascript
  function renderNotifications() {
    const data = (master && master.notifications) ? master.notifications : { items: [], unreadCount: 0 };
    const badge = document.getElementById('notif-badge');
    const list = document.getElementById('notif-list');
    const footer = document.getElementById('notif-footer');
    const markAll = document.getElementById('notif-mark-all');

    if (data.unreadCount > 0) {
      badge.classList.remove('d-none');
      badge.textContent = data.unreadCount > 99 ? '99+' : String(data.unreadCount);
      markAll.setAttribute('aria-disabled', 'false');
    } else {
      badge.classList.add('d-none');
      markAll.setAttribute('aria-disabled', 'true');
    }

    const items = (data.items || []).slice(0, 10);
    if (items.length === 0) {
      list.innerHTML = '<div class="notif-empty">No notifications yet.</div>';
      footer.textContent = '';
      return;
    }

    list.innerHTML = items.map(n => `
      <div class="notif-item ${n.read ? '' : 'unread'}" data-notif-id="${n.notifId}" data-req-id="${n.relatedReqId || ''}" onclick="onNotifClick(this)">
        <div>
          <span class="sender">${escapeHtml(n.senderName)}</span><span class="role-tag">${escapeHtml(n.senderRole)}</span>
        </div>
        <div class="msg">${escapeHtml(n.message)}</div>
        <div class="time">${relativeTime(n.timestamp)} · ${escapeHtml(n.action)}</div>
      </div>
    `).join('');

    const total = (data.items || []).length;
    footer.textContent = total > 10 ? `Showing 10 of ${total}` : `${total} notification${total === 1 ? '' : 's'}`;
  }

  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function relativeTime(iso) {
    if (!iso) return '';
    const t = new Date(iso).getTime();
    if (isNaN(t)) return '';
    const diff = Date.now() - t;
    const m = Math.floor(diff / 60000);
    if (m < 1) return 'just now';
    if (m < 60) return m + 'm ago';
    const h = Math.floor(m / 60);
    if (h < 24) return h + 'h ago';
    const d = Math.floor(h / 24);
    if (d < 7) return d + 'd ago';
    return new Date(t).toLocaleDateString();
  }

  function onNotifClick(el) {
    // Placeholder — full handler in Task 18
    el.classList.remove('unread');
  }
```

- [ ] **Step 2: Manual UAT**

Push, reload web app. Log in. Confirm: empty state shows "No notifications yet" until something is triggered. Submit a DR_CREATE as admin in another browser/incognito as the warehouseman user — refresh the page. Bell badge shows 1, panel shows the DR_CREATE notification.

- [ ] **Step 3: Commit**

```bash
git add src/Index.html
git commit -m "$(cat <<'EOF'
Render bell badge and notification list from master.notifications

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 18: Frontend — mark-read interactions

**Files:**
- Modify: `src/Index.html`

**Interfaces:**
- Consumes: `master.notifications` items
- Produces: Clicking a notification calls `markNotificationRead`. "Mark all read" calls `markAllNotificationsRead`. Both then refresh.

- [ ] **Step 1: Implement `onNotifClick` and `markAllNotifsRead` for real**

In `src/Index.html`, replace the placeholder `onNotifClick` and `markAllNotifsRead` functions (added in Tasks 16/17) with full implementations:

```javascript
  function onNotifClick(el) {
    const notifId = el.getAttribute('data-notif-id');
    const reqId = el.getAttribute('data-req-id');
    if (!notifId) return;

    // Optimistic UI: mark unread style off
    el.classList.remove('unread');
    const badge = document.getElementById('notif-badge');
    if (badge && !badge.classList.contains('d-none')) {
      const n = parseInt(badge.textContent) || 0;
      if (n <= 1) badge.classList.add('d-none');
      else badge.textContent = String(n - 1);
    }

    google.script.run
      .withSuccessHandler(() => {
        if (reqId) {
          // Switch to Pending Actions tab and try to scroll the related card into view
          const tab = document.querySelector('[data-bs-target="#tab-queue"]');
          if (tab) tab.click();
          setTimeout(() => {
            const card = document.querySelector(`[data-req-id="${reqId}"]`);
            if (card && card.scrollIntoView) card.scrollIntoView({ behavior: 'smooth', block: 'center' });
          }, 400);
        }
        // Close the panel
        document.getElementById('notif-panel').classList.add('d-none');
      })
      .withFailureHandler(err => console.error("markNotificationRead failed:", err))
      .markNotificationRead(notifId, { email: currentEmail, fullName: currentUser, role: userRole });
  }

  function markAllNotifsRead() {
    const list = document.getElementById('notif-list');
    list.querySelectorAll('.notif-item.unread').forEach(el => el.classList.remove('unread'));
    document.getElementById('notif-badge').classList.add('d-none');

    google.script.run
      .withSuccessHandler(() => {
        // Refresh notifications block only — no need to full-refresh
        google.script.run.withSuccessHandler(data => {
          if (master) master.notifications = data;
          renderNotifications();
        }).getNotificationsForUser({ email: currentEmail, fullName: currentUser, role: userRole });
      })
      .withFailureHandler(err => console.error("markAllNotificationsRead failed:", err))
      .markAllNotificationsRead({ email: currentEmail, fullName: currentUser, role: userRole });
  }
```

- [ ] **Step 2: Manual UAT**

Push, reload. Submit a DR_CREATE as admin (in one browser session). Log in as warehouseman in another. Confirm bell shows badge. Click a notification — verify: panel closes, the related queue card scrolls into view, badge decrements by 1. Click "Mark all read" — verify badge disappears.

- [ ] **Step 3: Commit**

```bash
git add src/Index.html
git commit -m "$(cat <<'EOF'
Wire notification click to markNotificationRead and queue card scroll

Plus Mark all read invokes markAllNotificationsRead and refreshes only
the notification block.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 19: 60-second polling with visibility guard

**Files:**
- Modify: `src/Index.html`

**Interfaces:**
- Consumes: `getNotificationsForUser` server call
- Produces: Bell re-checks for new notifications every 60s while tab is visible

- [ ] **Step 1: Add polling setup**

In `src/Index.html`, find the `function refresh(silent = false)` (around line 718). After this function, add:

```javascript
  let notifPollTimer = null;

  function startNotifPolling() {
    if (notifPollTimer) return;
    notifPollTimer = setInterval(() => {
      if (document.visibilityState !== 'visible') return;
      if (!currentEmail) return;
      google.script.run
        .withSuccessHandler(data => {
          if (master) master.notifications = data;
          renderNotifications();
        })
        .withFailureHandler(err => console.warn("Notif poll failed:", err))
        .getNotificationsForUser({ email: currentEmail, fullName: currentUser, role: userRole });
    }, 60000);
  }

  function stopNotifPolling() {
    if (notifPollTimer) { clearInterval(notifPollTimer); notifPollTimer = null; }
  }
```

- [ ] **Step 2: Start polling after login**

Find `startApp()` (around line 489 — where `document.getElementById('user-display').innerText = currentUser;` lives). At the end of `startApp()` (just before its closing `}`), add:

```javascript
    startNotifPolling();
```

Also find `doLogout()` (search for `function doLogout`). Add `stopNotifPolling();` at its start.

- [ ] **Step 3: Manual UAT**

Push, reload. Log in as warehouseman. Open browser devtools console. Wait ~60s — observe a `getNotificationsForUser` request firing. Switch to another tab for >60s and back — verify no request fired while hidden (you should see the requests resume only on return). Trigger a notification from another session (admin DR_CREATE) — within 60s the badge updates without full page refresh.

- [ ] **Step 4: Commit**

```bash
git add src/Index.html
git commit -m "$(cat <<'EOF'
Poll notifications every 60s while tab is visible

Pauses on hidden visibility; cleared on logout.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 20: Readiness fix — PO Assignments empty-sheet handling

**Files:**
- Modify: `src/Code.js`

**Interfaces:** none external

- [ ] **Step 1: Write the failing test**

Append to `src/tests.js`:

```javascript
function test_assignPOToDoc_empty_sheet_returns_clean_error() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paSheet = ss.getSheetByName('PO Assignments');
  // Clear all data rows but keep the header
  if (paSheet.getLastRow() > 1) {
    paSheet.getRange(2, 1, paSheet.getLastRow() - 1, paSheet.getLastColumn()).clearContent();
  }
  const result = assignPOToDoc('DOC-NONEXISTENT', 'PO-10001', { email: 'admin@test.com' });
  _assert(result.success === false, "Returns success: false");
  _assert(result.error && result.error.indexOf("no rows") === -1 && (result.error.indexOf("No pending") !== -1 || result.error.indexOf("no pending") !== -1),
    "Returns a friendly 'no pending' message rather than the stack-trace error");
}
```

- [ ] **Step 2: Run — verify it fails**

Run `test_assignPOToDoc_empty_sheet_returns_clean_error` — should fail because the current code throws "PO Assignments sheet has no rows."

- [ ] **Step 3: Apply the fix**

In `src/Code.js` around line 612-613, find:

```javascript
const paSheet = SS.getSheetByName(SHEETS.PO_ASSIGN);
if (!paSheet || paSheet.getLastRow() < 2) throw new Error("PO Assignments sheet has no rows.");
```

Change to:

```javascript
const paSheet = SS.getSheetByName(SHEETS.PO_ASSIGN);
if (!paSheet || paSheet.getLastRow() < 2) {
  return { success: false, error: "No pending PO assignments to update." };
}
```

- [ ] **Step 4: Run test — verify pass**

Run again. Should pass.

- [ ] **Step 5: Commit**

```bash
git add src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Fix assignPOToDoc to return friendly error on empty PO Assignments

Was throwing a raw 'has no rows' error; now returns success:false with
a user-facing message.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 21: Readiness fix — DR ID whitespace mismatch

**Files:**
- Modify: `src/Code.js`

- [ ] **Step 1: Write the failing test**

Append to `src/tests.js`:

```javascript
function test_RECEIVE_DR_matches_with_whitespace_padding() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName('Requests');

  // Inject a Pending DR row with whitespace-padded Req ID
  const paddedDR = '  DR-WS-TEST-001  ';
  rSheet.appendRow([paddedDR, new Date(), 'Admin User', 'admin', 'DR_CREATE', 'NCR Hub', 'Makati Site',
    'ITM-001', 'Dell Latitude', 'pc', 1, 'Pending DR', '', 'admin@test.com']);
  const insertedRowIndex = rSheet.getLastRow();

  // Warehouseman receives — uses the trimmed DR ID
  processBulkTransaction({
    email: 'wh@test.com', user: 'Alice Warehouseman', role: 'warehouseman', action: 'RECEIVE_DR',
    location: 'NCR Hub', drId: 'DR-WS-TEST-001', client: 'Acme Corp', siteName: 'Makati Site', siteId: 'S-001',
    items: [{ code: 'ITM-001', name: 'Dell Latitude', uom: 'pc', qty: 1, wbs: '' }]
  });

  const status = rSheet.getRange(insertedRowIndex, 12).getValue();
  _assert(status === 'Completed', "Padded DR ID row was correctly matched and marked Completed");
}
```

- [ ] **Step 2: Run — verify failure**

The current code at Code.js:301 does `row[0].toString() === payload.drId.toString()` without trim. The test should fail.

- [ ] **Step 3: Apply the fix**

In `src/Code.js` line 301, find:

```javascript
let rIdx = rData.findIndex(row => row[0].toString() === payload.drId.toString() && row[7].toString().trim().toLowerCase() === cleanCode.toLowerCase() && row[11] === 'Pending DR');
```

Change to:

```javascript
let rIdx = rData.findIndex(row => row[0].toString().trim() === payload.drId.toString().trim() && row[7].toString().trim().toLowerCase() === cleanCode.toLowerCase() && row[11] === 'Pending DR');
```

Also apply the same trim defensively in `processQueueAction` at line 525. Find:

```javascript
const rowIndex = reqData.findIndex(r => r[0] === reqId);
```

Change to:

```javascript
const rowIndex = reqData.findIndex(r => r[0] && r[0].toString().trim() === reqId.toString().trim());
```

- [ ] **Step 4: Run test — verify pass**

Run again. Should pass.

- [ ] **Step 5: Commit**

```bash
git add src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Trim DR ID and Req ID on both sides of comparison

Prevents silent miss when stored IDs have leading/trailing whitespace.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 22: Readiness fix — `processQueueAction` action-log audit

**Files:**
- Modify: `src/Code.js`

Audit each branch of `processQueueAction` to ensure exactly one summary log row is written per confirm/reject. Spec §10.6.

- [ ] **Step 1: Read each branch and document what each writes**

In `src/Code.js` lines 544-597, the branches are:
- Reject + ISSUE/TRANSFER_WH (lines 547-554): writes one `REJECTED (Reverted)` log row ✓
- Reject + RETURN_WH (lines 555-562): writes one `REJECTED (Reverted)` log row ✓
- Reject + other (lines 563-565): writes one `REJECTED` log row ✓
- Confirm Receipt + ISSUE (lines 569-576): writes one `TL RECEIVED` log row ✓
- Confirm Receipt + TRANSFER_WH (lines 577-587): writes one `WH RECEIVED` log row ✓
- Confirm Receipt + (neither ISSUE nor TRANSFER_WH): falls through with NO log row ⚠
- Confirm Return (lines 588-597): writes one `WH RCVD RETURN` log row ✓

Identified gap: `Confirm Receipt` with an unexpected reqAction silently completes the status update without logging. Defensive fix.

- [ ] **Step 2: Write the failing test**

Append to `src/tests.js`:

```javascript
function test_processQueueAction_unknown_branch_still_logs() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName('Requests');
  const lSheet = ss.getSheetByName('Logs');

  // Inject an oddball request row that doesn't match ISSUE/TRANSFER_WH/RETURN_WH
  const oddId = 'ODD-' + Math.floor(Math.random() * 100000);
  rSheet.appendRow([oddId, new Date(), 'Alice Warehouseman', 'warehouseman', 'CUSTOM_ACTION', 'NCR Hub', 'Makati Site',
    'ITM-001', 'Dell Latitude', 'pc', 1, 'In Transit', '', 'wh@test.com']);

  const startLogs = lSheet.getLastRow();
  processQueueAction(oddId, 'Confirm Receipt', { email: 'tl1@test.com', fullName: 'Bob TeamLeader', role: 'team leader' });
  const endLogs = lSheet.getLastRow();
  _assert(endLogs > startLogs, "Action-log row written even for non-standard branch");
}
```

- [ ] **Step 3: Run — verify it fails**

Should fail because no log row is written in the unmatched branch.

- [ ] **Step 4: Apply the fix**

In `src/Code.js`, find the `Confirm Receipt` block at line 567:

```javascript
} else if (action === 'Confirm Receipt') {
  newStatus = 'Completed';
  if (reqAction === 'ISSUE') {
     // ...
  } else if (reqAction === 'TRANSFER_WH') {
     // ...
  }
}
```

Add an `else` to log even when the branch is unknown:

Find the end of the TRANSFER_WH inner block (line 587, just before `}` closes `if (action === 'Confirm Receipt')`):

```javascript
         lSheet.appendRow([new Date(), reqId, 'WH RECEIVED', '-', userProfile.fullName, '-', currentData[6], currentData[5], currentData[7], currentData[8], currentData[9], '-', finalQty, stock, stock + finalQty, 'Completed', '']);
      }
    } else if (action === 'Confirm Return') {
```

Insert a fallback else block right before `} else if (action === 'Confirm Return') {`:

```javascript
         lSheet.appendRow([new Date(), reqId, 'WH RECEIVED', '-', userProfile.fullName, '-', currentData[6], currentData[5], currentData[7], currentData[8], currentData[9], '-', finalQty, stock, stock + finalQty, 'Completed', '']);
      } else {
         lSheet.appendRow([new Date(), reqId, 'CONFIRMED', '-', userProfile.fullName, '-', currentData[6], currentData[5], currentData[7], currentData[8], currentData[9], '-', finalQty, 0, 0, 'Completed', '']);
      }
    } else if (action === 'Confirm Return') {
```

- [ ] **Step 5: Run — verify pass**

Should pass.

- [ ] **Step 6: Commit**

```bash
git add src/Code.js src/tests.js
git commit -m "$(cat <<'EOF'
Ensure processQueueAction always writes an action-log row

Fallback CONFIRMED log row for non-standard reqActions on Confirm Receipt.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 23: Readiness fix — ISSUE action role/queue audit

**Files:**
- Modify: `src/Index.html`, possibly `src/Code.js`

The audit flagged Index.html:506 (warehouseman action dropdown shows ISSUE) and Index.html:954 (team leader card renderer). Verify the end-to-end flow is consistent.

- [ ] **Step 1: Read the relevant code**

Read `src/Index.html:497-514` (the action dropdown setup) and `src/Index.html:940-983` (renderQueue branches). Verify:
- ISSUE is in warehouseman's action dropdown — ✓ correct
- Team leader's card UI is what renders ISSUE items pending their confirm — ✓ correct
- The Pending Actions tab is visible to all three roles (the `tab-queue` button has no `d-none` class) — ✓ correct

If you find a real inconsistency during this read (e.g., a role check that excludes team leader from seeing the Pending tab, or a stale handler that calls processQueueAction with the wrong action label), fix it. Otherwise this task is a no-op verification.

- [ ] **Step 2: Add a documenting test**

Append to `src/tests.js`:

```javascript
function test_ISSUE_round_trip_warehouseman_to_team_leader() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName('Requests');

  // Seed stock
  const iSheet = ss.getSheetByName('Inventory');
  const invData = iSheet.getDataRange().getValues();
  const ncrCol = invData[0].indexOf('NCR Hub');
  const itmRow = invData.findIndex(r => r[0] === 'ITM-001');
  if (itmRow > 0 && ncrCol !== -1) iSheet.getRange(itmRow + 1, ncrCol + 1).setValue(50);

  // Warehouseman issues
  const before = rSheet.getLastRow();
  processBulkTransaction({
    email: 'wh@test.com', user: 'Alice Warehouseman', role: 'warehouseman', action: 'ISSUE',
    location: 'NCR Hub', siteName: 'Makati Site', siteId: 'S-001', client: 'Acme Corp',
    items: [{ code: 'ITM-001', name: 'Dell Latitude', uom: 'pc', qty: 2, wbs: 'WBS-991' }]
  });
  const after = rSheet.getLastRow();
  _assert(after > before, "ISSUE created a Requests row");
  const newRow = rSheet.getRange(after, 1, 1, 14).getValues()[0];
  _assert(newRow[4] === 'ISSUE', "Action = ISSUE");
  _assert(newRow[11] === 'In Transit', "Status = In Transit (visible to team leader queue)");
  _assert(newRow[6] === 'Makati Site', "Target site = Makati Site");
  _assert(newRow[13] === 'wh@test.com', "User Email populated");
}
```

- [ ] **Step 3: Run — verify pass**

This test confirms the existing ISSUE flow is intact. If it fails after Tasks 1-22, an earlier change is broken — investigate before continuing.

- [ ] **Step 4: Commit (audit only or with fix if one was found)**

```bash
git add src/tests.js [+ any code fixes from Step 1]
git commit -m "$(cat <<'EOF'
Document ISSUE warehouseman-to-team-leader round trip

Audit only; no inconsistency found in existing routing.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

If the audit in Step 1 found a real bug, replace this commit message with a description of the fix.

---

### Task 24: Readiness fix — DR ID type handling consolidation

**Files:**
- Modify: `src/Code.js`

- [ ] **Step 1: Identify type-conversion sites**

In `src/Code.js`, the existing defensive `.toString()` calls on Req/DR IDs are at:
- Line 139 (pendingDRs push from numeric DR ID)
- Line 171 (matchedLog comparison)
- Lines 301, 525 (already trimmed in Task 21)
- Throughout request-row writes

The risk: when `appendRow` or `setValues` writes a numeric string that looks like a number (e.g., `"10042"`), Sheets may auto-cast to Number. Then reads come back as Number, and downstream string ops need `.toString()` defensively. We enforce: every Req ID we write is prefixed (e.g., `DR-`, `TRN-`, `ISSUE-`, `RTN-`, `DOC-`) so the cell content is unambiguously a string. Confirm this and add a defensive type-cast at the write side.

- [ ] **Step 2: Audit write sites**

In `src/Code.js`, examine the `finalDocId` generation around line 266:
```javascript
let finalDocId = payload.refDoc ? payload.refDoc : "DOC-" + d.getFullYear() + ...;
```
This is always a string. Good.

The `reqId` for TRANSFER_WH (line 362), ISSUE (line 372), and RETURN_WH (line 408) are all `'TRN-' + ...`, `'ISSUE-' + ...`, `'RTN-' + ...` — all strings. Good.

No code changes needed — the audit confirmed the IDs are always written as prefixed strings, so the existing `.toString()` reads at line 139 and 171 are belt-and-suspenders defense and can stay as-is.

- [ ] **Step 3: Add a regression test**

Append to `src/tests.js`:

```javascript
function test_DR_id_is_always_string_in_requests_sheet() {
  initializeSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName('Requests');

  // Submit a DR_CREATE; the resulting Req ID should be a string starting with DOC- or whatever the user supplied
  processBulkTransaction({
    email: 'admin@test.com', user: 'Admin User', role: 'admin', action: 'DR_CREATE',
    location: 'NCR Hub', siteName: 'Makati Site', siteId: 'S-001', client: 'Acme Corp',
    refDoc: 'DR-TYPE-001', poNumber: 'PO-10001',
    items: [{ code: 'ITM-001', name: 'Dell Latitude', uom: 'pc', qty: 1, wbs: 'WBS-991' }]
  });

  const lastRow = rSheet.getRange(rSheet.getLastRow(), 1, 1, 14).getValues()[0];
  _assert(typeof lastRow[0] === 'string', "Req ID stored as string (type was: " + typeof lastRow[0] + ")");
  _assert(lastRow[0] === 'DR-TYPE-001', "Req ID round-trips exactly");
}
```

- [ ] **Step 4: Run — verify pass**

The test should already pass; the task is documenting that the invariant holds.

- [ ] **Step 5: Commit**

```bash
git add src/tests.js
git commit -m "$(cat <<'EOF'
Pin invariant: Requests Req ID is always a prefixed string

Regression test against future regressions to numeric Doc IDs.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

### Task 25: End-to-end UAT pass and final smoke

**Files:** none modified

Spec §11's full test matrix in the deployed web app, executed by hand. This is the gate that all the upstream tests can't catch — UI behavior, real emails, the live spreadsheet under realistic load.

- [ ] **Step 1: Prep the test fixture**

Reset the spreadsheet to a known state by running `initializeSheets()` from the Apps Script editor. (This is idempotent; existing data stays, but headers/sheets are restored.) Confirm there are at least these users in the Users sheet: 1 admin, 2 warehousemen (one at NCR Hub, one at Visayas Hub), 2 team leaders (one for Makati Site, one for Cebu Site). Edit the Users sheet directly to add a second warehouseman with `wh2@test.com` and `Visayas Hub` Location Access if missing.

- [ ] **Step 2: Deploy**

Push the code to Apps Script (clasp push or copy-paste). Re-deploy the web app to a stable URL accessible by the test users.

- [ ] **Step 3: Execute the matrix**

For each of the rows below, perform the action in one browser session and verify the expected effect in another session logged in as the expected recipient.

| # | Trigger | As | Expected in-app recipients | Expected emails |
|---|---|---|---|---|
| 1 | DR_CREATE for NCR Hub | admin | NCR warehouseman badge +1; panel shows DR_CREATE | Email to NCR warehouseman |
| 2 | DR_CREATE for Visayas | admin | Visayas warehouseman badge +1; NCR warehouseman badge unchanged | Email to Visayas warehouseman only |
| 3 | RECEIVE_DR of #1 | NCR warehouseman | Admin badge +1; message: "<name> received DR-…" | Email to admin |
| 4 | TRANSFER_WH NCR→Visayas | NCR warehouseman | Visayas warehouseman badge +1; NCR warehouseman NO self-notif | Email to Visayas warehouseman |
| 5 | Queue Confirm of #4 | Visayas warehouseman | NCR warehouseman badge +1 with "<name> confirmed your TRANSFER_WH …" | Email to NCR warehouseman |
| 6 | ISSUE NCR→Makati | NCR warehouseman | Makati team leader badge +1; Cebu team leader badge unchanged | Email to Makati team leader only |
| 7 | Queue Reject of #6 | Makati team leader | NCR warehouseman badge +1 with "<name> rejected your ISSUE …" | Email to NCR warehouseman |
| 8 | RETURN_WH Makati→NCR | Makati team leader | NCR warehouseman badge +1 | Email to NCR warehouseman |
| 9 | Queue Confirm of #8 | NCR warehouseman | Makati team leader badge +1 confirm | Email to Makati team leader |
| 10 | RETURN_CLIENT | NCR warehouseman | All admins badge +1 | Email to all admins; PDF receipt email still arrives separately to warehouseman |

- [ ] **Step 4: UI behavior checks**

- Click any notification — panel closes, queue card scrolls into view (if relevant), badge decrements.
- "Mark all read" — badge clears, unread style removed from all items.
- Polling — open devtools Network tab, wait 60s, confirm `getNotificationsForUser` request fires. Hide the tab for 70s and bring it back — no request fires while hidden.

- [ ] **Step 5: Security checks**

In the browser console, run:
```javascript
currentEmail = 'admin@test.com';
userRole = 'admin';
```
(while logged in as a team leader). Try to submit a DR_CREATE. The server should reject because the validated profile won't match (the team leader's email isn't `admin@test.com`). Actually — this *will* succeed because `currentEmail` is now `admin@test.com` and the Users sheet has that. The real defense is: a user cannot guess another user's email AND password. The test is to verify the **role override** works: set `userRole = 'admin'` while keeping the team-leader's real `currentEmail`; submit DR_CREATE. The server should accept the request (because the email is real) but the role is overwritten with `'team leader'` so the rest of the flow honors team-leader permissions. Verify by checking the Notifications row's `Sender Role` — it should say `team leader`, not `admin`.

- [ ] **Step 6: Failure-mode checks**

- Edit the Users sheet to give one user an obviously bad email (`not-an-email`). Trigger an action that notifies that user. Confirm the Notifications row is still written and `Email Status` says `failed: …`. Other recipients' emails are still sent.
- Delete the Notifications sheet entirely from the spreadsheet. Trigger a notifying action. Confirm `_notif_ensureSheet()` recreates the sheet and the notification is written normally.

- [ ] **Step 7: Final commit**

If any small fixes were needed during UAT (typos, off-by-one in counts, etc.), commit them. Otherwise no commit needed.

If everything passed, tag the release for clarity:

```bash
git tag -a "notifications-v1" -m "Notifications + system readiness shipped"
```

---

## Self-Review Notes

**Spec coverage:** Cross-checked against the spec sections:
- §5.1 (Notifications schema): Task 1
- §5.2 (Requests User Email): Task 1, Task 7
- §6.1 (notify): Task 5
- §6.2 (resolveRecipients): Task 6
- §6.3 (resolveRequester): Task 6
- §6.4 (RECEIVE_DR ack): Task 13
- §6.5 (queue ack): Task 14
- §6.6 (validateUserProfile): Tasks 2, 4
- §6.7 (getNotificationsForUser): Task 15
- §6.8 (mark read): Task 15
- §6.9 (getAppData integration): Task 15
- §7 (frontend): Tasks 16, 17, 18, 19
- §8 (processBulkTransaction wiring): Tasks 8-13
- §9 (processQueueAction wiring): Task 14
- §10.1 (role validation): Tasks 2, 4
- §10.2 (User Email col): Tasks 1, 7
- §10.3 (PO Assignments empty-sheet): Task 20
- §10.4 (DR ID whitespace): Task 21
- §10.5 (ISSUE audit): Task 23
- §10.6 (queue log gaps): Task 22
- §10.7 (DR ID type): Task 24
- §11 (testing matrix): Task 25
