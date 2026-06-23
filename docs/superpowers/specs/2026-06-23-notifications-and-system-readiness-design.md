# Notifications + System Readiness for New-Central

**Date:** 2026-06-23
**Scope:** `src/Code.js`, `src/Index.html`, `src/setup.js`
**Constraint:** Preserve all existing actions, submit/finalize logic, queue behavior, and login flow.

---

## 1. Background

New-Central is a Google Apps Script web app for inventory management with three roles (admin, warehouseman, team leader) backed by a custom `Users` sheet (Email, Full Name, Password, Salt, Role, Location Access, Site Access). Login is handled by `user.js#loginUser` — the app does NOT use `Session.getActiveUser()` (the Google account). The frontend keeps a `userProfile` object after login and passes it on every backend call.

Eight actions cycle work between roles:

- `DR_CREATE` (admin) → leaves `Pending DR` for warehouseman to `RECEIVE_DR`
- `RECEIVE_DR` (warehouseman) — completes immediately
- `PURCHASE_LOG` (admin) — completes immediately
- `TRANSFER_WH` (warehouseman, source) → leaves `Pending Receipt` for warehouseman (target)
- `ISSUE` (warehouseman) → leaves `In Transit` for team leader
- `USAGE` (team leader) — completes immediately
- `RETURN_CLIENT` (warehouseman) — completes immediately, emails return PDF
- `RETURN_WH` (team leader) → leaves `Pending Return` for warehouseman

Today the only signal a recipient gets that work is waiting for them is the "Pending Actions" queue inside the app — they have to open the app and look. There is no email notification, no in-app bell, and no acknowledgment signal back to the original requester when the recipient acts. The user wants both:

1. **Notifications** — in-app bell + email — fire whenever a new action lands on someone's plate, and again as an acknowledgment when they confirm or reject it.
2. **System readiness** — fix bugs that affect notification correctness or are obvious crash risks, since we're already in these files.

## 2. Goals

1. When an action creates work for another role, every targeted user receives (a) an in-app notification in the bell dropdown and (b) an email.
2. When a recipient confirms or rejects the work in the queue, the original requester receives a confirmation/rejection notification in-app and by email.
3. Recipient targeting respects `Location Access` and `Site Access` columns in the Users sheet — a warehouseman at "NCR Hub" never gets notified for activity at "Visayas Hub".
4. All recipient/sender identities resolve via the **app's `Users` sheet** (looked up by email), never via `Session.getActiveUser()`.
5. Critical bugs that would let notifications fire to the wrong people or crash the request flow are fixed in the same release.

## 3. Non-Goals

- No per-user "email notifications on/off" preference (deferred — single global on/off behavior).
- No notification batching/digest (one email per recipient per event).
- No SMS, Slack, or push notifications.
- No retroactive notifications for historical Requests rows — only new activity from the release point forward.
- No refactor of `handleActionChange()`, the Requests schema beyond one added column (`User Email`), the Logs schema, or login flow.
- The action-routing inconsistency for `ISSUE` and the action-log gaps in `processQueueAction` are addressed; deeper queue refactor is out of scope.

## 4. Architecture summary

```
                              ┌────────────────────┐
   ┌──────────────────┐       │   Notifications    │
   │  Code.js         │──────▶│   sheet (new)      │
   │  notify(...)     │       │                    │
   │                  │       │  one row per       │
   │  - writes rows   │       │  (recipient×event) │
   │  - MailApp.send  │       └────────────────────┘
   └──────────────────┘                ▲
            ▲                          │
            │ called from              │ read by
            │                          │
   ┌──────────────────────────┐   ┌───────────────────────┐
   │ processBulkTransaction() │   │ getNotificationsFor   │
   │ processQueueAction()     │   │ User(userProfile)     │
   └──────────────────────────┘   └───────────────────────┘
                                          ▲
                                          │ included in getAppData() response
                                          │
                                  ┌───────────────────────┐
                                  │  Index.html bell UI   │
                                  │  - badge w/ unread #  │
                                  │  - dropdown panel     │
                                  │  - 60s poll           │
                                  └───────────────────────┘
```

Sender identity comes from `userProfile.email` (re-validated server-side against the Users sheet). Recipients are resolved by querying the Users sheet for users whose `Role` matches the target role AND whose `Location Access` / `Site Access` includes the action's location/site (where applicable).

## 5. Schema changes

### 5.1 New sheet: `Notifications`

Columns (added to `setup.js#initializeSheets`):

| # | Column | Type | Notes |
|---|---|---|---|
| 0 | Notif ID | string | `NOTIF-yyyyMMdd-HHmmss-<rand4>` |
| 1 | Timestamp | date | server-side `new Date()` |
| 2 | Recipient Email | string | from Users sheet col 0 |
| 3 | Recipient Name | string | from Users sheet col 1 |
| 4 | Recipient Role | string | from Users sheet col 4 |
| 5 | Sender Email | string | from `userProfile.email` (re-validated) |
| 6 | Sender Name | string | from `userProfile.fullName` |
| 7 | Sender Role | string | from validated server-side role |
| 8 | Action | string | one of the 8 action codes, or `CONFIRM` / `REJECT` |
| 9 | Related Req ID | string | the Requests-sheet Req ID, or DR ID for RECEIVE_DR ack |
| 10 | Message | string | human-readable, e.g. `Admin Shane created DR-1023 (3 items) — pending your receipt` |
| 11 | Read | bool | `false` on insert; `true` after `markNotificationRead` |
| 12 | Read At | date \| empty | timestamp when marked read |
| 13 | Email Status | string | `sent` \| `failed: <err>` \| `skipped: no email` |

Frozen header row, bold dark background, same styling as the other sheets.

### 5.2 New column on `Requests`: `User Email`

Append a 14th column `User Email` to the Requests schema in `setup.js`. New rows written by `processBulkTransaction()` set this column to `userProfile.email`. This lets `processQueueAction()` resolve the original requester's email directly for acknowledgment notifications without name-based lookup. Existing Requests rows have an empty value; acknowledgment for those falls back to name lookup in Users sheet (best-effort).

## 6. Backend additions (Code.js)

All new code goes in `Code.js`. No backend file split.

### 6.1 `notify(recipients, action, sender, message, relatedReqId)`

```
recipients: array of {email, name, role} — required
action:     string (action code or 'CONFIRM' or 'REJECT')
sender:     {email, name, role} — required, must be server-validated
message:    string — required, human-readable
relatedReqId: string — optional, links back to a Requests row or DR ID
```

Behavior:
1. For each recipient, generate a `Notif ID`.
2. Build one row per recipient and batch-append to `Notifications` sheet in a single `setValues` call.
3. For each recipient with a non-empty email, call `MailApp.sendEmail(email, subject, body)` (one email per recipient). Wrap each in `try/catch` — a single email failure must not abort the whole notify call. Set `Email Status` column per recipient to `sent` / `failed: <err>` / `skipped: no email`. The `Email Status` value is written in the same batch — we send the emails first, collect statuses, then write the final batch.
4. Returns `{ inserted: N, emailsSent: M, emailsFailed: K }`. Caller logs but does not propagate failure — notification failure must never block the underlying business transaction.

Subject: `[Inventory] <Action> by <SenderName> — <short detail>`
Body:
```
Hi <RecipientName>,

<SenderName> (<SenderRole>) just performed <Action>.

<Message>
Related ID: <RelatedReqId>

Open the inventory app to confirm or reject:
<web-app-url from ScriptApp.getService().getUrl()>

— Inventory System
```

### 6.2 `resolveRecipients(action, payload)`

Returns the list of `{email, name, role}` who should be notified. Reads Users sheet once per call (within `processBulkTransaction` lock).

Mapping:

| Action | Target role | Scope filter |
|---|---|---|
| `DR_CREATE` | warehouseman | `Location Access` empty OR includes `payload.location` |
| `RECEIVE_DR` | (no fan-out — only ack to creator, see §6.4) | — |
| `PURCHASE_LOG` | (no notification) | — |
| `TRANSFER_WH` | warehouseman | `Location Access` empty OR includes `payload.targetLoc` |
| `ISSUE` | team leader | `Site Access` empty OR includes `payload.siteName` |
| `USAGE` | (no notification) | — |
| `RETURN_CLIENT` | admin | (no scope filter — admins are global) |
| `RETURN_WH` | warehouseman | `Location Access` empty OR includes `payload.location` (the warehouse the return is going back to) |

Empty access string means "all" — preserves the same logic the existing queue filtering uses in `getAppData()` (Code.js:121-122, 132, 137, 145).

### 6.3 `resolveRequester(reqId)`

Given a `reqId`, returns `{email, name, role}` for the user who originally created that request:
1. Read Requests sheet, find row with matching Req ID.
2. If `User Email` column (new col 13) is populated, look up that email in the Users sheet — return the row.
3. Else fall back to matching `User Name` (col 2) against Users sheet col 1.
4. If neither resolves, return `null`. Caller treats this as "skip acknowledgment notification" and logs a warning.

### 6.4 Acknowledgment after `RECEIVE_DR`

`RECEIVE_DR` completes immediately, not via the queue. After the success path in `processBulkTransaction()` (around Code.js:319), call `resolveRequester(payload.drId)` (the DR's own Req ID) and `notify([creator], 'CONFIRM', sender, "Warehouseman <name> received DR-<id> (<N> items)", payload.drId)`.

### 6.5 Acknowledgment after queue confirm/reject

In `processQueueAction()` after the status update at ~Code.js:599, call `resolveRequester(reqId)` and `notify([creator], action === 'Reject' ? 'REJECT' : 'CONFIRM', sender, message, reqId)` where `sender` is the validated profile of the user calling `processQueueAction` and `message` describes the outcome:
- Confirm: `"<RecipientName> (<RecipientRole>) confirmed your <Action> request <reqId>"`
- Reject: `"<RecipientName> (<RecipientRole>) rejected your <Action> request <reqId>"`

### 6.6 `validateUserProfile(userProfile)` — server-side trust gate

Used at the top of `processBulkTransaction`, `processQueueAction`, `markNotificationRead`, `markAllNotificationsRead`, and `getNotificationsForUser`:

1. If `userProfile` is null/undefined or missing `email` → throw `"Authentication required."`
2. Read Users sheet, find row with `Email` (col 0) equal to `userProfile.email` (case-insensitive trim).
3. If no row found → throw `"User account not recognized."`
4. Build a **validated** profile from the sheet row (NOT from the client payload): `{email, fullName, role, locAccess, siteAccess}`. Return this. The caller uses ONLY the returned object — never the client-supplied one — for role/scope decisions and for `sender` in `notify`.

This stops a tampered frontend from claiming a different role or identity. It does NOT verify the user is currently authenticated (Apps Script doesn't have a session-token model for HTML service); it verifies the claimed identity exists in the Users sheet. Combined with the password-gated login flow, this is the strongest guarantee available in the existing architecture.

### 6.7 `getNotificationsForUser(userProfile)`

1. `validateUserProfile(userProfile)` → validated.
2. Read Notifications sheet (skip if doesn't exist yet — return `{ items: [], unreadCount: 0 }`).
3. Filter rows where `Recipient Email` equals validated email (case-insensitive).
4. Sort descending by Timestamp; take the last 30.
5. Return:
   ```
   {
     items: [{notifId, timestamp, sender, action, message, relatedReqId, read}, ...],
     unreadCount: <number of unread across all matches, not just the 30>
   }
   ```

### 6.8 `markNotificationRead(notifId, userProfile)` / `markAllNotificationsRead(userProfile)`

- `markNotificationRead`: validate → find row by Notif ID → verify Recipient Email matches validated email (else throw "Not your notification") → set Read=true, Read At=now.
- `markAllNotificationsRead`: validate → find all unread rows for validated email → batch update Read=true, Read At=now in a single `setValues` call.

### 6.9 Integration with `getAppData()`

Append to the return object:
```js
return {
  inventory, logs, dropdowns, pending, pendingDRs, receiptDocs,
  notifications: getNotificationsForUser(userProfile)  // {items, unreadCount}
};
```

If `userProfile` is missing or invalid (`validateUserProfile` throws), catch and substitute `{ items: [], unreadCount: 0 }` so the rest of `getAppData` still works for backward compatibility.

## 7. Frontend additions (Index.html)

### 7.1 Bell icon

In the existing header row alongside the user name greeting:

```
[Logo] Inventory System          🔔 [3]   Welcome, Shane (admin)   [Logout]
```

A bell SVG inside a clickable `<button>`, with an absolutely-positioned red circular badge showing unread count. Badge hidden when count is 0. Match existing header styling.

### 7.2 Dropdown panel

Click the bell → fixed-position dropdown anchored bottom-right of the bell:

- 320px wide
- Header: `Notifications` + `Mark all read` link (only enabled if unreadCount > 0)
- List of up to 10 most recent notifications, newest first
- Each item: sender name (bold) + role tag, action, message (truncated to 2 lines), relative time (`2m ago`, `1h ago`, `Yesterday`, etc.)
- Unread items have a left blue stripe and light-blue background
- Click an item → mark read, close dropdown, and (if `relatedReqId` is present and the user has a matching queue item) scroll the queue card into view and pulse-highlight it. Otherwise, just mark read.
- Footer: count of total notifications fetched, if more exist beyond 10
- Click-outside or Esc closes the dropdown

### 7.3 Refresh strategy

- On startup: `master.notifications` arrives in the initial `getAppData()` payload. Render badge and pre-populate dropdown.
- On poll: every 60 seconds while the tab is visible, call `google.script.run.getNotificationsForUser(userProfile)` and re-render badge + dropdown items. Skip polling when tab is hidden (use `document.visibilityState`).
- On submit success: `submitItems()` already refreshes app data after a successful transaction. Reuse that path — notifications come back fresh in the next `getAppData()` payload.

### 7.4 Marking as read

- Single read: `google.script.run.withSuccessHandler(...).markNotificationRead(notifId, userProfile)` then optimistically update local state.
- Mark all: `markAllNotificationsRead(userProfile)` then set all local items to read.

## 8. Wiring inside `processBulkTransaction()`

After the main body completes successfully but before `lock.releaseLock()` (so that notification writes are also lock-protected against concurrent reads), call notification helpers. If notify throws or returns errors, log and continue — never fail the underlying transaction.

Insertion points (line numbers approximate to current state, will shift slightly during implementation):

| After Code.js line | Action | Call |
|---|---|---|
| ~319 (`RECEIVE_DR` branch ends) | RECEIVE_DR | `notify(resolveRequester(payload.drId) → [creator], 'CONFIRM', sender, msg, payload.drId)` |
| ~343 (DR_CREATE per-item loop) — actually after the items loop completes, once per DR | DR_CREATE | `notify(resolveRecipients('DR_CREATE', payload), 'DR_CREATE', sender, msg, finalDocId)` |
| ~365 (TRANSFER_WH per-item) — once per transfer reqId | TRANSFER_WH | `notify(resolveRecipients('TRANSFER_WH', payload), 'TRANSFER_WH', sender, msg, reqId)` |
| ~374 (ISSUE per-item) — once per ISSUE reqId | ISSUE | `notify(resolveRecipients('ISSUE', payload), 'ISSUE', sender, msg, reqId)` |
| ~410 (RETURN_WH per-item) — once per RETURN_WH reqId | RETURN_WH | `notify(resolveRecipients('RETURN_WH', payload), 'RETURN_WH', sender, msg, reqId)` |
| After PDF generation in RETURN_CLIENT block (~Code.js:480) | RETURN_CLIENT | `notify(resolveRecipients('RETURN_CLIENT', payload), 'RETURN_CLIENT', sender, msg, finalDocId)` |

**Deduplication**: `DR_CREATE`, `TRANSFER_WH`, `ISSUE`, `RETURN_WH` all loop over items and create one Requests row per item, but the notification should be one per transaction (per Doc/Req grouping), not per item. Solution: collect notification triggers in arrays during the item loop, then flush once after the loop with a summary message like `"Admin Shane created DR-1023 (3 items: ITM-001, CBL-100, ITM-002) — pending your receipt"`.

## 9. Wiring inside `processQueueAction()`

After `rSheet.getRange(rowNum, 12).setValue(newStatus)` at ~Code.js:599 and before the return:

```js
let creator = resolveRequester(reqId);
if (creator) {
  let outcome = (action === 'Reject') ? 'REJECT' : 'CONFIRM';
  let msg = (action === 'Reject')
    ? `${validated.fullName} (${validated.role}) rejected your ${reqAction} request ${reqId}`
    : `${validated.fullName} (${validated.role}) confirmed your ${reqAction} request ${reqId}`;
  notify([creator], outcome, validated, msg, reqId);
}
```

Wrapped in try/catch; never blocks the queue action's success return.

## 10. Readiness fixes bundled in

Each is independently verifiable; all are in scope for this same release. Items are listed in priority order.

### 10.1 [Critical] Server-side identity & role validation
- **Problem:** `processBulkTransaction`, `processQueueAction`, etc., trust `userProfile.role` from the frontend. A tampered client could claim `admin` and execute admin-only actions. With notifications added, a tampered client could also impersonate the sender in the notification record.
- **Fix:** Add `validateUserProfile()` (§6.6) and call it at the top of every server entry point. Use only the validated profile for downstream logic and for `sender` in `notify`.
- **Files:** `Code.js` (function added; called from `processBulkTransaction`, `processQueueAction`, `getAppData`, `getNotificationsForUser`, `markNotificationRead`, `markAllNotificationsRead`, `assignPOToDoc`).

### 10.2 [Critical] Add `User Email` column to Requests sheet
- **Problem:** `processQueueAction` can't reliably look up the original requester by name (names can have whitespace, can change, can collide). The audit also flagged `Code.js:477` (`r[1].toString().trim() === payload.user.toString().trim()`) as fragile.
- **Fix:** Add `User Email` as the 14th Requests column in `setup.js`. Write `userProfile.email` in this column in every `processBulkTransaction` Requests insert. `resolveRequester` reads it first, falls back to name match for legacy rows. The RETURN_CLIENT email lookup at Code.js:477 also switches to `payload.userEmail` (passed from the client) → look up in Users sheet → email → fallback to name only if email column is empty.
- **Files:** `Code.js`, `setup.js`, `Index.html` (submit payload now includes `userEmail: userProfile.email`).

### 10.3 [Important] PO Assignments empty-sheet crash
- **Problem:** `assignPOToDoc` at Code.js:612-613 throws when the PO Assignments sheet exists but is empty (`getLastRow() < 2`).
- **Fix:** Change to a clean `return { error: "No pending PO assignments found." }` or equivalent guard that doesn't throw — let the UI handle the empty state.
- **Files:** `Code.js`.

### 10.4 [Important] DR ID whitespace mismatch on read
- **Problem:** Code.js:301 does `row[0].toString() === payload.drId.toString()` without trim. If a Requests row's Req ID has surrounding whitespace, the RECEIVE_DR match fails silently — status never updates to Completed.
- **Fix:** Trim both sides in the comparison. Apply the same trim to all `row[0]` comparisons throughout `Code.js` (e.g., Code.js:525 `processQueueAction`).
- **Files:** `Code.js`.

### 10.5 [Important] ISSUE action role/queue inconsistency
- **Problem:** ISSUE is created by warehouseman and consumed by team leader, but the audit flagged some UI/queue role-mapping inconsistencies between Index.html:506 (warehouseman action dropdown) and Index.html:954 (team leader card renderer).
- **Fix:** Audit the ISSUE flow end-to-end. Confirm warehouseman-creates → 'In Transit' Requests row → team leader sees in their queue → confirm/reject. Fix any place the role check is wrong or stale. Add an explicit `processQueueAction` action-log row for the team-leader receipt (currently logged but worth verifying matches the pattern for TRANSFER_WH).
- **Files:** `Code.js`, `Index.html`.

### 10.6 [Important] `processQueueAction` action-log gap on RETURN_WH confirm
- **Problem:** Code.js:597 writes "WH RCVD RETURN" log row for `Confirm Return`. Confirm matrix verified, but the audit noted that the audit-trail row was missing or inconsistent in some paths. Verify each branch writes exactly one summary log row including the rejected/reverted cases (Code.js:554, 562, 564, 576, 586, 596).
- **Fix:** Audit each branch. Add any missing log rows. Confirm format consistency.
- **Files:** `Code.js`.

### 10.7 [Cleanup] DR ID type handling consolidation
- **Problem:** Multiple `.toString()` conversions on `row[0]` in different code paths suggest past type confusion. The Requests sheet should consistently store Req IDs as strings.
- **Fix:** When writing Requests rows, ensure `finalDocId` and other Req IDs are always strings before `appendRow`/`setValues`. Where stale rows might have number-typed IDs, the existing `.toString()` defensive code stays as a fallback.
- **Files:** `Code.js`.

## 11. Testing

### 11.1 Test data setup
- Reset spreadsheet via `initializeSheets()` (which seeds default Users with admin/2 warehousemen/2 team leaders across 2 locations and 4 sites).
- Add a second warehouseman with `Location Access = "NCR Hub"` to test fan-out within a role/scope.

### 11.2 Notification matrix (one row per notifying action)

| Trigger | Performed by | Expected in-app recipients | Expected emails | Verify |
|---|---|---|---|---|
| `DR_CREATE` for NCR Hub | admin | All warehousemen with `Location Access` covering NCR Hub | Same | Badge +1 for each; one email each; Visayas-only warehousemen get nothing |
| `RECEIVE_DR` | warehouseman | Admin who created the DR | Same | Admin badge +1, message says "Warehouseman X received DR-… (N items)" |
| `TRANSFER_WH` NCR→Visayas | warehouseman (NCR) | Warehousemen covering Visayas Hub | Same | NCR warehouseman gets NO self-notification |
| `ISSUE` from NCR Hub to Makati Site | warehouseman | Team leaders covering Makati Site | Same | Cebu team leaders get nothing |
| `RETURN_CLIENT` | warehouseman | All admins | Same | Existing PDF email still arrives separately to the submitter |
| `RETURN_WH` Makati→NCR | team leader | Warehousemen covering NCR Hub | Same | — |
| Queue Confirm (TRANSFER_WH, ISSUE, RETURN_WH) | queue recipient | Original requester | Same | Message says "X confirmed your <action> request <reqId>". (RECEIVE_DR ack is covered by the `RECEIVE_DR` row above — it doesn't go through `processQueueAction`.) |
| Queue Reject (TRANSFER_WH, ISSUE, RETURN_WH) | queue recipient | Original requester | Same | Message says "X rejected your <action> request <reqId>" |

### 11.3 In-app UI checks
- Bell badge increments after submit (within next poll cycle, or immediately on next `getAppData()`).
- Clicking dropdown shows newest-first list.
- Clicking an item marks it read, closes the dropdown, and scrolls to the queue card (if relatedReqId matches a card).
- "Mark all read" zeros the badge and grays previously-unread items.
- Polling pauses when tab is hidden.

### 11.4 Identity & security checks
- Edit the client-side `userProfile` in the browser console to claim `role: 'admin'` while logged in as a team leader; submit a DR_CREATE → server must reject with `"User account not recognized."` or the role-validated profile must overwrite the claim so the action fails its own role check.
- Edit `userProfile.email` to a fake address → server must reject with `"User account not recognized."`.

### 11.5 Readiness regression checks
- Whitespace-padded DR ID in Requests sheet → RECEIVE_DR still matches.
- Empty PO Assignments sheet → `assignPOToDoc` returns a clean error, not a stack trace.
- A Requests row with no `User Email` value (legacy row) → `resolveRequester` falls back to name lookup and finds the user.

### 11.6 Failure-mode checks
- One recipient's email is malformed → other recipients still get their emails; the bad row's `Email Status` column shows `failed: …`. Notification ROW still appears in-app for that user (in-app independent of email).
- Notifications sheet doesn't exist yet (first run after deploy, before re-running `initializeSheets`) → `notify` creates it on the fly OR `getAppData` returns empty notifications and a clear warning in the server log; choose whichever is least disruptive. Recommended: have `notify` insert the sheet if missing (matches the pattern used elsewhere for `Discrepancies` and `PO Assignments`).

## 12. Risks

- **MailApp quota:** 100/day on consumer Gmail accounts, 1,500/day on Workspace. Worst-case fan-out (e.g., DR_CREATE notifying 5 warehousemen) at ~30 actions/day = 150 emails/day, comfortably under Workspace limits. If on consumer Gmail this could be tight; note in deployment instructions.
- **Email noise:** Several users in the same role/scope all get every notification. Mitigation deferred (per-user toggle is non-goal). Recommend training users to filter emails by subject prefix `[Inventory]`.
- **Notifications-sheet growth:** No automatic pruning. Over time the sheet grows unboundedly. Reasonable at this app's scale, but add a manual archival recipe to the operations notes; do not implement auto-prune now.
- **Polling cost:** 60s `getNotificationsForUser` poll × many users × all day. Apps Script quotas per user are generous, but the spreadsheet read on every poll is a `getDataRange().getValues()` plus filter. If the Notifications sheet grows large this becomes slow. Mitigation: in §6.7 we already only return the last 30 items; if needed later, swap to a tail-read using `getRange(lastRow-200, …)` rather than full-sheet read.
- **Race between transaction and notification:** Both run inside the same `LockService` lock, so a recipient calling `getNotificationsForUser` won't miss a freshly-written notification.
- **Sender identity drift:** Validated sender comes from the Users sheet at the moment of the action. If the user's email or full name changes later, the notification record retains the historical value. Acceptable.

## 13. Files touched

| File | Changes |
|---|---|
| `src/Code.js` | Add `notify`, `resolveRecipients`, `resolveRequester`, `validateUserProfile`, `getNotificationsForUser`, `markNotificationRead`, `markAllNotificationsRead`. Wire calls into `processBulkTransaction` and `processQueueAction`. Apply readiness fixes 10.1, 10.2, 10.3, 10.4, 10.5, 10.6, 10.7. Append `notifications` to `getAppData` return. |
| `src/Index.html` | Add bell button + dropdown panel HTML in the header. Add CSS for badge/panel/items. Add JS: render `master.notifications`, click handlers, 60s `setInterval` poll guarded by `document.visibilityState`, `markNotificationRead` / `markAllNotificationsRead` calls. Pass `userEmail: userProfile.email` in every submit payload. |
| `src/setup.js` | Add `Notifications` sheet definition with the 14 columns from §5.1. Add `User Email` column to the Requests schema. Re-run `initializeSheets()` is idempotent — existing sheets only add missing columns. |

## 14. Implementation order recommendation

Suggested for the implementation plan that follows this spec:

1. Schema & infra: add `Notifications` sheet + `User Email` column to Requests in `setup.js`. Verify `initializeSheets()` is still idempotent.
2. `validateUserProfile()` + wire into all entry points (readiness fix 10.1).
3. `notify()`, `resolveRecipients()`, `resolveRequester()`.
4. Wire `notify` into `processBulkTransaction` (one action at a time, test each).
5. Wire `notify` into `processQueueAction`.
6. `getNotificationsForUser` + integrate into `getAppData`.
7. Bell UI in `Index.html` (badge, dropdown, click-to-mark-read).
8. 60s polling with visibility guard.
9. Readiness fixes 10.3, 10.4, 10.5, 10.6, 10.7.
10. End-to-end test matrix from §11.
