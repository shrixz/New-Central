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
