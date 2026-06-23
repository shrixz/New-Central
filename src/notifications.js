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
  let webAppUrl = '';
  try { webAppUrl = ScriptApp.getService().getUrl() || ''; } catch (_) {}
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
    throw new Error("Not your notification");
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
