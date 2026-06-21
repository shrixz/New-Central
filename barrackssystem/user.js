/**
 * Authentication & Security Logic
 * First-time detection: column D (Salt) is empty -> user is new -> modal opens
 *   and the temp password in col C is compared as plaintext.
 * finalizePassword: generates a fresh salt, writes SHA-256(password + salt) to col C
 *   and the salt to col D. Subsequent logins compare the hash.
 */
function loginUser(identifier, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = ss.getSheetByName('Users').getDataRange().getValues();
    const input = identifier.toString().trim().toLowerCase();

    const rowIndex = data.findIndex(row =>
      (row[0] && row[0].toString().trim().toLowerCase() === input) ||
      (row[1] && row[1].toString().trim().toLowerCase() === input)
    );

    if (rowIndex === -1) return { error: "User account not found." };

    const email = data[rowIndex][0] ? data[rowIndex][0].toString().trim() : "";
    const fullName = data[rowIndex][1];
    const storedPass = data[rowIndex][2];
    const salt = data[rowIndex][3];

    const role = data[rowIndex][4] ? data[rowIndex][4].toString().toLowerCase() : "team leader";
    const locAccess = data[rowIndex][5] ? data[rowIndex][5].toString().trim() : "";
    const siteAccess = data[rowIndex][6] ? data[rowIndex][6].toString().trim() : "";

    const isNewUser = (!salt || salt.toString().trim() === "");

    if (isNewUser) {
      if (password.toString() === storedPass.toString()) {
        return { success: true, email, fullName, role, locAccess, siteAccess, isNew: true };
      }
    } else {
      if (hashPassword(password, salt) === storedPass) {
        return { success: true, email, fullName, role, locAccess, siteAccess, isNew: false };
      }
    }
    return { error: "Incorrect password." };
  } catch (e) { return { error: e.toString() }; }
}

function finalizePassword(email, newPassword) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    const target = email.toString().trim().toLowerCase();
    const rowIndex = data.findIndex(r => r[0] && r[0].toString().trim().toLowerCase() === target);

    if (rowIndex === -1) return "User verification failed.";

    const salt = Utilities.getUuid();
    const hashedPassword = hashPassword(newPassword, salt);

    sheet.getRange(rowIndex + 1, 3, 1, 2).setValues([[hashedPassword, salt]]);
    SpreadsheetApp.flush();
    return "SUCCESS";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

function hashPassword(password, salt) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt);
  return digest.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

function recoverPassword(identifier) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    const input = identifier.toString().trim().toLowerCase();

    const rowIndex = data.findIndex(row =>
      (row[0] && row[0].toString().trim().toLowerCase() === input) ||
      (row[1] && row[1].toString().trim().toLowerCase() === input)
    );

    if (rowIndex === -1) return "❌ Name or Email not found.";

    const email = data[rowIndex][0];
    const fullName = data[rowIndex][1];

    const tempPass = Math.random().toString(36).slice(-8).toUpperCase();

    sheet.getRange(rowIndex + 1, 3).setValue(tempPass);
    sheet.getRange(rowIndex + 1, 4).setValue("");
    SpreadsheetApp.flush();

    try {
      const subject = "Inventory System - Password Reset";
      const body = `Hello ${fullName},\n\nYour password has been reset.\n\nTemporary Password: ${tempPass}\n\nPlease log in and set a new password.\n\nRegards,\nInventory System`;

      GmailApp.sendEmail(email, subject, body);

      return "✅ SUCCESS: A temporary password has been sent to your registered email address.";
    } catch (mailErr) {
      console.error("Mail Failure: " + mailErr.toString());
      return "⚠️ Reset worked, but EMAIL FAILED. Your temporary password is: " + tempPass;
    }
  } catch (e) {
    return "❌ Error: " + e.toString();
  }
}
