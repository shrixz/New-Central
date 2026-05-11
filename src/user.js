/**
 * Authentication & Security Logic
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

    const email = data[rowIndex][0];
    const fullName = data[rowIndex][1];
    const storedPass = data[rowIndex][2];
    const salt = data[rowIndex][3];
    
    // Pull Access Control Columns. Defaulting to team leader if empty.
    const role = data[rowIndex][4] ? data[rowIndex][4].toString().toLowerCase() : "team leader";
    const locAccess = data[rowIndex][5] ? data[rowIndex][5].toString().trim() : "";
    const siteAccess = data[rowIndex][6] ? data[rowIndex][6].toString().trim() : "";
    
    const isNewUser = (!salt || salt.toString().trim() === "");

    if (isNewUser) {
      if (password.toString() === storedPass.toString()) {
        return { success: true, fullName, role, locAccess, siteAccess, isNew: true };
      }
    } else {
      if (hashPassword(password, salt) === storedPass) {
        return { success: true, fullName, role, locAccess, siteAccess, isNew: false };
      }
    }
    return { error: "Incorrect password." };
  } catch (e) { return { error: e.toString() }; }
}

function finalizePassword(name, newPassword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[1].toString().trim() === name.toString().trim());
  
  if (rowIndex === -1) return "User verification failed.";

  const salt = Utilities.getUuid(); 
  const hashedPassword = hashPassword(newPassword, salt);
  
  sheet.getRange(rowIndex + 1, 3, 1, 2).setValues([[hashedPassword, salt]]);
  SpreadsheetApp.flush();
  return "SUCCESS";
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
    
    sheet.getRange(rowIndex + 1, 3, 1, 2).setValues([[tempPass, ""]]);
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