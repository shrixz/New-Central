// temporary function to force the permission popup
function FORCE_PERMISSION_POPUP() {
  // 1. Trigger Spreadsheet permission
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log("Spreadsheet Access: OK");

  // 2. Trigger Drive permission (for your logo)
  const folders = DriveApp.getFolders();
  console.log("Drive Access: OK");

  // 3. Trigger Gmail permission (The one failing)
  GmailApp.sendEmail(Session.getActiveUser().getEmail(), "Permission Test", "If you see this, permissions are granted.");
  console.log("Gmail Access: OK");
  
  Browser.msgBox("Success! Permissions are now fully authorized.");
}