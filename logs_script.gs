function logActivity(functionName, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Logs_Script_Activity");

  const timestamp = new Date();
  const userEmail = Session.getActiveUser().getEmail();

  logSheet.appendRow([
    timestamp,
    functionName,
    userEmail,
    details
  ]);
}
