function onCreateKitchenManagerSubmit(e) {
  // Prevent cross-trigger execution
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "Kitchen_Managers_Responses") return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const responseSheet = ss.getSheetByName("Kitchen_Managers_Responses");
  const targetSheet = ss.getSheetByName("Kitchen_Managers");
  const kitchensSheet = ss.getSheetByName("Kitchens");

  // Extract form values
  const timestamp = e.values[0];
  const managerName = e.values[1];
  const managerEmail = e.values[2];
  const kitchenAssignedName = e.values[3];

  // Lookup Kitchen_ID from Kitchens tab
  const kitchensData = kitchensSheet.getDataRange().getValues();
  let kitchenId = "";

  for (let i = 1; i < kitchensData.length; i++) {
    if (kitchensData[i][1] === kitchenAssignedName) {
      kitchenId = kitchensData[i][0];
      break;
    }
  }

  // Generate Manager_ID
  const lastRow = targetSheet.getLastRow();
  let newManagerID = "MGR_001";

  if (lastRow > 1) {
    const lastID = targetSheet.getRange(lastRow, 1).getValue();
    const lastNumber = parseInt(lastID.split("_")[1]);
    newManagerID = "MGR_" + String(lastNumber + 1).padStart(3, "0");
  }

  // Append structured row
  targetSheet.appendRow([
    newManagerID,
    managerName,
    managerEmail,
    kitchenId,
    kitchenAssignedName,
    timestamp
  ]);

  // Log the event
  logActivity(
    "onCreateKitchenManagerSubmit",
    `Assigned manager ${newManagerID} (${managerName}) to kitchen ${kitchenId} (${kitchenAssignedName})`
  );
}
