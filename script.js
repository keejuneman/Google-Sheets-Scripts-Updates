function moveToNextWeek() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Your Sheet Name");
  var currentRow = sheet.getRange("A2").getValue();
  var dayOfWeek = new Date().getDay();
  var a3Value = sheet.getRange("A3").getValue(); 

  if (dayOfWeek === 1 & a3Value === 'no') {
    currentRow = currentRow + 13;
  }
  
  var cell = sheet.getRange(currentRow, 3);
  sheet.setActiveRange(cell);

  sheet.getRange("A2").setValue(currentRow);
  sheet.getRange("A2").setFontColor("#ffffff");

  if (a3Value === "no" & dayOfWeek === 1) {
    sheet.getRange("A3").setValue("yes");
    sheet.getRange("A3").setFontColor("#ffffff"); 

  }

  if (a3Value === "yes" & dayOfWeek !== 1) {
    sheet.getRange("A3").setValue("no");
    sheet.getRange("A3").setFontColor("#ffffff"); 

  }
}
