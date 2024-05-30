function extractUncheckedPeople() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getRange('A3:B300'); 
    var values = range.getValues();
    
    var uncheckedPeopleWithAt = [];
    var uncheckedPeopleWithoutAt = [];
    
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] && !values[i][1]) {
        uncheckedPeopleWithAt.push('@' + values[i][0]); 
        uncheckedPeopleWithoutAt.push(values[i][0]); 
      }
    }
    
    sheet.getRange('E3').setValue(uncheckedPeopleWithAt.join(', '));
    sheet.getRange('E7').setValue(uncheckedPeopleWithoutAt.join(', '));
  }
  
  function deleteValues() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    sheet.getRange('C3:C300').clearContent();
    sheet.getRange('E3').clearContent();
    sheet.getRange('E7').clearContent();
  }
  