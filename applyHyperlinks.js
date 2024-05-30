function applyHyperlinksToNonEmptyCells() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getRange("A:A");
    var values = range.getValues(); 
    
    for (var i = 0; i < values.length; i++) {
      var cell = values[i][0];
      if (cell) { 
        sheet.getRange(i+1, 1).setFormula('=HYPERLINK("' + cell + '", "> 링크")');
      }
    }
  }
  