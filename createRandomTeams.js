function createTeams() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var names = sheet.getRange("A2:A").getValues().filter(String).map(function(row){return row[0];});
    var formationCriterion = sheet.getRange("B2").getValue();
    var criterionValue = sheet.getRange("C2").getValue();
    var teams = [];
    var result = "";
  
    sheet.getRange("D2").clearContent();
    if (names.length == 0 || criterionValue == 0) {
      sheet.getRange("D2").setValue("팀 편성을 위한 데이터가 충분하지 않습니다.");
      return;
    }
    
    // Shuffle names randomly
    names = names.sort(function() { return 0.5 - Math.random(); });
    
    if (formationCriterion == "팀 개수") {
      var membersPerTeam = Math.ceil(names.length / criterionValue);
      for (var i = 0; i < names.length; i++) {
        if (i % membersPerTeam == 0) teams.push([]);
        teams[teams.length - 1].push(names[i]);
      }
    } else if (formationCriterion == "팀 별 구성원 수") {
      for (var i = 0; i < names.length; i++) {
        if (i % criterionValue == 0) teams.push([]);
        teams[teams.length - 1].push(names[i]);
      }
    } else {
      sheet.getRange("D2").setValue("올바른 편성 기준을 선택해주세요.");
      return;
    }
    
    for (var i = 0; i < teams.length; i++) {
      result += (i+1) + "조 : " + teams[i].join(", ") + "\n";
    }
    
    sheet.getRange("D2").setValue(result);
  }
  