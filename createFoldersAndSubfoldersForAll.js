function createFoldersAndSubfoldersForAll() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // E4 셀에서 부모 드라이브 링크 읽기
    var parentFolderUrl = sheet.getRange("E4").getValue();
    var parentFolderId = parentFolderUrl.match(/[-\w]{25,}/)[0]; // URL에서 폴더 ID 추출
    var parentFolder = DriveApp.getFolderById(parentFolderId); // 부모 폴더 객체 가져오기
    
    // A3:A와 B3:C 범위의 데이터 읽기
    var mainFolderNames = sheet.getRange("A3:A").getValues();
    var subFolderNamesArray = sheet.getRange("B3:C").getValues();
    
    for (var i = 0; i < mainFolderNames.length; i++) {
      if (mainFolderNames[i][0] === "") break; // A열에 값이 없으면 반복 중단
      
      var mainFolderName = mainFolderNames[i][0];
      var subFolderNamesString = subFolderNamesArray[i][0];
      
      if (subFolderNamesString === "") continue; // B열에 값이 없으면 이번 루프 건너뛰기
      
      // 줄바꿈을 기준으로 하위 폴더 이름 배열 생성
      var subFolderNames = subFolderNamesString.split('\n');
      
      // 메인 폴더 생성 (이미 존재하면 해당 폴더 사용)
      var mainFolder;
      var mainFolders = parentFolder.getFoldersByName(mainFolderName);
      if (mainFolders.hasNext()) {
        mainFolder = mainFolders.next();
      } else {
        mainFolder = parentFolder.createFolder(mainFolderName);
      }
      
      // 메인 폴더 링크 C열에 기입
      sheet.getRange(i + 3, 3).setValue(mainFolder.getUrl());
      
      // 하위 폴더들 생성
      for (var j = 0; j < subFolderNames.length; j++) {
        var subFolderName = subFolderNames[j].trim(); // 공백 제거
        if (subFolderName === "") continue; // 비어있는 이름 건너뛰기
        
        // 하위 폴더 생성 (이미 존재하면 건너뛰기)
        var subFolders = mainFolder.getFoldersByName(subFolderName);
        if (!subFolders.hasNext()) {
          mainFolder.createFolder(subFolderName);
        }
      }
    }
  }
  
  function deleteValues_Drive2() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    sheet.getRange('A3:C').clearContent();
  }
  