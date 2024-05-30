function copyAndRenameFiles() {
    // 스프레드시트와 특정 시트를 가져옵니다.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    
    // E4 셀의 구글 드라이브 폴더 주소를 가져옵니다.
    var folderUrl = sheet.getRange("E4").getValue();
    var folderId = folderUrl.match(/[-\w]{25,}/);
    var folder = DriveApp.getFolderById(folderId);
    
    // E7 셀의 파일 주소를 가져옵니다.
    var fileUrl = sheet.getRange("E7").getValue();
    var fileId = fileUrl.match(/[-\w]{25,}/);
    var file = DriveApp.getFileById(fileId);
    
    // A열과 C열의 데이터 범위를 가져옵니다.
    var names = sheet.getRange("A3:A").getValues().filter(String);
    var linksRange = sheet.getRange("C3:C" + (names.length + 2));
    
    var links = [];
    
    // 각 이름에 대해 파일의 사본을 생성하고 이름을 변경합니다.
    for (var i = 0; i < names.length; i++) {
      var newName = names[i][0];
      var newFile = file.makeCopy(newName, folder);
      var newFileUrl = newFile.getUrl();
      
      // C열에 새 파일의 링크를 추가합니다.
      links.push([newFileUrl]);
    }
    
    // C3:C 셀에 링크를 기입합니다.
    linksRange.setValues(links);
  }
  