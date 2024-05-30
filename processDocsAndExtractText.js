function processDocsAndExtractText() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var namesRange = sheet.getRange("A4:A");
    var names = namesRange.getValues().flat().filter(function(name) {
      return name !== ""; // 공백이 아닌 이름만 저장
    });
    
    var folderUrl = sheet.getRange("C3").getValue();
    var folderId = folderUrl.match(/[-\w]{25,}/);
    var folder = DriveApp.getFolderById(folderId[0]);
    var files = folder.getFiles();
    var extractedTexts = [];
    
    while (files.hasNext()) {
      var file = files.next();
      if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
        var doc = DocumentApp.openById(file.getId());
        var text = doc.getBody().getText();
        // 줄바꿈을 기준으로 텍스트 분리하여 extractedTexts 리스트에 추가
        var lines = text.split('\n').filter(function(line) { return line.trim() !== ''; });
        extractedTexts = extractedTexts.concat(lines);
      }
    }
  
    var processedTexts = extractedTexts
      .filter(function(text) {
        return /[가-힣]/.test(text); // 한글이 포함된 요소만 남김
      })
      .map(function(text) {
        return text.replace(/[^가-힣]/g, ""); // 한글이 아닌 글자들 제거
      })
      .filter(function(value, index, self) {
        return self.indexOf(value) === index; // 중복 요소 제거
      })
      .filter(function(text) {
        return names.includes(text); // 임시 리스트1에 존재하는 항목만 남김
      })
      .sort(); // 이름순으로 정렬
    
    // 결과를 B4:B셀에 출력하는 부분을 수정
    if (processedTexts.length > 0) { // 처리된 텍스트가 있을 경우에만 실행
      var outputRange = sheet.getRange(4, 2, processedTexts.length, 1);
      outputRange.setValues(processedTexts.map(function(name) { return [name]; }));
    } else {
      // 처리된 텍스트가 없을 경우, 사용자에게 알림을 줄 수 있습니다.
      SpreadsheetApp.getUi().alert('처리된 텍스트가 없습니다.');
    }
  }
  