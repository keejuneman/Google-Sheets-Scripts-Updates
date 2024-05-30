function createFolderAndCopyFile() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataRange = sheet.getRange('A3:C' + sheet.getLastRow());
    var data = dataRange.getValues();
    
    var parentFolderLink = sheet.getRange('F4').getValue();
    var originalFileLink = sheet.getRange('F7').getValue();
    
    // 부모 폴더와 원본 파일의 ID 추출
    var parentFolderId = parentFolderLink.match(/[-\w]{25,}/);
    var originalFileId = originalFileLink.match(/[-\w]{25,}/);
    
    if (parentFolderId && originalFileId) {
      var parentFolder = DriveApp.getFolderById(parentFolderId[0]);
      var originalFile = DriveApp.getFileById(originalFileId[0]);
      
      data.forEach(function(row, index) {
        var name = row[0]; // 이름
        var folderName = row[1]; // 폴더 이름
        var newFileName = row[2]; // 생성 파일 이름
        
        // 모든 필드가 채워져 있는지 확인
        if (folderName && newFileName) {
          // 폴더 생성
          var folder;
          var folders = parentFolder.getFoldersByName(folderName);
          if (folders.hasNext()) {
            folder = folders.next();
          } else {
            folder = parentFolder.createFolder(folderName);
          }
          
          // 원본 파일 사본 생성 및 이름 변경
          var newFile = originalFile.makeCopy(newFileName, folder);
          
          // 사본 파일의 URL을 시트의 D열에 저장
          var fileUrl = newFile.getUrl();
          sheet.getRange('D' + (index + 3)).setValue(fileUrl);
        }
      });
      
      SpreadsheetApp.flush(); // 변경사항 적용
      SpreadsheetApp.getUi().alert('폴더 및 사본 파일 생성 완료');
    } else {
      Logger.log('잘못된 링크입니다. 부모 폴더 또는 원본 파일의 링크를 확인해주세요.');
      SpreadsheetApp.getUi().alert('잘못된 링크입니다. 부모 폴더 또는 원본 파일의 링크를 확인해주세요.');
  
    }
  }
  
  
  
  function createFolder() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataRange = sheet.getRange('A3:C' + sheet.getLastRow());
    var data = dataRange.getValues();
    
    var parentFolderLink = sheet.getRange('F4').getValue();
    
    // 부모 폴더와 원본 파일의 ID 추출
    var parentFolderId = parentFolderLink.match(/[-\w]{25,}/);
    
    if (parentFolderId) {
      var parentFolder = DriveApp.getFolderById(parentFolderId[0]);
      
      data.forEach(function(row, index) {
        var name = row[0]; // 이름
        var folderName = row[1]; // 폴더 이름
        
        // 모든 필드가 채워져 있는지 확인
        if (folderName) {
          // 폴더 생성
          var folder;
          var folders = parentFolder.getFoldersByName(folderName);
          if (folders.hasNext()) {
            folder = folders.next();
          } else {
            folder = parentFolder.createFolder(folderName);
          }
        }
      });
      
      SpreadsheetApp.flush(); // 변경사항 적용
      SpreadsheetApp.getUi().alert('폴더 생성 완료');
    } else {
      Logger.log('잘못된 링크입니다. 부모 폴더 또는 원본 파일의 링크를 확인해주세요.');
      SpreadsheetApp.getUi().alert('잘못된 링크입니다. 부모 폴더 또는 원본 파일의 링크를 확인해주세요.');
  
    }
  }
  
  function deleteValues_Drive1() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    sheet.getRange('A3:D').clearContent();
  }