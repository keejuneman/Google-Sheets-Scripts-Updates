var DriveApp = DriveApp;
var GmailApp = GmailApp;

function SendMail() {
  try {
    var activeSS = SpreadsheetApp.getActiveSpreadsheet();
    var activeSheet = activeSS.getActiveSheet();

    // G27 셀에서 서명 가져오기
    var signature = activeSheet.getRange("H27").getValue();

    for(let index = 3; index < 200; index++) { 
      var email = activeSheet.getRange("B"+index).getValue(); 
      var carboncopy = activeSheet.getRange("C"+index).getValue(); 
      var mail_subject = activeSheet.getRange("D"+index).getValue(); 
      var mail_body = activeSheet.getRange("E"+index).getValue();
      var file_urls = activeSheet.getRange("F"+index).getValue(); // 파일 URL들 가져오기(줄바꿈으로 구분)
      
      // 이메일 본문에 서명 추가
      mail_body += "<br><br>--<br>" + signature;

      if (email !== "") {
        var attachments = []; // 첨부 파일 배열 초기화
        
        // 파일 URL이 하나 이상 있으면 각각 처리
        if (file_urls !== "") {
          var urls = file_urls.split('\n'); // 줄바꿈으로 URL 분리
          for(var i = 0; i < urls.length; i++) {
            var fileId = getFileIdFromUrl(urls[i]); // 파일 ID 추출
            if(fileId) {
              var file = DriveApp.getFileById(fileId); // 파일 가져오기
              attachments.push(file); // 첨부 파일로 추가
            }
          }
        }

        // 이메일 보내기
        MailApp.sendEmail({
          to: email, 
          subject: mail_subject, 
          htmlBody: mail_body,
          attachments: attachments, // 첨부 파일 추가
          cc: carboncopy
        });

        console.log(email, mail_subject, "--> 발송 성공");
        
      } else {
        console.log("B" + index + "에 메일 주소 없음. 발송 중단 처리");
        break; 
      }              
    }
  } catch(err) {
    console.log("발송 실패 - " + err);
  }
}


// function SendMail() {
//   try {
//     var activeSS = SpreadsheetApp.getActiveSpreadsheet();
//     var activeSheet = activeSS.getActiveSheet();

//     // 현재 활성 사용자의 이메일 주소 가져오기
//     var user = Session.getActiveUser().getEmail();
//     // 사용자의 'SendAs' 설정 가져오기
//     var sendAs = GmailApp.getSendAs()[0]; // 첫 번째 SendAs 설정 사용
//     // 서명 정보 가져오기
//     var signature = sendAs.getSignature();

//     for(let index = 3; index <= 200; index++) { 
//       var email = activeSheet.getRange("B"+index).getValue(); 
//       var carboncopy = activeSheet.getRange("C"+index).getValue(); 
//       var mail_subject = activeSheet.getRange("D"+index).getValue(); 
//       var mail_body = activeSheet.getRange("E"+index).getValue();
//       var file_urls = activeSheet.getRange("F"+index).getValue(); // 파일 URL들 가져오기(줄바꿈으로 구분)

//       // 이메일 본문에 서명 추가
//       mail_body += "\n\n--\n" + signature;

//       if (email !== "") {
//         var attachments = []; // 첨부 파일 배열 초기화
        
//         // 파일 URL이 하나 이상 있으면 각각 처리
//         if (file_urls !== "") {
//           var urls = file_urls.split('\n'); // 줄바꿈으로 URL 분리
//           for(var i = 0; i < urls.length; i++) {
//             var fileId = getFileIdFromUrl(urls[i]); // 파일 ID 추출
//             if(fileId) {
//               var file = DriveApp.getFileById(fileId); // 파일 가져오기
//               attachments.push(file); // 첨부 파일로 추가
//             }
//           }
//         }

//         // 이메일 보내기
//         GmailApp.sendEmail(email, mail_subject, mail_body, {
//           attachments: attachments, // 첨부 파일 추가
//           cc: carboncopy
//         });

//         console.log(email, mail_subject, "--> 발송 성공");
//         SpreadsheetApp.getUi().alert("발송 성공");
//       } else {
//         console.log("B" + index + "에 메일 주소 없음. 발송 중단 처리");
//         break; 
//       }              
//     }
//   } catch(err) {
//     console.log("발송 실패 - " + err);
//   }
// }


// Google Drive URL에서 파일 ID 추출하는 함수
function getFileIdFromUrl(url) {
  var match = /\/d\/([a-zA-Z0-9-_]+)/.exec(url);
  return match ? match[1] : null;
}


function getFileAddress() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange('A3:A' + sheet.getLastRow()); // Assuming data starts from row 2 in column A
  var values = range.getValues();
  var driveURL = sheet.getRange('H4').getValue(); // Assuming the Google Drive URL is in cell G4

  for (var i = 0; i < values.length; i++) {
    var folderName = values[i][0];
    var folder = getFolder(driveURL, folderName);
    if (folder) {
      var files = folder.getFiles();
      if (files.hasNext()) {
        var file = files.next();
        var fileURL = file.getUrl();
        sheet.getRange(i + 3, 6).setValue(fileURL); // Assuming column E is column 5
      }
    }
  }
}

function getFolder(driveURL, folderName) {
  var folderId = getFolderIdFromURL(driveURL);
  var folder = DriveApp.getFolderById(folderId);
  var subFolders = folder.getFoldersByName(folderName);
  if (subFolders.hasNext()) {
    return subFolders.next();
  }
  return null;
}

function getFolderIdFromURL(url) {
  var id = null;
  var match = url.match(/[-\w]{25,}/);
  if (match) {
    id = match[0];
  }
  return id;
}

function deleteValues_Gmail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange('A3:F').clearContent();
}
