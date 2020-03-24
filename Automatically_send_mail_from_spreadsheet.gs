/* Automatically send an e-mail with an attached file
***Use with Google spreadsheet***
Column 1: e-mail address
Column 2: subject
Column 3: message
Column 4: attached file name
*/
// Change this folder name to match the location of attached files
var myFileParentFolderName = "2019-2_207108_sec06_optics2";

function onOpen() {
 var ss = SpreadsheetApp.getActiveSpreadsheet(),
     options = [
      {name:"Send Mail", functionName:"sendEmails2"},
     ];
 ss.addMenu("Email Sender", options);
}

var EMAIL_SENT = "EMAIL_SENT";

// Send an email with one attachment: a file from Google Drive (as a PDF)             
function sendEmails2() {  
  var debug = 0;       
  var numCols = 4;
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process, skip header
  var numRows = 20;   // Number of rows to process
  // Fetch the range of cells A2:(numRows,numCols)
  var dataRange = sheet.getRange(startRow, 1, numRows, numCols)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0];  // First column
    var subject = row[1];       // Second column
    var message = row[2];       // Third column
    var attachedFile = row[3];     // Fourth column
    if (debug) {
      Logger.log(subject,message,attachedFile)
    }
    var emailSent = row[4];     // Ffifth column    
  /* First method: use file iterator*/
//    var files = DriveApp.getFilesByName(attachedFile);
//    if (files.hasNext()) {
//      var file = files.next();
//      Logger.log("id=%s, name=%s", file.getId(), file.getName())
//      MailApp.sendEmail(emailAddress, subject, message, {attachments: [file.getBlob()],name: 'Nithiwadee Thaicharoen'})
//    }
  /* Second method: if it's inside a folder*/
    if (emailSent != EMAIL_SENT)  {  // Prevents sending duplicates    
      var folders = DriveApp.getFoldersByName(myFileParentFolderName)
      while(folders.hasNext()){
        var files = DriveApp.getFilesByName(attachedFile);
        if (files.hasNext()) {
          var file = files.next();
          Logger.log("id=%s, name=%s", file.getId(), file.getName())
          MailApp.sendEmail(emailAddress, subject, message, {attachments: [file.getBlob()],name: 'Nithiwadee Thaicharoen'})
        }
        folder = folders.next().getName(); 
        Logger.log(folder)      
      }   
      sheet.getRange(startRow + i, 5).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
