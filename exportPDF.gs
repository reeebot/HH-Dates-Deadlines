// # TO-DO #
// - cloudconvert googlecloud integration


function exportPDF() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var urlID = ss.getId();
  var ssID = DriveApp.getFileById(urlID);
  
  var backupsFolder = DriveApp.getFolderById("13_ozP4EGLMv9ECiafWKhgzBGGHR18M5r"); // backups folder
  var projectName = ss.getSheetByName('calendar').getRange('A2').getValue(); // project name
  var projectFolder = backupsFolder.createFolder(projectName); // create project folder
  var backupSS = ssID.makeCopy(projectName, projectFolder); // create spreadsheet backup
  
  var lastRowNum = ss.getSheetByName('pivot').getLastRow() + 6;
  var sheets = ss.getSheets();
  var gidID = sheets[1].getSheetId()
  
  var url = "https://docs.google.com/spreadsheets/d/"+urlID+"/export?gid="+gidID+
    "&exportFormat=pdf&format=pdf"+
    "&portrait=true&single=true"+
    "&top_margin=0&bottom_margin=0&left_margin=0&right_margin=0&fitw=true"+
    "&ir=false&ic=false"+
    "&r1=0&c1=0"+"&r2="+lastRowNum+"&c2=12"
                                                        
  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  
  var response = UrlFetchApp.fetch(url, params).getBlob().setName(projectName);
  // save to drive
  projectFolder.createFile(response);
  
  var input = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('input');
  input.setCurrentCell(input.getRange('D11'));
  // msg
  Browser.msgBox("Form & PDF Saved to Google Drive");
  
}
