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
  
  var GETurl = "https://docs.google.com/spreadsheets/d/"+urlID+"/export?gid="+gidID+
    "&exportFormat=pdf&format=pdf"+
    "&portrait=true&single=true"+
    "&top_margin=0&bottom_margin=0&left_margin=0&right_margin=0&fitw=true"+
    "&ir=false&ic=false"+
    "&r1=0&c1=0"+"&r2="+lastRowNum+"&c2=12"
                                                        
  var GETparams = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  
  // download pdf
  var GETresponse = UrlFetchApp.fetch(GETurl, GETparams).getBlob().setName(projectName);
  // save to drive
  var pdfFile = projectFolder.createFile(GETresponse);
  var pdfBlob = pdfFile.getBlob();  // for conversion to jpg
  
  
  //////////////////////////
  // convert pdf to jpg using convertapi
  
  var POSTpayload = {
    "Parameters": [
        {
            "Name": "File",
            "FileValue": {
                "Name": projectName + ".pdf",
                "Data": Utilities.base64Encode(pdfBlob.getBytes())
            }
        }
    ]
  };
  
  var POSToptions = {
    'method' : 'post',
    'payload' : JSON.stringify(POSTpayload),
    'contentType' : 'application/json'
  };
  
  var POSTresponse = JSON.parse(UrlFetchApp.fetch('https://v2.convertapi.com/convert/pdf/to/jpg?Secret=uDOxzsD1OCMYbQ02', POSToptions).getContentText());
  
  // extract base64 filedata from json response & convert into jpg
  var jpgRaw = POSTresponse.Files[0].FileData;
  var jpgFile = Utilities.newBlob(Utilities.base64Decode(jpgRaw),  MimeType.JPEG, projectName+".jpg").getAs('image/jpeg');
  
  // save jpg to drive
  projectFolder.createFile(jpgFile);
  
  
  ////////////////////////////
  // go to input D11
  var input = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('input');
  input.setCurrentCell(input.getRange('D11'));
  // msg
  Browser.msgBox("PDF / JPG / Form Saved to Google Drive");
}
