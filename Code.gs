function onOpen() {
  var input = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('input');
  var startCell = input.getRange('C2');
  input.setCurrentCell(startCell);
  // msg
  Browser.msgBox("Welcome to the -- Hops & Homes -- Dates & Deadlines Form Creation Suite");
}



function divideDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MAIN_FORM'); // apply to sheet name only
  var rows = sheet.getRange('A10:L48'); // range to apply formatting to
  var numRows = rows.getNumRows(); // no. of rows in the range named above
  var testvalues = sheet.getRange('B1:B48').getValues(); // array of values to be tested (1st column of the range named above)

  rows.setBorder(false, false, false, false, false, false, "red", SpreadsheetApp.BorderStyle.SOLID_MEDIUM); // remove existing borders before applying rule below

  for (var i = 0; i <= numRows - 1; i++) {
    var n = i + 1;
    if (testvalues[i] > testvalues[i+1]) { // test applied to array of values
      sheet.getRange('a' + n + ':l' + n).setBorder(true, null, null, null, null, null, "#d7e3c1", SpreadsheetApp.BorderStyle.SOLID_THICK); // format if true
    };
  };
  // remove last horizontal line
  var sheetPivot = ss.getSheetByName('pivot');
  var lastRowNumPivot = sheetPivot.getLastRow() + 6;
  var lastRowRange = sheet.getRange('A' + lastRowNumPivot + ':L' + lastRowNumPivot).setBorder(false, false, false, false, false, false, "red", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // save form & export PDF
  exportPDF();
};



function Reset() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('input');
  sheet.getRange('C2:C8').clearContent();
  sheet.getRange('C12:C60').clearContent();
  sheet.getRange('B12:B60').clearContent();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MAIN_FORM'); // apply to sheet name only
  var rows = sheet.getRange('A10:L48'); // range to apply formatting to
  var numRows = rows.getNumRows(); // no. of rows in the range named above
  var testvalues = sheet.getRange('B1:B48').getValues(); // array of values to be tested (1st column of the range named above)

  rows.setBorder(false, false, false, false, false, false, "red", SpreadsheetApp.BorderStyle.SOLID_MEDIUM); // remove existing borders before applying rule below

  var input = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('input');
  input.setCurrentCell(input.getRange('C2'));
  // msg
  Browser.msgBox("Form Has Been Cleared");
}



function createCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('calendar');
  var calendarName = sheet.getRange('A2').getValue();
  var calendar = CalendarApp.createCalendar(calendarName).setTimeZone("America/Denver");  // Create Calendar with specified name in Denver timezone

  var startRow = 2;  // First row of data to process / exempts header row
  var numRows = sheet.getLastRow();   // Number of rows to process
  var numColumns = sheet.getLastColumn();
  
  var dataRange = sheet.getRange(startRow, 1, numRows-1, numColumns); 
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];  // row of data
    var name = row[1];  // Event Name
    var fdate = Utilities.formatDate(new Date(row[2]), "America/Denver", "MMMM dd, yyyy HH:mm:ss Z");
    var date = new Date(fdate);
    
    var event = calendar.createAllDayEvent(name, date);
    event.addPopupReminder(900); // Reminder Popup at 9am day prior
  };
  
  // calendarAPI: add calendar to daily email notifications agenda
  var notificationArgs = {
    "notificationSettings": {
      "notifications": [
        {
          "method": "email",
          "type": "agenda"
        }
      ]
    }
  }
  var calendarId = calendar.getId();
  Calendar.CalendarList.update(notificationArgs, calendarId);
  
  // msg
  Browser.msgBox("Appointments Added to Calendar");
};
