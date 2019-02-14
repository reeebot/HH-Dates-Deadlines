function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('input');
  var cell = sheet.getRange('C2');
  sheet.setCurrentCell(cell);
}


function DivideDates() {
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

  // switch to finished report tab
  SpreadsheetApp.setActiveSheet(ss.getSheets()[1]);
  // msg
  Browser.msgBox("Dates and Deadlines Created Successfully");
};


function Refresh() {
  SpreadsheetApp.flush();
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
  // change sheet to input
  SpreadsheetApp.setActiveSheet(SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]);
  // msg
  Browser.msgBox("Form Has Been Cleared");
}


function createCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('calendar');
  var calendar = CalendarApp.createCalendar(sheet.getRange('A2').getValue());  // Calendar name
  Logger.log(calendar);
  
  var startRow = 2;  // First row of data to process / exempts header row
  var numRows = sheet.getLastRow();   // Number of rows to process
  var numColumns = sheet.getLastColumn();
  
  var dataRange = sheet.getRange(startRow, 1, numRows-1, numColumns); 
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];  // row of data
    var name = row[1];  // Event Name
    //var date = new Date(row[2]);  // Event date
    var fdate = Utilities.formatDate(new Date(row[2]), "GMT-7", "MMMM dd, yyyy HH:mm:ss Z");
    var date = new Date(fdate);
    Logger.log(fdate)
    Logger.log(date)
    var event = calendar.createAllDayEvent(name, date);
    event.addPopupReminder(900); // Reminder Popup at 9am day prior
    
    
    // #### TODO #### //
    // calendar API:
    // - add to notifications agenda
    // - override default reminders?
    
    
  };
  // msg
  Browser.msgBox("Appointments Added to Calendar");
};



function saveForm() {
  
  // #### TODO #### //
  // - make a copy of spreadsheet with last name + address in filename
  
  //var sheet = ss.getSheetByName('Template').copyTo(ss);
  
  
  // #### TODO #### //  
  // - save report as pdf & jpg
  
};




 
/**
 * Lists 10 upcoming events in the user's calendar.
 */
function listUpcomingEvents() {
  var calendarId = 'Moser - 7302 Ocean Ridge Street, Wellington, CO 80549';
  var optionalArgs = {
    timeMin: (new Date()).toISOString(),
    showDeleted: false,
    singleEvents: true,
    maxResults: 10,
    orderBy: 'startTime'
  };
  var response = Calendar.Events.list(calendarId, optionalArgs);
  var events = response.items;
  if (events.length > 0) {
    for (i = 0; i < events.length; i++) {
      var event = events[i];
      var when = event.start.dateTime;
      if (!when) {
        when = event.start.date;
      }
      Logger.log('%s (%s)', event.summary, when);
    }
  } else {
    Logger.log('No upcoming events found.');
  }
}





