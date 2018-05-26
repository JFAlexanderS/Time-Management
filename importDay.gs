/*
  Imports all events in a day from the active calendar into a Google Spreadsheet.
*/

function importDay() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var today = new Date();

  // Change to the name of the calendar you would like to import from
  const calendar = ""

  // Pulls events from yesterday instead of today.
  // today.setDate(today.getDate()-1);

  var events = CalendarApp.getCalendarsByName(calendar)[0].getEventsForDay(today);
  timezone = CalendarApp.getCalendarsByName(calendar)[0].getTimeZone();
  
  for(var i = 0; i < events.length; i++)
    sheet.appendRow([events[i].getTitle(), events[i].getDescription(), 
                     Utilities.formatDate(events[i].getStartTime(), timezone, "HH:mm:ss"),
                     Utilities.formatDate(events[i].getEndTime(), timezone, "HH:mm:ss"),
                     (events[i].getEndTime()-events[i].getStartTime())/(60*60*1000),
                     Utilities.formatDate(events[i].getStartTime(), timezone, "yyyy-MM-dd")])
    
  sheet.sort(3)
}

