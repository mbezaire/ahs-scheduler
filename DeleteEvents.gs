function DeleteEvents() {
  /*  The DeleteEvents function in Calendar Scheduler takes input from
      the AHS Personal Scheduler Google Spreadsheet and then deletes the
      appointments for each class on each school day from the specified
      Google Calendar. It only deletes events made by CreateEvents, denoted
      by the text "This event was created by the webapp AHS Personal Calendar."
      in the description of the event.

      Author: marianne.bezaire@andoverma.us
      Date Updated: September 4, 2023

      Github code: https://github.com/mbezaire/ahs-scheduler
  */


  // Prevent script running if another is already running:
  // https://stackoverflow.com/questions/67066779/how-to-prevent-google-apps-script-trigger-if-a-function-is-already-running
  var isItRunning;

  isItRunning = CacheService.getDocumentCache().put("itzRunning", "true",600);//Keep this value in Cache for up to X minutes
  //There are 3 types of Cache - if the Apps Script project is not
  //bound to a document then use ScriptCache
  if (isItRunning) {//If this is true then another instance of this
    //function is running which means that you dont want this
    //instance of this function to run - so quit
    return;//Stop running this instance of this function
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Spreadsheet = " + ss);

  var nowdate = new Date();

  var bnames = ss.getSheets()[0];
  var range = bnames.getRange("M6:N7");
  var TermDates = range.getValues();
  Logger.log(TermDates);

  var range = bnames.getRange("H5");
  if (range.getValue()=="All") {
    var Term = 9;
  } else {
    var Term = range.getValue();
  }

  if (Term==1) {
    var startYear = new Date(TermDates[0][0]);
    var endYear = new Date(TermDates[0][1]);
  } else if (Term==2) {
    var startYear = new Date(TermDates[1][0]);
    var endYear = new Date(TermDates[1][1]);
  } else if (Term==9) {
    var startYear = new Date(TermDates[0][0]);
    var endYear = new Date(TermDates[1][1]);
  }
  Logger.log(startYear + " - " + endYear);

  var bnames = ss.getSheets()[0];
  var range = bnames.getRange("H3");
  var CalName = range.getValue();

  var range = bnames.getRange("H4");
  var updateAll = (range.getValue()=='All');

  var range = bnames.getRange("N5");
  var sleepTime = range.getValue();

  Logger.log("Will be deleting events from Calendar: " + CalName);

  var events = CalendarApp.getCalendarsByName(CalName)[0].getEvents(startYear, endYear,
    {search: "This event was created by the webapp AHS Personal Calendar."});
  Logger.log("Events = " + events.length);
  bnames.getRange("H10").setValue("You'll need to run the script again; still working.");

  for (e=0; e<events.length; e++) {
    if (nowdate<=events[e].getStartTime() || updateAll==true) {
      events[e].deleteEvent();
      Utilities.sleep(sleepTime)
    }
    if (e%20==0) {
      Logger.log("Deleted " + e + " events so far...");
    }
  }
  bnames.getRange("N2:N4").setValues([[0],[1],[ 0]]);
  bnames.getRange("H10").setValue("All done, last updated at " + Date());
  CacheService.getDocumentCache().remove("itzRunning");
}
