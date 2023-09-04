// The half day schedule was updated by the school the last
// week of September. The logic is now updated for future users.
//
// For people who used this calendar prior to October 1, 2021,
// use this function to update your half-day schedule.

function FixHalfDays() {
  /*  The FixHalfDays function in Calendar Scheduler takes input from
      the AHS Personal Scheduler Google Spreadsheet and then updates the
      Google Calendar appointments for half days.

      Author: marianne.bezaire@andoverma.us
      Date Updated: September, 2023

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

    Logger.log("Will be deleting half-day events from Calendar: " + CalName);

    // "This event was created by the webapp AHS Personal Calendar."

    var events = CalendarApp.getCalendarsByName(CalName)[0].getEvents(startYear, endYear,
      {search: "Half Day, ER "});
    Logger.log("Events = " + events.length);
    bnames.getRange("H10").setValue("You'll need to run the script again; still working.");

    for (e=0; e<events.length; e++) {
        var half_events = CalendarApp.getCalendarsByName(CalName)[0].getEvents(events[e].getStartTime(), events[e].getEndTime(),
      {search: "This event was created by the webapp AHS Personal Calendar."});
    Logger.log("Half day - single day's Events = " + half_events.length);

      for (g=0; g<half_events.length; g++) {
        if ((nowdate<=half_events[g].getStartTime() || updateAll==true) && half_events[g].isAllDayEvent()==false) {
          half_events[g].deleteEvent();
          Utilities.sleep(sleepTime)
        }
        if (g%20==0 && g>0) {
          Logger.log("Deleted " + g + " half-day events so far...");
        }
      }
    }
    bnames.getRange("N2:N4").setValues([[0],[1],[ 0]]);
    bnames.getRange("H10").setValue("Deleted old half-days at " + Date() + ", now adding updated ones...");

    var scheds = ss.getSheets()[5];
    var range = scheds.getRange("AA2:AA5");
    const halfStartTimes = range.getValues();// ["8:15", "9:22", "10:27"];
    var range = scheds.getRange("AB2:AB5"); //
    const halfEndTimes = range.getValues();// ["9:18", "10:23", "11:30"];

    var range = scheds.getRange("T8:Z8");
    const halfDayTypeBlocks = range.getValues()[0];// ["ABC","DEF","GAB","CDE","FGA","BCD","EFG"];

    // read in i and block - if the script times out in the middle of nested for loops, we want
    // to be able to continue where we left off. Where we left off is saved in a range of cells
    // in the sheet:
    var bnames = ss.getSheets()[0];

    // Initialize the counters for the nested for loops
    var termStart = 0; //  Semester
    var iStart = 1; //  Day in Semester
    var blockStart = 0; // Block in Day

    // Find whether user wants to work with events in Fall (1), Spring (2), or all year (9/All)
    var range = bnames.getRange("H5");
    if (range.getValue()=="All") {
      var Term = 9; // Fall and Spring Semesters
    } else {
      var Term = range.getValue(); // One Semester only
    }

    var range = bnames.getRange("N5");
    var sleepTime = range.getValue();

    Logger.log("termstart = " + termStart + "istart = " + iStart + ", blockStart = " + blockStart)
    bnames.getRange("H10").setValue("You'll need to run the script again; still working.");

    var range = bnames.getRange("H3");
    var CalName = range.getValue();
    if (CalName=="") {
      var calendars = CalendarApp.createCalendar('MyClasses');
      bnames.getRange("H3").setValue("MyClasses");
      return
    }
    Logger.log("Will be adding updated half_day events to Calendar: " + CalName);

    var range = bnames.getRange("H2");
    var allDayDesc = range.getValue();

    var range = bnames.getRange("H4");
    var updateAll = (range.getValue()=='All');

    if (Term==1) {var Sheets2Do=[3];}
    else if (Term==2) {var Sheets2Do=[4]}
    else if (Term==9) {var Sheets2Do=[3,4];}

    for (term2do=termStart;term2do<Sheets2Do.length;term2do++) {
    var sheet = ss.getSheets()[Sheets2Do[term2do]];
    var range = sheet.getDataRange();
    var values = range.getValues();

    
    var range = bnames.getRange(2 + (Sheets2Do[term2do]-3)*9, 2, 8);
    var blocknames = range.getValues();
    var range = bnames.getRange(2 + (Sheets2Do[term2do]-3)*9, 3, 8);
    var blockletters = range.getValues();
    var range = bnames.getRange(2 + (Sheets2Do[term2do]-3)*9, 5, 8);
    var blockrooms = range.getValues();
    var range = bnames.getRange(2 + (Sheets2Do[term2do]-3)*9, 6, 8);
    var lunches = range.getValues();
    Logger.log("Will be adding events from Term: " + sheet.getName());
    var blankcheck=0
    for (q=0;q<blocknames.length;q++) {
      if (blocknames[q]=="") {
        blankcheck += 1
      }
    }

    for (q=0;q<blockletters.length;q++) {
      if (blockletters[q]=="") {
        blankcheck += 1
      }
    }

    for (q=0;q<blockrooms.length;q++) {
      if (blockrooms[q]=="") {
        blankcheck += 1
      }
    }

    for (q=0;q<lunches.length;q++) {
      if (lunches[q]=="") {
        blankcheck += 1
      }
    }

    //Only run this check for students
    var range = bnames.getRange("H6");
    var PersonType = range.getValue();
    if (blankcheck>0 && PersonType=="Student") {
      bnames.getRange("H10").setValue("Make sure all the classes, letters, rooms, and lunches"
      + " are filled out for the semester you want to schedule");
      return
    }


    var durations = [0,0,0,0,0,0,0,0]; // Not implemented yet
    for (var i = iStart; i < values.length; i++) {  // start at 1 due to header in 0th row
      var dayType = values[i][2];
      var dayDesc = values[i][1] + ", " + values[i][2];
      Logger.log("Day Description: " + dayDesc);

      var calendars = CalendarApp.getCalendarsByName(CalName)[0]; // getDefaultCalendar();

      if (calendars==null)
      {
        var calendars = CalendarApp.createCalendar(CalName);
        bnames.getRange("H10").setValue("Just created Calendar " + CalName);
      }

      var nowdate = new Date();

      if (values[i][1]=="Half Day") {
        var todayBlocks = halfDayTypeBlocks[dayType[dayType.length - 1]-1];
        var startTimes = halfStartTimes;
        var endTimes = halfEndTimes;
      }
      else
      {
        continue;
      }

      // Block #
      for (var block = blockStart; block < todayBlocks.length; block++) {

        var mystart = new Date(startTimes[block]);
        var myend = new Date(endTimes[block]);

        Logger.log(mystart.getHours() + ":" + mystart.getMinutes() + ":00 - " +myend.getHours() + ":"
          + myend.getMinutes() + ":00")

        var startTime = new Date(values[i][0].getFullYear() + "-" + (values[i][0].getMonth()+1)
          + "-" + (values[i][0].getDate()) + " " + mystart.getHours() + ":" + mystart.getMinutes() + ":00"); 

        var endTime = new Date(values[i][0].getFullYear() + "-" + (values[i][0].getMonth()+1)
          + "-" + (values[i][0].getDate()) + " " + myend.getHours() + ":" + myend.getMinutes() + ":00");  

        if (nowdate>startTime && updateAll==false) {
          continue;
        }

        var durToday = endTime.getHours()*60 + endTime.getMinutes() - startTime.getHours()*60 
          - startTime.getMinutes();
        var blockLetter = todayBlocks[block];
        Logger.log("todayBlocks = " + todayBlocks + ", block = " + block + ", blockLetter = "
          + blockLetter + ", blockletters = " + blockletters);
        var ind = blockletters.join('').indexOf(blockLetter);

        if (blocknames[ind][0]=="") {
          Logger.log("empty blocknames for ind = " + ind + ", move on from i=" + i)
          continue
        } // if no class specified (can happen for Staff) then skip making a calendar entry


        var caltitle = blocknames[ind][0] + " (" + blockLetter + ")"; // Look up name on 1st tab given block letter
        var calloc =  "Room #" + blockrooms[ind]; // Look up room on first tab given block letter
        var mydesc = "Session #" + values[i][3+ind]

        mydesc += "\n\nThis event was created by the webapp AHS Personal Calendar." +
          " To delete all these events or update in bulk, use the buttons in your Google Sheet entitled" +
          " 'AHS Personal Calendar', available at: " + ss.getUrl() + "";

        Logger.log("startTime=" + startTime + ", endTime=" + endTime + ", caltitle=" + caltitle
          + ", calloc=" + calloc)
        durations[ind] += durToday;

        var event = calendars.createEvent(caltitle, startTime, endTime, {location: calloc, description: mydesc});
    
        // write out i and block
        bnames.getRange("N2:N4").setValues([[term2do],[i],[ block+1]]);
        Utilities.sleep(sleepTime);
      } // added all the blocks for a day
      blockStart = 0; // if the script had to restart where it left off for one day, blockStart>0, but of course
      // the next day we want to start again at 0 to add all blocks.
      bnames.getRange("H10").setValue("You'll need to run the script again; still working.");
    } // added all the days for this term
    
    //SpreadsheetApp.getUi().prompt('Your class appointments in Google Calendar have been added/updated.');;
    bnames.getRange("N2:N4").setValues([[0],[1],[ 0]]);
    bnames.getRange("H10").setValue("All done, last updated at " + Date());
  } // finished all the terms for this function

  iStart = 1;
  CacheService.getDocumentCache().remove("itzRunning");
}// finished the fcn
