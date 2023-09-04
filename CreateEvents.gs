function CreateEvents() {
  /*  The CreateEvents function in Calendar Scheduler takes input from
      the AHS Personal Scheduler Google Spreadsheet and then populates
      a Google Calendar with appointments for each class on each school day.

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
  Logger.log("Will be pulling data from Spreadsheet " + ss)

  // Get class times for full day, half day, and delayed (1,1.5,2 hrs) days
  // These data are set in the spreadsheet and can be updated there
  var scheds = ss.getSheets()[5];
  var range = scheds.getRange("J2:J6");
  const fullStartTimes = range.getValues();// ["8:15", "9:23", "10:47", "11:49", "13:49"];
  var range = scheds.getRange("K2:K6");
  const fullEndTimes =  range.getValues();// ["9:17", "10:41", "11:47", "13:43", "14:51"];

  var range = scheds.getRange("AA2:AA6");
  const halfStartTimes = range.getValues();// ["8:15", "9:05","9:52" "10:41"];
  var range = scheds.getRange("AB2:AB6"); //
  const halfEndTimes = range.getValues();// ["9:18", "10:23", "11:30"];

  var range = scheds.getRange("J16:R19"); //
  const LunchClassTimes = range.getValues();// ["9:18", "10:23", "11:30"];

  // The rotating cycle of class meetings over 8 days
  var range = scheds.getRange("B8:I8");
  const dayTypeBlocks =range.getValues()[0];//  ["ACHEG",	"BDFGE",	"AHDCF",	"BAHGE",	"CBFDG",	"AHEFC",	"BADEG",	"CBHFD"];
  var range = scheds.getRange("B9:I9");
  const hTypeBlocks = range.getValues()[0];

  var range = scheds.getRange("T8:Z8");
  const halfDayTypeBlocks = range.getValues()[0];// ACHEG	BADEG	AFCHH	BDFGE	BAHGE	CBHFD	AHHDF


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

  var range = bnames.getRange(2,14,3); // N2:N4
  var loopvals = range.getValues();
  termStart = loopvals[0];
  iStart = loopvals[1];
  blockStart = loopvals[2];

  Logger.log("termstart = " + termStart + "istart = " + iStart + ", blockStart = " + blockStart)
  bnames.getRange("H10").setValue("You'll need to run the script again; still working.");

  var range = bnames.getRange("H3");
  var CalName = range.getValue();
  if (CalName=="") {
    var calendars = CalendarApp.createCalendar('MyClasses');
    bnames.getRange("H3").setValue("MyClasses");
    return
  }
  Logger.log("Will be adding events to Calendar: " + CalName);

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

    if (allDayDesc==1 && (nowdate<=startTime || updateAll==true)) {
      var event = calendars.createAllDayEvent(dayDesc, new Date(values[i][0]), 
        {description: "This event was created by the webapp AHS Personal Calendar." +
        " To delete all these events or update in bulk, use the buttons in your Google Sheet entitled" +
        " 'AHS Personal Calendar', available at: " + ss.getUrl() + ""});
    }

    if (["X","Y","Z","EXAM"].indexOf(dayType) > -1) {
      var startTime = new Date(values[i][0].getFullYear() + "-" + (values[i][0].getMonth()+1) + "-" 
        + (values[i][0].getDate()) + " 8:15"); 
      var endTime = new Date(values[i][0].getFullYear() + "-" + (values[i][0].getMonth()+1) + "-"
        + (values[i][0].getDate()) + " 14:51");
      caltitle = dayType;

      Logger.log("Different day startTime=" + startTime + ", endTime=" + endTime + ", caltitle="
        + caltitle + ", dayDesc=" + dayDesc);

      if (nowdate<=startTime || updateAll==true) {

        event = calendars.createEvent("Day " + caltitle, startTime, endTime,
          {description: dayDesc + "\n\nThis event was created by the webapp AHS Personal Calendar." +
          " To delete all these events or update in bulk, use the buttons in your Google Sheet entitled" +
          " 'AHS Personal Calendar', available at: " + ss.getUrl() + ""});
      }

      continue;
    } else if (values[i][1]=="Full Day") {
      var todayBlocks = dayTypeBlocks[dayType-1];
      var startTimes = fullStartTimes;
      var endTimes = fullEndTimes;
    } else if (values[i][1]=="Half Day") {
      var todayBlocks = halfDayTypeBlocks[dayType[dayType.length - 1]-1];
      var startTimes = halfStartTimes;
      var endTimes = halfEndTimes;
    }

    // Block #
    for (var block = blockStart; block < todayBlocks.length; block++) {

      var mystart = new Date(startTimes[block]);
      var myend = new Date(endTimes[block]);

      Logger.log("what have we here: " + mystart.getHours() + ":" + mystart.getMinutes() + ":00 - " +myend.getHours() + ":"
        + myend.getMinutes() + ":00")

        // class ends at ... LunchClassTimes[lunch][0]
        // class starts again at ... LunchClassTimes[lunch][1]

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

      if (blockLetter=="H") 
      { // Get the H block number
        var Hnum = hTypeBlocks[dayType-1]
        if (Hnum == undefined)
        {
          Hnum = "";
        }
        Logger.log("Today's H Block is " + Hnum)
        if(blocknames[ind][0][0]=="*") // some staff teach a few HBlocks
        {
          if (blocknames[ind][0].includes(Hnum.slice(1))==false) {
            continue
          }
        }
        var caltitle = Hnum + " Block (" + blockLetter + ")"; // Look up name on 1st tab given block 
      } else {
        var caltitle = blocknames[ind][0] + " (" + blockLetter + ")"; // Look up name on 1st tab given block letter
      }

        var calloc =  "Room #" + blockrooms[ind]; // Look up room on first tab given block letter
        var mydesc = "Session #" + values[i][3+ind]

        // Not implemented yet, for duration:
        // + " for this class.\n\nMinutes of class before today: " + durations[ind] + "\nMinutes of class today: "
        // + durToday + "\nMinutes of class after today: " + durations[ind]+durToday
        + "\n\nThis event was created by the webapp AHS Personal Calendar."+
        " To delete all these events or update in bulk, use the buttons in your Google Sheet entitled" +
        " 'AHS Personal Calendar', available at: " + ss.getUrl() + "";
        Logger.log("startTime=" + startTime + ", endTime=" + endTime + ", caltitle=" + caltitle
          + ", calloc=" + calloc)
        durations[ind] += durToday;
        if (block==3 && values[i][1]!="Half Day") {// reg full day, 0-based block 4
          var lunch = lunches[ind];
          if (lunch=="") {
            Logger.log("No lunch specified, making Block 3 the full time");

          } else {
            Logger.log("Todays lunch is #" + lunch);
            mydesc = "Lunch #" + lunch + "\n" + mydesc;
            if (lunch==1) {
            // updateStartTime

              var mystart = new Date(LunchClassTimes[lunch-1][1]);
              var startTime = new Date(values[i][0].getFullYear() + "-" + (values[i][0].getMonth()+1) + "-"
              + (values[i][0].getDate()) + " " + mystart.getHours() + ":" + mystart.getMinutes() + ":00"); 
            } else if (lunch==4) {
              // class ends at ... LunchClassTimes[lunch-1][0]

              var myend = new Date(LunchClassTimes[lunch-1][0]);
              var endTime = new Date(values[i][0].getFullYear() + "-" + (values[i][0].getMonth()+1) + "-" 
              + (values[i][0].getDate()) + " " + myend.getHours() + ":" + myend.getMinutes() + ":00"); 
            } else {
              Logger.log("lunch = "+lunch+", tmes="+LunchClassTimes)
              var myend = new Date(LunchClassTimes[lunch-1][0]);
              var NewendTime = new Date(values[i][0].getFullYear() + "-" + (values[i][0].getMonth()+1) + "-" 
              + (values[i][0].getDate()) + " " + myend.getHours() + ":" + myend.getMinutes() + ":00"); 
              var event = 
                calendars.createEvent(caltitle, startTime, NewendTime, {location: calloc, description: mydesc});

              // class starts again at ...  updateStartTime
              var mystart = new Date(LunchClassTimes[lunch-1][1]);
              var startTime = new Date(values[i][0].getFullYear() + "-" + (values[i][0].getMonth()+1) + "-"
              + (values[i][0].getDate()) + " " + mystart.getHours() + ":" + mystart.getMinutes() + ":00"); 
            }
          }
        }
        if (startTime >= endTime)
          Logger.log(startTime + " - " + endTime)
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
