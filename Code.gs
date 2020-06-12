// NUMERIC CONSTANTS
var MILLIS_PER_MINUTE = 60 * 1000;
var MILLIS_PER_HOUR = 60 * MILLIS_PER_MINUTE;
var MILLIS_PER_DAY = 24 * MILLIS_PER_HOUR;
var MILLIS_PER_WEEK = 7 * MILLIS_PER_DAY;

// NUMBER OF ADDITIONAL WEEKS TO LOOK AHEAD AND MAKE TASKS
var lookAheadWeeks = 2;

// ADD CUSTOM MENU
function onOpen(e) {
  
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Custom Functions")
  .addItem("Update Tasks", 'updateTasks')
  .addToUi();
  
}

function updateTasks() {
  
  // GET TIMEZONE
  var timezone = Session.getScriptTimeZone();
  
  // GET CURRENT TIME
  var currDate = new Date();
  var currDateTime = currDate.getTime();
  var currDateDay = currDate.getDay();
  var currDateHour = currDate.getUTCHours() - currDate.getTimezoneOffset()/60;
  
  // GET LOOK AHEAD DATE
  var lookAheadDate = new Date(new Date().setDate(currDate.getDate() + 7*lookAheadWeeks));
  
  // GET SPREADSHEET CELL REFERENCES
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklySheet = ss.getSheetByName("Weekly Tasks");
  var biweeklySheet = ss.getSheetByName("Bi-weekly Tasks")
  var monthlySheet = ss.getSheetByName("Monthy Tasks");
  
  // GET SPREADSHEET DATA
  var weeklyTaskEntries = weeklySheet.getDataRange().getValues();
  var biweeklyTaskEntries = biweeklySheet.getDataRange().getValues();
  var monthlyTaskEntries = monthlySheet.getDataRange().getValues();
  
  // REMOVE HEADERS FROM SPREADSHEET DATA
  weeklyTaskEntries.splice(0,1);
  biweeklyTaskEntries.splice(0,1);
  monthlyTaskEntries.splice(0,1);
  
  // COMBINE SPREADSHEET DATA
  var taskEntries = weeklyTaskEntries.concat(biweeklyTaskEntries,monthlyTaskEntries);
  
  // GET MOST RECENT SUNDAY (IE TODAY) @ TIME = MIDNIGHT
  var thisSunday = new Date(currDateTime);
  thisSunday.setDate(currDate.getDate() - currDateDay);
  thisSunday.setHours(0,0,0,0);
  
  // GET MOST RECENT MONDAY @ TIME = MIDNIGHT
  var thisMonday = new Date(thisSunday.getTime());
  thisMonday.setDate(thisSunday.getDate() + 1);
  
  // GET MOST RECENT TUESDAY @ TIME = MIDNIGHT
  var thisTuesday = new Date(thisSunday.getTime());
  thisTuesday.setDate(thisSunday.getDate() + 2);
  
  // GET MOST RECENT WEDNESDAY @ TIME = MIDNIGHT
  var thisWednesday = new Date(thisSunday.getTime());
  thisWednesday.setDate(thisSunday.getDate() + 3);
  
  // GET MOST RECENT THURSDAY @ TIME = MIDNIGHT
  var thisThursday = new Date(thisSunday.getTime());
  thisThursday.setDate(thisSunday.getDate() + 4);
  
  // GET MOST RECENT FRIDAY @ TIME = MIDNIGHT
  var thisFriday = new Date(thisSunday.getTime());
  thisFriday.setDate(thisSunday.getDate() + 5);
  
  // GET MOST RECENT SATURDAY @ TIME = MIDNIGHT
  var thisSaturday = new Date(thisSunday.getTime());
  thisSaturday.setDate(thisSunday.getDate() + 6);
  
  // GET LIST OF TASK LISTS
  var taskLists = Tasks.Tasklists.list({maxResults:100});
  
  // GENERATE LIST OF EXISTING TASKS, INCLUDING COMPLETED TASKS
  var oldTasks = [];
  
  for (var j=0;j<taskLists.items.length;j++) {
    
    var currOldTasks = Tasks.Tasks.list(taskLists.items[j].id, {showCompleted:true,showHidden:true,maxResults:100});
    
    var currOldTasksItems = currOldTasks.items;
    
    var currNextPageToken = currOldTasks.nextPageToken;
    
    // WHETHER OR NOT TO GO THROUGH ALL TASKS, NOT JUST PREVIOUS 100
    var checkAllTasks = 1;
    
    while (currNextPageToken && checkAllTasks) {
      
      currOldTasks= Tasks.Tasks.list(taskLists.items[j].id, {showCompleted:true,showHidden:true,maxResults:100,pageToken:currNextPageToken});
      
      currOldTasksItems = currOldTasksItems.concat(currOldTasks.items);
      
      currNextPageToken = currOldTasks.nextPageToken;
      
    }
    
    // SKIP IF TASK LIST IS EMPTY
    if (!currOldTasksItems) break;
    
    for (var k=0;k<currOldTasksItems.length;k++) {
      
      // IF OLD TASK HAS NO DUE DATE, SKIP IT
      if (!currOldTasksItems[k].due) break;
      
      var currOldTaskTitle = currOldTasksItems[k].title;
    
      var currOldTaskDesc = currOldTasksItems[k].notes;
    
      var currOldTaskDateString = fixDate(currOldTasksItems[k].due);
      var currOldTaskDate = new Date(currOldTaskDateString);
    
      oldTasks.push([currOldTaskTitle,currOldTaskDesc,currOldTaskDate]);
      
    }
  }
  
  // ADD APPROPRIATE TAG
  for (var i=0;i<taskEntries.length;i++) {
    
    if (i<weeklyTaskEntries.length) {
      taskEntries[i].unshift("W");
    } else if (i<weeklyTaskEntries.length+biweeklyTaskEntries.length) {
      taskEntries[i].unshift("B");
    } else {
      taskEntries[i].unshift("M");
    }
    
  }
  
  // TASKS THAT MAY NEED TO BE ADDED
  var newTasks = [];
  
  // LOOK THROUGH TASKS ON SPREADSHEET
  for (var i=0;i<taskEntries.length;i++) {
    
    var currTask = taskEntries[i];
    
    var taskType = currTask[0];
    var taskTitle = currTask[1];
    var taskDesc = currTask[2];
    var taskTargetList = currTask[3];
    
    var repMap = null;
    var onWeek = null;
    var startWeek = null;
    var repDay = null;
    var extraTag = null;
    var taskDates = null;
    
    // GET DAYS ON WHICH TASK REPEATS
    if (taskType=="W") {
      
      repMap = currTask.slice(4,11);
      taskDates = currTask.slice(11,currTask.length);
      
    } else if (taskType=="B")  {
      
      repMap = currTask.slice(4,11);
      startWeek = fixYear(new Date(currTask.slice(11,12)));
      taskDates = currTask.slice(12,currTask.length);
      
    } else {
      
      repDay = currTask[4];
      extraTag = currTask[5];
      taskDates = currTask.slice(6,currTask.length);
      
    }
    
    var taskBeginDates = []; var taskEndDates = [];
    
    // CREATE LISTS OF BEGINNING AND END DATES
    for (var j=0;j<taskDates.length/2;j++) {
      
      if (taskDates[2*j] instanceof Date && taskDates[2*j+1] instanceof Date) {
        
        taskBeginDates[j] = fixYear(taskDates[2*j]);
        
        if (taskDates[2*j+1] instanceof Date) {
          taskEndDates[j] = fixYear(taskDates[2*j+1]);
        } else { 
          taskEndDates[j] = lookAheadDate;
        }
      
      } else if (taskDates[2*j].length>0) {
        
        taskBeginDates[j] = fixYear(new Date(taskBeginDates[2*j]));
        
        if (taskDates[2*j+1].length>0) {
          taskEndDates[j] = fixYear(new Date(taskEndDates[2*j+1]));
        } else { 
          taskEndDates[j] = lookAheadDate;
        }
        
      } else break;
      
    }
    
    // FIND CORRESPONDING TASK LIST
    for (var j=0;j<taskLists.items.length;j++) {
    
      if (taskLists.items[j].title == taskTargetList) {
      
        var taskList = taskLists.items[j];
        break;
      
      }
    }
    
    // IF NO TASK LIST FOUND, THROW ERROR
    if (!taskList) throw "No such tasklist";
    
    // LOOP THROUGH lookAheadWeeks NUMBER OF WEEKS
    for (var j=0;j<=lookAheadWeeks;j++) {
      
      var currSunday = new Date(thisSunday.getTime());
      currSunday.setDate(currSunday.getDate() + 7*j);
      
      var currMonday = new Date(thisMonday.getTime());
      currMonday.setDate(currMonday.getDate() + 7*j);
      
      var currTuesday = new Date(thisTuesday.getTime());
      currTuesday.setDate(currTuesday.getDate() + 7*j);
      
      var currWednesday = new Date(thisWednesday.getTime());
      currWednesday.setDate(currWednesday.getDate() + 7*j);
      
      var currThursday = new Date(thisThursday.getTime());
      currThursday.setDate(currThursday.getDate() + 7*j);
      
      var currFriday = new Date(thisFriday.getTime());
      currFriday.setDate(currFriday.getDate() + 7*j);
      
      var currSaturday = new Date(thisSaturday.getTime());
      currSaturday.setDate(currSaturday.getDate() + 7*j);
      
      var currWeek = [currSunday,currMonday,currTuesday,currWednesday,currThursday,currFriday,currSaturday];
      var inRange = [0,0,0,0,0,0,0];
      
      // LOOP THROUGH EVERY DAY IN THIS WEEK
      for (var k=0;k<currWeek.length;k++) {
        
        // IF NO BEGIN/END DATES, ASSUME IN RANGE
        inRange[k] = taskEndDates.length==0;
        
        // LOOP THROUGH ALL BEGIN/END DATE PAIRS TO DETERMINE IF IN RANGE
        for (var m=0;m<taskEndDates.length;m++) {
          
          if (taskEndDates[m]===null) break;
          
          inRange[k] = inRange[k] || dateInRange(currWeek[k],taskBeginDates[m],taskEndDates[m]);
          
        }
        
        if (taskType=="B") {
          
          onWeek = onOffWeek(startWeek,currSunday,timezone);
          
        }
        
        // IF TO BE REPEATED THIS DAY AND IN RANGE, ADD TO NEW LIST
        if (taskType=="W" && repMap[k] && inRange[k]) {
          
          newTasks.push([taskTitle,taskDesc,currWeek[k],taskList]);
          
        } else if (taskType=="B" && onWeek && repMap[k] && inRange[k]) {
          
          newTasks.push([taskTitle,taskDesc,currWeek[k],taskList]);
          
        } else if (taskType=="M" && currWeek[k].getDate()==repDay && inRange[k] && !extraTag) {
          
          newTasks.push ([taskTitle,taskDesc,currWeek[k],taskList]);
          
        } else if (taskType=="M" && currWeek[k].getDate()==repDay && inRange[k] && extraTag.length>0) {
          
          newTasks.push ([taskTitle,taskDesc,nearestBusinessDay(currWeek[k]),taskList]);
          
        }
      }
    }
  }

  // FOR EVERY NEW TASK, DETERMINE IF TASK WITH SAME NAME/DATE EXISTS IN ANY LIST
  for (var i=newTasks.length-1;i>=0;i--) {
    
    for (var j=0;j<oldTasks.length;j++) {
      
      if (sameDay(newTasks[i][2],oldTasks[j][2]) && newTasks[i][0] == oldTasks[j][0]) {
        
        newTasks.splice(i,1); break;
        
      }
    }
  }
  
  Logger.log(newTasks)
  
  // FOR EVERY TASK STILL IN NEW LIST, PARSE AND ADD TO GOOGLE TASKS
  for (var i=0;i<newTasks.length;i++) {
    
    var currNewTask = newTasks[i];

    var taskToAdd = {
      title: currNewTask[0],
      notes: currNewTask[1],
      due: currNewTask[2].toISOString()
    }
    
    // ADD TO GOOGLE TASKS
    Tasks.Tasks.insert(taskToAdd,currNewTask[3].id);
    
  }
}

function dateInRange(entryDate,beginPeriod,endPeriod) {
  
  var entryDateTime = entryDate.getTime();
  var beginPeriodTime = beginPeriod.getTime();
  var endPeriodTime = endPeriod.getTime();
  
  return (entryDateTime>=beginPeriodTime) & (entryDateTime<=endPeriodTime);
  
}

function mod(n, p) {
  
    return n - p * Math.floor(n/p);
  
}

function fixYear(date) {
  
  // IF DATE IS ROUNDED DOWN TO 1900s, CORRECT BY ADDING 100 YEARS
  if (date.getFullYear()<1970) date.setFullYear(date.getFullYear() + 100);
  
  return date;
  
}

function fixDate(ISOString) {
  
  // PARSE ISO STRING TO DATE STRING "MM/DD/YYYY" FOR EASE OF USE
  var indexFirst = ISOString.indexOf('-');
  var indexSecond = ISOString.indexOf('-',indexFirst+1);
  var indexThird = ISOString.indexOf('T',indexSecond+1);
  
  var year = ISOString.slice(0,indexFirst)
  var month = ISOString.slice(indexFirst+1,indexSecond);
  var day = ISOString.slice(indexSecond+1,indexThird);
  
  var fixedDateString = month.concat("/",day,"/",year);
  
  return fixedDateString;

}

function onOffWeek(startDate,currDate,timezone) {
  
  // DETERMINE IF ON WEEK OR OFF WEEK (EVEN OR ODD # OF WEEKS SINCE START WEEK)
  var startDateSunday = new Date(new Date(startDate).setHours(0,0,0,0) - mod(startDate.getDay(),7) * MILLIS_PER_DAY);
  var currDateSunday = new Date(new Date(currDate).setHours(0,0,0,0) - mod(currDate.getDay(),7) * MILLIS_PER_DAY);
  
  // CATCH DAYLIGHT SAVINGS TIME-RELATED ISSUES
  var timezoneDifference = savingsError(startDateSunday,currDateSunday)
  
  var weeksSince = Math.floor((currDateSunday.getTime() - startDateSunday.getTime() + timezoneDifference * MILLIS_PER_HOUR) / MILLIS_PER_WEEK);
  
  var onWeek = mod(weeksSince,2) == 0;
  
  return onWeek;
  
}

function nearestBusinessDay(currDate) {
  
  // IF SATURDAY OR SUNDAY, FIND PREVIOUS FRIDAY
  var currDateTime = currDate.getTime();
  var currDateDay = currDate.getDay();
  
  var businessDate = null;
  
  if (currDateDay==0) {
    
    businessDate = new Date(new Date(currDateTime).setHours(0,0,0,0) - 2 * MILLIS_PER_DAY);
    
  } else if (currDateDay==6) {
    
    businessDate = new Date(new Date(currDateTime).setHours(0,0,0,0) - 1 * MILLIS_PER_DAY);
    
  } else {
    
    businessDate = currDate;
    
  }
  
  return businessDate;
  
}

function sameDay(d1, d2) {
  
  // FIND IF d1 AND d2 ARE THE SAME DAY, MORE RELIABLY THAN getTime() ARITHMETIC
  var sameYear = d1.getFullYear() === d2.getFullYear();
  var sameMonth = d1.getMonth() === d2.getMonth();
  var sameDate = d1.getDate() === d2.getDate();
  
  return sameYear && sameMonth && sameDate
  
}

function savingsError(beginDate,endDate) {
  
  var beginDate = beginDate.getTimezoneOffset();
  var endDate = endDate.getTimezoneOffset();
  
  var timezoneDifference = (beginDate - endDate) / 60;
  
  return timezoneDifference
  
}