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
  var yearlySheet = ss.getSheetByName("Yearly Tasks");
  var nthDayWeekMonthSheet = ss.getSheetByName("Nth Day of Week of Month");
  
  // GET SPREADSHEET DATA
  var weeklyTaskEntries = weeklySheet.getDataRange().getValues();
  var biweeklyTaskEntries = biweeklySheet.getDataRange().getValues();
  var monthlyTaskEntries = monthlySheet.getDataRange().getValues();
  var yearlyTaskEntries = yearlySheet.getDataRange().getValues();
  var nthDayWeekMonthEntries = nthDayWeekMonthSheet.getDataRange().getValues();
  
  // REMOVE HEADERS FROM SPREADSHEET DATA
  weeklyTaskEntries.splice(0,1);
  biweeklyTaskEntries.splice(0,1);
  monthlyTaskEntries.splice(0,1);
  yearlyTaskEntries.splice(0,1);
  nthDayWeekMonthEntries.splice(0,1);
  
  // COMBINE SPREADSHEET DATA
  var taskEntries = weeklyTaskEntries.concat(biweeklyTaskEntries,monthlyTaskEntries,yearlyTaskEntries,nthDayWeekMonthEntries);
  
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
    } else if (i<weeklyTaskEntries.length+biweeklyTaskEntries.length+monthlyTaskEntries.length) {
      taskEntries[i].unshift("M");
    } else if (i<weeklyTaskEntries.length+biweeklyTaskEntries.length+monthlyTaskEntries.length+yearlyTaskEntries.length) {
      taskEntries[i].unshift("Y");
    } else {
      taskEntries[i].unshift("N");
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
    
    var repMap, onWeek, startWeek, repDay, repMonth, extraTag, taskDates;
    var repWeekDayInstance, repWeekDay, afterFirstWeekDay;
    
    // FOR EACH TASK TYPE, COLLECT APPROPRIATE INFORMATION
    if (taskType=="W") {
      
      repMap = currTask.slice(4,11);
      taskDates = currTask.slice(11,currTask.length);
      
    } else if (taskType=="B")  {
      
      repMap = currTask.slice(4,11);
      startWeek = fixYear(new Date(currTask.slice(11,12)));
      taskDates = currTask.slice(12,currTask.length);
      
    } else if (taskType=="M") {
      
      repDay = currTask[4];
      extraTag = currTask[5];
      taskDates = currTask.slice(6,currTask.length);
      
    } else if (taskType=="Y") {
      
      repMonth = monthStringToNum(currTask[4]);
      repDay = currTask[5];
      taskDates = currTask.slice(6,currTask.length);
      
    } else if (taskType=="N") {
      
      repWeekDayInstance = currTask[4];
      repWeekDay = dayStringToNum(currTask[5]);
      repMonth = monthStringToNum(currTask[6]);
      afterFirstWeekDay = dayStringToNum(currTask[7]);
      taskDates = currTask.slice(8,currTask.length);
      
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
        
        // IF TO BE REPEATED THIS DAY AND IN RANGE, ADD TO NEW LIST
        if (taskType=="W" && repMap[k] && inRange[k]) {
          
          newTasks.push([taskTitle,taskDesc,currWeek[k],taskList]);
          
        } else if (taskType=="B" && onOffWeek(startWeek,currSunday,timezone) && repMap[k] && inRange[k]) {
          
          newTasks.push([taskTitle,taskDesc,currWeek[k],taskList]);
          
        } else if (taskType=="M" && currWeek[k].getDate()==repDay && inRange[k] && !extraTag) {
          
          newTasks.push ([taskTitle,taskDesc,currWeek[k],taskList]);
          
        } else if (taskType=="M" && currWeek[k].getDate()==repDay && inRange[k] && extraTag.length>0) {
          
          newTasks.push ([taskTitle,taskDesc,nearestBusinessDay(currWeek[k]),taskList]);
          
        } else if (taskType=="Y" && currWeek[k].getMonth()==repMonth && currWeek[k].getDate()==repDay && inRange[k]) {
          
          newTasks.push([taskTitle,taskDesc,currWeek[k],taskList]);
          
        } else if (taskType=="N" && (currWeek[k].getMonth()==repMonth || repMonth==12) && currWeek[k].getDay()==repWeekDay
        && getNumOfWeekDay(currWeek[k], afterFirstWeekDay)==repWeekDayInstance && inRange[k]) {
          
          newTasks.push([taskTitle,taskDesc,currWeek[k],taskList])
          
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

// CHECK IF entryDate IS BETWEEN beginPeriod AND endPeriod, INCLUSIVE
function dateInRange(entryDate,beginPeriod,endPeriod) {
  
  var entryDateTime = entryDate.getTime();
  var beginPeriodTime = beginPeriod.getTime();
  var endPeriodTime = endPeriod.getTime();
  
  return (entryDateTime>=beginPeriodTime) & (entryDateTime<=endPeriodTime);
  
}

// IF DATE IS ROUNDED DOWN TO 1900s, CORRECT BY ADDING 100 YEARS
function fixYear(date) {
  
  if (date.getFullYear()<1970) date.setFullYear(date.getFullYear() + 100);
  
  return date;
  
}

// PARSE ISO STRING TO DATE STRING "MM/DD/YYYY" FOR EASE OF USE
function fixDate(ISOString) {
  
  var indexFirst = ISOString.indexOf('-');
  var indexSecond = ISOString.indexOf('-',indexFirst+1);
  var indexThird = ISOString.indexOf('T',indexSecond+1);
  
  var year = ISOString.slice(0,indexFirst)
  var month = ISOString.slice(indexFirst+1,indexSecond);
  var day = ISOString.slice(indexSecond+1,indexThird);
  
  var fixedDateString = month.concat("/",day,"/",year);
  
  return fixedDateString;

}

// DETERMINE IF ON WEEK OR OFF WEEK (EVEN OR ODD # OF WEEKS SINCE START WEEK)
function onOffWeek(startDate,currDate) {
  
  var currDate = new Date(currDate.getTime());
  
  var startDateSunday = new Date(startDate.getTime());
  startDateSunday.setHours(0,0,0,0);
  startDateSunday.setDate(startDate.getDate()-mod(startDate.getDay(),7));
  
  var currDateSunday = new Date(currDate.getTime());
  currDateSunday.setHours(0,0,0,0);
  currDateSunday.setDate(currDate.getDate()-mod(currDate.getDay(),7));
  
  if (startDateSunday.getTime()>currDateSunday.getTime()) return -1;
  
  var weeksSince = 0;
  
  while (!sameDay(currDateSunday,startDateSunday)) {
    
    weeksSince+=1;
    
    currDateSunday.setDate(currDateSunday.getDate()-7);
    
  }
  
  return mod(weeksSince,2)==0;
  
}

// IF SATURDAY OR SUNDAY, FIND PREVIOUS FRIDAY
function nearestBusinessDay(currDate) {
  
  var currDateTime = currDate.getTime();
  var currDateDay = currDate.getDay();
  
  var businessDate;
  
  if (currDateDay==0 || currDateDay==6) {
    
    businessDate = new Date(currDate.getTime());
    businessDate.setDate(currDate.getDate()-mod(currDate.getDay()+2,7));
    
  }
  
  return businessDate;
  
}

// CONVERT NAME OF MONTH TO VALUE BETWEEN 0 & 11, OTHERWISE 12 IF 'ALL', OTHERWISE -1
function monthStringToNum(monthString) {
  
  if (monthString=="January") {
    return 0;
  } else if (monthString=="February") {
    return 1;
  } else if (monthString=="March") {
    return 2;
  } else if (monthString=="April") {
    return 3;
  } else if (monthString=="May") {
    return 4;
  } else if (monthString=="June") {
    return 5;
  } else if (monthString=="July") {
    return 6;
  } else if (monthString=="August") {
    return 7;
  } else if (monthString=="September") {
    return 8;
  } else if (monthString=="October") {
    return 9;
  } else if (monthString=="November") {
    return 10;
  } else if (monthString=="December") {
    return 11;
  } else if (monthString=="All") {
    return 12;
  } else {
    return -1;
  }
  
}

// CONVERT NAME OF DAY TO VALUE BETWEEN 0 & 6, OTHERWISE -1
function dayStringToNum(dayString) {
  
  if (dayString=="Sunday") {
    return 0;
  } else if (dayString=="Monday") {
    return 1;
  } else if (dayString=="Tuesday") {
    return 2;
  } else if (dayString=="Wednesday") {
    return 3;
  } else if (dayString=="Thursday") {
    return 4;
  } else if (dayString=="Friday") {
    return 5;
  } else if (dayString=="Saturday") {
    return 6;
  } else {
    return -1;
  }
  
}

// FIND WEEK DAY NUMBER OF DAY IN MONTH, ADJUSTING FOR DECLARED OFFSET
function getNumOfWeekDay(currDate,afterFirstWeekDay) {
  
  var currDateDate = currDate.getDate();
  
  var firstDateOfWeekDay = getFirstDateOfWeekDay(currDate,afterFirstWeekDay)+1;
  
  var adjustedCurrDateDate = currDateDate-firstDateOfWeekDay;
  
  if (currDateDate-firstDateOfWeekDay <= 7) {
    return 1;
  } else if (currDateDate-firstDateOfWeekDay <= 14) {
    return 2;
  } else if (currDateDate-firstDateOfWeekDay <= 21) {
    return 3;
  } else if (currDateDate-firstDateOfWeekDay <= 28) {
    return 4;
  } else if (currDateDate-firstDateOfWeekDay <= 31) {
    return 5;
  } else {
    return -1;
  }
  
}

// FIND DATE OF FIRST OCCURENCE OF WEEK DAY IN MONTH
function getFirstDateOfWeekDay(currDate,weekDay) {
  
  if (weekDay==-1) return -1;
  
   currDate = new Date(currDate.getTime())
   currDate.setDate(1);
  
  while (currDate.getDay()!=weekDay) {
    currDate.setDate(currDate.getDate()+1);
  }
  
  return currDate.getDate();
  
}

// FIND IF d1 AND d2 ARE THE SAME DAY
function sameDay(d1,d2) {
  
  // MORE RELIABLE THAN getTime() ARITHMETIC
  var sameYear = d1.getFullYear() === d2.getFullYear();
  var sameMonth = d1.getMonth() === d2.getMonth();
  var sameDate = d1.getDate() === d2.getDate();
  
  return sameYear && sameMonth && sameDate
  
}

// ADJUST FOR DAYLIGHT SAVINGS ERROR -- NEEDS REVISION
function savingsError(beginDate,endDate) {
  
  var beginDate = beginDate.getTimezoneOffset();
  var endDate = endDate.getTimezoneOffset();
  
  var timezoneDifference = (beginDate - endDate) / 60;
  
  return timezoneDifference
  
}

// MOD p OF n
function mod(n, p) {
  
    return n - p * Math.floor(n/p);
  
}