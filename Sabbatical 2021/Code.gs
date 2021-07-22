function OnCellChange(event) {
  //return;
  
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (spreadsheet.getActiveSheet().getName() === "Main Table")
    { // only process changes from Main Table sheet
      var range = spreadsheet.getCurrentCell();

      if(range.getNumRows() === 1 
        && range.getNumColumns() === 1 
        && range.getColumn() === 6) 
      {
        var val = range.getValue();
        SetStatusRangeColors(range, val);

        // Move row based on new status
        if (val === "Done")
        { 
          // Move item to bottom of Done group
          MoveRowToGroup('Done Tasks');
        }
        else if (val === "Doing")
        {
          MoveRowToGroup('Doing');
        }
      }
    }
  } catch(err) 
  {
    Browser.msgBox(err);
  }
};

function getFirstEmptyGroupRow(groupName) {
  var spr = SpreadsheetApp.getActiveSpreadsheet();

  // Find group header
  var groupHeaderRowIndex = findValueInColumn('B', groupName);

  // Get values in B column and loop through them, starting at the group header
  var column = spr.getRange('B:B');
  var values = column.getValues(); // get all data in one call
  var ct = groupHeaderRowIndex;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);
}

function getFirstEmptyRowByColumnArray() {
  var spr = SpreadsheetApp.getActiveSpreadsheet();
  var column = spr.getRange('B:B');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);
}

function findValueInColumn(columnLetter, searchValue){
  //let term = searchValue;
  let data = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(columnLetter + ':' + columnLetter).getValues();
  let row = data.findIndex(users => {return users[0] == searchValue});  
  //Browser.msgBox(row+1);
  return row+1;
}

function MoveRowToGroup(targetGroupName){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var rowIndex = spreadsheet.getActiveCell().getRowIndex();
  var activeRowRangeA1String = rowIndex.toString() + ":" + rowIndex.toString();

  var nextEmptyRowIndex = getFirstEmptyGroupRow(targetGroupName);
  if (nextEmptyRowIndex > 0)
    spreadsheet.getActiveSheet().moveRows(spreadsheet.getRange(activeRowRangeA1String), nextEmptyRowIndex);
  else
    Browser.msgBox('Could not find the next empty row in the ' + targetGroupName + ' group');
}

function CopyDailyRecurringTasksToMainTable()
{
  CopyRecurringTasksToMainTable("Daily");
}

function CopyWeeklyRecurringTasksToMainTable()
{
  CopyRecurringTasksToMainTable("Weekly");
}

// Copy all the recurring tasks of a task type to the Main Table if the task has not been copied yet
function CopyRecurringTasksToMainTable(recurringTaskType)
{
  var timeZone = "GMT-8";
  // Set up constants based on recurring task type
  var recurringTaskSheetName = "unknown";
  var nbrTaskColumns = -1;
  var targetGroupName = "";
  var hoursColumnIndex = -1;
  var weekDayColumnIndex = -1;
  var dueDate = new Date();
  var priorityColumnIndex = -1;
  if (recurringTaskType==="Daily")
  {
    recurringTaskSheetName = "Daily Recurring Tasks";
    nbrTaskColumns = 6;
    targetGroupName = "Recurring Daily";
    hoursColumnIndex = 4;
    priorityColumnIndex = 3;
    dueDate = new Date();
  }
  else if ((recurringTaskType==="Weekly"))
  {
    recurringTaskSheetName = "Weekly Recurring Tasks";
    nbrTaskColumns = 7;
    targetGroupName = "Recurring Weekly";
    hoursColumnIndex = 5;
    weekDayColumnIndex = 3;
    priorityColumnIndex = 4;
    dueDate.setDate(dueDate.getDate()+6);
  }
  else
    Browser.msgBox('Invalid recurring task type ' + recurringTaskType);

  try
  {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Switch to Main Table sheet
    ss.setActiveSheet(ss.getSheetByName("Main Table"));

    // Select all Daily Recurring Tasks
    var recurringTasksSheet = ss.getSheetByName(recurringTaskSheetName);
    var recurringTasks = recurringTasksSheet.getRange(2, 1, recurringTasksSheet.getLastRow() - 1, nbrTaskColumns);

    // Copy Recurring Tasks to Main Table
    var nextEmptyRowIndex = getFirstEmptyGroupRow(targetGroupName);
    if (nextEmptyRowIndex > 0)
    {
      // loop through rows and copy one by one
      var values = recurringTasks.getValues();
      var sourceRowIndex = 2;
      values.forEach(function(row)
      {
        // Check if this is the first time on the current day and past the start time of the task
        var lastDate = new Date(row[nbrTaskColumns-1]);
        var currentDate = new Date();
        var currentDateOnly = new Date().setHours(0,0,0,0);
        if (isValidDate(lastDate) == false)
          lastDate = new Date("01/01/21");

        // Day Last Copied < Current Day and Current Hour > Recurring Task Hour
        // for weekly tasks (weekdayColumnIndex != 0) also check the day of the week
        if ((currentDateOnly - lastDate.setHours(0,0,0,0) > 0) && (currentDate.getHours() >= Number(row[hoursColumnIndex]))
            && (weekDayColumnIndex===-1 || row[weekDayColumnIndex] === Utilities.formatDate(currentDate, timeZone, "EEEE")))
        { // last time task was inserted was before today

          // create empty row in Recurring Daily group
          ss.getActiveSheet().insertRows(nextEmptyRowIndex, 1);

          // Add date after recurring task name and add recurring task to Main Tasks sheet
          ss.getActiveSheet().getRange(nextEmptyRowIndex, 2).setValue(row[0] + ' (' + Utilities.formatDate(new Date(), timeZone, "MM/dd/yy") + ')');
          
          // Set Status field
          var statusRange = ss.getActiveSheet().getRange(nextEmptyRowIndex, 6);
          statusRange.setValue('To Do');
          SetStatusRangeColors(statusRange, 'To Do');
        
          // Set Due Date
          ss.getActiveSheet().getRange(nextEmptyRowIndex, 7).setValue(Utilities.formatDate(dueDate, timeZone, "MM/dd/yyyy"));
          
          // Set Priority
          var priorityRange = ss.getActiveSheet().getRange(nextEmptyRowIndex, 8);
          priorityRange.setValue(row[priorityColumnIndex]);
          SetPriorityRangeColors(priorityRange, row[priorityColumnIndex]);

          // update recurring task LastCopied value
          recurringTasksSheet.getRange(sourceRowIndex, nbrTaskColumns).setValue(Utilities.formatDate(new Date(), timeZone, "MM/dd/yy HH:mm"));

          Logger.log('Inserted row: ' + row);

          nextEmptyRowIndex++;
        }
        sourceRowIndex++;
      
      });
    }
    else
      Browser.msgBox('Could not find the next empty row in the ' + targetGroupName + ' group');

    //Browser.msgBox('Copy Daily Recurring Tasks to Main Table');
  } catch(err) 
  {
    Browser.msgBox(err);
  }
}

// Check is date value is valod
function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

// Sets the font color and background color of a status cell based on the new status
function SetStatusRangeColors(range, newStatus)
{
  if (newStatus === "Done")
  { 
    // Change background and font color
    range.setBackground('#00c875');
    range.setFontColor('white');
  }
  else if (newStatus === "To Do")
  {
    range.setBackground('#cfe2f3');
    range.setFontColor('white');
  }
  else if (newStatus === "Doing")
  {
    range.setBackground('#fdab3d');
    range.setFontColor('white');
  }
  else if (newStatus === "Waiting")
  {
    range.setBackground('#a25ddc');
    range.setFontColor('white');
  }
  else if (newStatus === "On Hold")
  {
    range.setBackground('#e2445c');
    range.setFontColor('white');
  }
  else
    range.setBackground('#c4c4c4');
}

// Sets the font color and background color of a status cell based on the new status
function SetPriorityRangeColors(range, newPriority)
{
  if (newPriority === "High")
  { 
    // Change background and font color
    range.setBackground('#ff158a');
    range.setFontColor('white');
  }
  else if (newPriority === "Medium")
  {
    range.setBackground('#fdab3d');
    range.setFontColor('white');
  }
  else if (newPriority === "Low")
  {
    range.setBackground('#9cd326');
    range.setFontColor('white');
  }
  else
    range.setBackground('#c4c4c4');
}

// Trigger that is executed every minute
function OnMinuteTimer()
{
  CopyDailyRecurringTasksToMainTable();
  CopyWeeklyRecurringTasksToMainTable();
}

function TestDaily()
{
    CopyDailyRecurringTasksToMainTable();
 //   CopyWeeklyRecurringTasksToMainTable();
}

function TestWeekly()
{
    CopyWeeklyRecurringTasksToMainTable();
}

function TestDebug(){
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var color = ss.getActiveSheet().getRange('H9').getBackground();

  var dueDate = new Date("07/29/21"); //.setDate(new Date().getDate()+6);
  var testDate = new Date();
  var a1date = dueDate.getDate();
  dueDate.setDate(a1date+6);
  //value.setMonth((a1date+1) % 12);

}

function getUserTimeZone() {
  Logger.log("Script Time Zone: " + Session.getScriptTimeZone());
}
