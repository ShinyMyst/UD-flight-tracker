function main() {
  // Prepare Tabs
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var tasksTabName = 'Tasks'; 
  var emailTabName = 'Control'; 
  var taskSheet = spreadsheet.getSheetByName(tasksTabName);
  var emailSheet = spreadsheet.getSheetByName(emailTabName);

  // Set Variables
  var taskNameCol = 3 // Col holding name of task
  var startCol = 4    // First col with an RA name
  var emailRow = 2    // Row with first email address
  var emailCol = 2    // Col holding email addresses

  var missingTasks = getMissingTasks(taskSheet, taskNameCol, startCol)
  sendReminders(emailSheet, emailRow, emailCol, missingTasks);
}

////////////
// Functions
////////////
function getMissingTasks(taskSheet, taskNameCol, startCol){
  // Returns a list of lists.  
  // Each sub-list contains all missing tasks in that column.

  var lastColumn = taskSheet.getLastColumn();
  var lastRow = taskSheet.getLastRow();
  var allTasksList = []
  
  // Iterate through each column for missing tasks
  for (var col = startCol; col <= lastColumn; col++) { 
    // Iterate through each row in the current column
    var colTaskList = [];
    for (var row = 1; row <= lastRow; row++) {
      var cellValue = taskSheet.getRange(row, col).getValue();     
      if (cellValue === '❌') {
        var taskName = taskSheet.getRange(row, taskNameCol).getValue();
        colTaskList.push(taskName)
      }
    }
    allTasksList.push(colTaskList)
  }
  return allTasksList;
}

function sendReminders(emailSheet, emailRow, emailCol, tasks) {
  /* Sends email reminders to RAs that have a missing task */
  var lastRow = emailSheet.getLastRow();
  
  // Get the range of the column from the start row to the last row
  var range = emailSheet.getRange(emailRow, emailCol, lastRow - emailRow + 1); //email row = startRow
  var values = range.getValues();

  // Iterate through each row and log the value
  for (var i = 0; i < values.length; i++) {
    var emailAddress = values[i][0]; 
    if (tasks[i].length > 0){
      // console.log('Email:' + emailAddress);
      // console.log('Title: [REMINDER] Missing Tasks')
      // console.log('You are missing the following tasks.')
      // console.log(tasks[i])
      // console.log('Please check your Flight Tracker to get these compelted.')
      let subject = "[REMINDER] Missing Tasks"

      let body = 'You are missing the following tasks:\n' +
                  tasks[i].map(task => `  - ${task}`).join('\n') + 
                  '\n\nPlease check your Flight Tracker to get these completed.';

      sendEmail(emailAddress, subject, body)
    }
  }
}

function sendEmail(recipient, subject, body, cc=null){
    var options = {};
    if (cc !== null && cc.length > 0) {
        options.cc = cc.join(',');
    }
    GmailApp.sendEmail(recipient, subject, body, options);    
}

// TODO - We don't need two loops.
// Just send the email after getting list of tasks
