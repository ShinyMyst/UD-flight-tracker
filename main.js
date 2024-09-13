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
  
  var taskSubject = "[REMINDER] Missing Tasks"
  var taskCC = ['email'];


function main() {
  sendTaskReminder()
}


function sendTaskReminder(){
  // Iterate through each column (RA Names)
  let lastTaskCol = taskSheet.getLastColumn();
  for (let col = startCol; col <= lastTaskCol; col++) { 
    tasks = getTasks(col);
    // Prep and send email
    if (tasks.length > 0) {
      emailAddress = emailSheet.getRange(col-startCol+emailRow, emailCol).getValue();
      let body = 'You are missing the following tasks:\n' +
              tasks.map(tasks => `  - ${tasks}`).join('\n') + 
              '\n\nPlease check your Flight Tracker to get these completed.';

      //console.log(emailAddress)
      //console.log(emailSheet.getRange(col-startCol+emailRow, emailCol-1).getValue())
      //console.log(body)
      sendEmail(emailAddress, taskSubject, body, taskCC)
    }
  }
};


function getTasks(targetCol) {
  /* Returns a list of all tasks with an X in target column */
  let tasks = [];
  let lastRow = taskSheet.getLastRow();

  for (let row = 1; row <= lastRow; row++) {
    let cellValue = taskSheet.getRange(row, targetCol).getValue();     
    if (cellValue === '❌') {
      let taskName = taskSheet.getRange(row, taskNameCol).getValue();
      tasks.push(taskName)
    }
  }
  return tasks
};


function sendEmail(recipient, subject, body, cc=null){
    var options = {};
    if (cc !== null && cc.length > 0) {
        options.cc = cc.join(',');
    }
    GmailApp.sendEmail(recipient, subject, body, options);    
};
