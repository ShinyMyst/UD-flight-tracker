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
      let body = '<p>You are missing the following tasks:</p><ul>' +
              tasks.map(task => `<li>${task}</li>`).join('') + 
              '</ul><p>Please check your Flight Tracker to get these completed.</p>';

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
      console.log("FOUND an X")
      let taskCell = taskSheet.getRange(row, taskNameCol);  
      let richTextValue = taskCell.getRichTextValue(); 
      let taskName = richTextValue.getLinkUrl() 
        ? `<a href="${richTextValue.getLinkUrl()}">${richTextValue.getText()}</a>` 
        : richTextValue.getText();

      tasks.push(taskName)
    }
  }
  return tasks
};


function sendEmail(recipient, subject, body, cc=null){
    var options = {
      htmlBody: body
    };
    if (cc !== null && cc.length > 0) {
        options.cc = cc.join(',');
    }
    console.log("SENDING EMAIL")
    GmailApp.sendEmail(recipient, subject, '', options);    
};
