function main() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define the tab names using variables
  var emailTabName = 'Control'; 
  var tasksTabName = 'Tasks'; 

  // Get the sheets (tabs) by their names
  var emailSheet = spreadsheet.getSheetByName(emailTabName);
  var tasksSheet = spreadsheet.getSheetByName(tasksTabName);

  // Set Variables
  var taskCol = 3 
  var taskRow = 1
  var emailCol = 2
  var emailRow = 2
  // Email is assumed adjacent
    // Get the last column with data
  var lastColumn = tasksSheet.getLastColumn();
  
  // Get the last row with data
  var lastRow = tasksSheet.getLastRow();

  // Prep Collection
  var allTasks = []

  for (var col = 4; col <= lastColumn; col++) { // Change start col as a var
    // Iterate through each row in the current column
    //console.log("NEW RA");
    var raName = tasksSheet.getRange(taskRow, col).getValue();
    var taskList = [];
    for (var row = 1; row <= lastRow; row++) {
      // Get the value of the current cell
      var cellValue = tasksSheet.getRange(row, col).getValue();
      
      // Check if the cell value is ❌
      if (cellValue === '❌') {
        // Get values from the taskRow and taskCol
        var taskName = tasksSheet.getRange(row, taskCol).getValue();
        taskList.push(taskName)
      }
    }
  console.log(raName)
  console.log(taskList)
  //taskDict[raName] = taskList
  allTasks.push(taskList)
  }
  console.log(allTasks)

  // Send the Emails
    // Get the last row with data in the specified column
  var lastRow = emailSheet.getLastRow();
  
  // Get the range of the column from the start row to the last row
  var range = emailSheet.getRange(emailRow, emailCol, lastRow - emailRow + 1); //email row = startRow
  // Get the values from the range
  var values = range.getValues();

  // Iterate through each row and log the value
  for (var i = 0; i < values.length; i++) {
    var cellValue = values[i][0]; // Get the value in the cell
    console.log('Email:' + cellValue);
    console.log(allTasks[i])
  }
}

function getMissingTasks(){
  var lastColumn = tasksSheet.getLastColumn();
  var lastRow = tasksSheet.getLastRow();
  var allTasks = []

  for (var col = 4; col <= lastColumn; col++) { // Change start col as a var
    // Iterate through each row in the current column
    //console.log("NEW RA");
    var raName = tasksSheet.getRange(taskRow, col).getValue();
    var taskList = [];
    for (var row = 1; row <= lastRow; row++) {
      // Get the value of the current cell
      var cellValue = tasksSheet.getRange(row, col).getValue();
      
      // Check if the cell value is ❌
      if (cellValue === '❌') {
        // Get values from the taskRow and taskCol
        var taskName = tasksSheet.getRange(row, taskCol).getValue();
        taskList.push(taskName)
      }
    }
  console.log(raName)
  console.log(taskList)
  //taskDict[raName] = taskList
  allTasks.push(taskList)
  }
  console.log(allTasks)

}
