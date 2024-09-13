// Prepare Vars
var programTabName = 'Programs'; 
var emailTabName = 'Control'; 
var headerRow = 2
var monthCol = "Month:"

// Get Sheets
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var programSheet = spreadsheet.getSheetByName(programTabName);
var emailSheet = spreadsheet.getSheetByName(emailTabName);

function onEdit(e){
  var range = e.range;

  // Program Tab Operations
  if (range.getSheet().getName() == programTabName) {
      var headers = programSheet.getRange(headerRow, 1, 1, programSheet.getLastColumn()).getValues()[0];
      var monthColIndex = headers.indexOf(monthCol) + 1;
      // If Month Column was edited
      if (range.getColumn() == monthColIndex) {
          var editedValue = range.getValue();
          Logger.log("Edited value in 'Month:' column: " + editedValue);

      }

  }
}



