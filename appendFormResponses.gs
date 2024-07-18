function appendFormResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSheet = ss.getSheetByName('Form Responses 1');
  var artFilesSheet = ss.getSheetByName('Art Files');
  
  // Get the data from Form Responses 1, excluding the header row
  var data = formResponsesSheet.getRange('B2:L' + formResponsesSheet.getLastRow()).getValues();
  
  // Find the next blank row in the Art Files sheet
  var lastRow = artFilesSheet.getLastRow();
  var nextRow = lastRow + 1;
  
  // Append the data to the next blank row in Art Files
  artFilesSheet.getRange(nextRow, 1, data.length, data[0].length).setValues(data);
}
