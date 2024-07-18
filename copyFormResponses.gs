function copyFormResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSheet = ss.getSheetByName('Form Responses 1');
  var artFilesSheet = ss.getSheetByName('Art Files');
  
  // Get the data from Form Responses 1, excluding the header row
  var data = formResponsesSheet.getRange('B2:L' + formResponsesSheet.getLastRow()).getValues();
  
  // Clear existing data in Art Files
  artFilesSheet.getRange('A2:L' + artFilesSheet.getLastRow()).clearContent();
  
  // Copy the data to Art Files, starting from the second row
  artFilesSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
}
