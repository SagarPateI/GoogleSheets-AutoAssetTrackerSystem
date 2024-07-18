function combineFormResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSheet = ss.getSheetByName('Form Responses 1');
  
  // Get all the data from Form Responses 1
  var data = formResponsesSheet.getDataRange().getValues();
  
  // Map to store the most recent entries
  var recentEntries = {};
  
  // Header row
  var headers = data[0];
  
  // Iterate through data starting from the second row
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var timestamp = new Date(row[0]);
    var fileName = row[5]; // Assuming "File or Folder Name" is the 6th column (index 5)
    
    // Check if the fileName already exists in the map
    if (recentEntries[fileName]) {
      // Compare timestamps and keep the most recent row
      if (timestamp > recentEntries[fileName].timestamp) {
        recentEntries[fileName] = {row: row, timestamp: timestamp};
      }
    } else {
      // Add new entry to the map
      recentEntries[fileName] = {row: row, timestamp: timestamp};
    }
  }
  
  // Clear the Form Responses 1 sheet except for the header row
  formResponsesSheet.clear();
  formResponsesSheet.appendRow(headers);
  
  // Append the most recent entries back to the sheet
  for (var key in recentEntries) {
    formResponsesSheet.appendRow(recentEntries[key].row);
  }
}
