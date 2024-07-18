function combineAndAppendFormResponses() {
  // Step 1: Combine responses in Form Responses 1 sheet
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
      var existingRow = recentEntries[fileName].row;
      var existingTimestamp = recentEntries[fileName].timestamp;
      var shouldUpdate = false;
      
      // Compare each cell and update only if current cell is not blank and existing cell is blank
      for (var col = 0; col < row.length; col++) {
        if (row[col] !== '' && (existingRow[col] === '' || existingTimestamp < timestamp)) {
          existingRow[col] = row[col];
          shouldUpdate = true;
        }
      }
      
      // Update timestamp if any cell was updated
      if (shouldUpdate) {
        recentEntries[fileName] = {row: existingRow, timestamp: timestamp};
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
  
  // Step 2: Append or update Form Responses into Art Files sheet
  var artFilesSheet = ss.getSheetByName('Art Files');
  var formResponsesData = formResponsesSheet.getRange('C2:M' + formResponsesSheet.getLastRow()).getValues();

  // Get existing data in Art Files
  var artFilesData = artFilesSheet.getRange(2, 1, artFilesSheet.getLastRow() - 1, 12).getValues();

  // Create a map of File or Folder Name to row index in Art Files
  var fileMap = {};
  for (var i = 0; i < artFilesData.length; i++) {
    var fileName = artFilesData[i][3]; // Assuming "File or Folder Name" is the 4th column (index 3)
    fileMap[fileName] = i + 2; // Storing row index, starting from row 2
  }

  // Iterate through form responses and append or update rows in Art Files
  for (var j = 0; j < formResponsesData.length; j++) {
    var formData = formResponsesData[j];
    var formFileName = formData[3]; // Assuming "File or Folder Name" is the 4th column (index 3)
    var rowIndex = fileMap[formFileName];

    if (rowIndex) {
      // Update existing row
      artFilesSheet.getRange(rowIndex, 1, 1, formData.length).setValues([formData]);
    } else {
      // Append new row
      artFilesSheet.appendRow(formData);
    }
  }
}
