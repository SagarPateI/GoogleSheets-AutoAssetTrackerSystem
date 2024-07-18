function combineAndUpdateFormResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSheet = ss.getSheetByName('Form Responses 1');
  var artFilesSheet = ss.getSheetByName('Art Files');
  
  // Step 1: Update Form Responses 1 with data from Art Files
  updateFormResponsesFromArtFiles(formResponsesSheet, artFilesSheet);
  
  // Step 2: Combine responses in Form Responses 1 sheet
  combineFormResponses(formResponsesSheet);
  
  // Step 3: Append or update Form Responses into Art Files sheet
  appendOrUpdateFormResponses(formResponsesSheet, artFilesSheet);
}

function updateFormResponsesFromArtFiles(formResponsesSheet, artFilesSheet) {
  var formResponsesData = formResponsesSheet.getDataRange().getValues();
  var artFilesData = artFilesSheet.getDataRange().getValues();
  
  var formFileNames = formResponsesData.slice(1).map(row => row[5]); // Assuming "File or Folder Name" is the 6th column (index 5)
  var artFileNames = artFilesData.slice(1).map(row => row[3]); // Assuming "File or Folder Name" is the 4th column (index 3)
  
  // Iterate through Art Files and update Form Responses if needed
  artFilesData.slice(1).forEach(artRow => {
    var fileName = artRow[3]; // Assuming "File or Folder Name" is the 4th column (index 3)
    var artRowIndex = artFilesData.indexOf(artRow);
    var formRowIndex = formFileNames.indexOf(fileName);
    
    if (formRowIndex !== -1) {
      var formRow = formResponsesData[formRowIndex + 1];
      
      // Update relevant columns from Art Files to Form Responses
      formRow[0] = new Date(); // Update Timestamp to current date/time
      formRow[1] = ''; // Clear Email Address column
      
      // Update the rest of the columns based on Art Files
      formRow[2] = artRow[0]; // Asset Type
      formRow[3] = artRow[1]; // Asset Name
      formRow[4] = artRow[2]; // Asset Description
      formRow[5] = artRow[3]; // File or Folder Name
      formRow[6] = artRow[4]; // Status
      formRow[7] = artRow[5]; // Priority
      formRow[8] = artRow[6]; // Start Date
      formRow[9] = artRow[7]; // End Date
      formRow[10] = artRow[8]; // Assigned Team Member(s)
      formRow[11] = artRow[9]; // Issues or Optional Notes?
      formRow[12] = artRow[10]; // Agalleius Google Drive Link
      
      // Set the updated form row back to Form Responses sheet
      formResponsesSheet.getRange(formRowIndex + 2, 1, 1, formRow.length).setValues([formRow]);
    } else {
      // Append new row to Form Responses if not found
      formResponsesSheet.appendRow([
        new Date(), // Timestamp
        '', // Email Address
        artRow[0], // Asset Type
        artRow[1], // Asset Name
        artRow[2], // Asset Description
        artRow[3], // File or Folder Name
        artRow[4], // Status
        artRow[5], // Priority
        artRow[6], // Start Date
        artRow[7], // End Date
        artRow[8], // Assigned Team Member(s)
        artRow[9], // Issues or Optional Notes?
        artRow[10] // Agalleius Google Drive Link
      ]);
    }
  });
}

function combineFormResponses(formResponsesSheet) {
  var data = formResponsesSheet.getDataRange().getValues();
  var recentEntries = {};
  var headers = data[0];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var timestamp = new Date(row[0]);
    var fileName = row[5]; // Assuming "File or Folder Name" is the 6th column (index 5)
    
    if (recentEntries[fileName]) {
      if (timestamp > recentEntries[fileName].timestamp) {
        recentEntries[fileName] = {row: row, timestamp: timestamp};
      }
    } else {
      recentEntries[fileName] = {row: row, timestamp: timestamp};
    }
  }
  
  formResponsesSheet.clear();
  formResponsesSheet.appendRow(headers);
  
  for (var key in recentEntries) {
    formResponsesSheet.appendRow(recentEntries[key].row);
  }
}

function appendOrUpdateFormResponses(formResponsesSheet, artFilesSheet) {
  var formResponsesData = formResponsesSheet.getRange('C2:M' + formResponsesSheet.getLastRow()).getValues();
  var artFilesData = artFilesSheet.getRange(2, 1, artFilesSheet.getLastRow() - 1, 12).getValues();
  
  var fileMap = {};
  for (var i = 0; i < artFilesData.length; i++) {
    var fileName = artFilesData[i][3]; // Assuming "File or Folder Name" is the 4th column (index 3)
    fileMap[fileName] = i + 2; // Storing row index, starting from row 2
  }
  
  for (var j = 0; j < formResponsesData.length; j++) {
    var formData = formResponsesData[j];
    var formFileName = formData[3]; // Assuming "File or Folder Name" is the 4th column (index 3)
    var rowIndex = fileMap[formFileName];
    
    if (rowIndex) {
      artFilesSheet.getRange(rowIndex, 1, 1, formData.length).setValues([formData]);
    } else {
      artFilesSheet.appendRow(formData);
    }
  }
}
