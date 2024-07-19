function combineAndUpdateFormResponses() {
  // Configuration
  var config = {
    formResponsesSheetName: 'Form Responses 1',
    artFilesSheetName: 'Art Files',
    numColumns: 14,
    columns: {
      fileName: 5, // Index for "File or Folder Name" in Form Responses (0-based)
      timestamp: 0, // Index for Timestamp in Form Responses (0-based)
      emailAddress: 1, // Index for Email Address in Form Responses (0-based)
      artFileName: 3, // Index for "File or Folder Name" in Art Files (0-based)
      formColumns: {
        assetType: 2,
        assetName: 3,
        assetDescription: 4,
        fileOrFolderName: 5,
        status: 6,
        priority: 7,
        startDate: 8,
        endDate: 9,
        assignedTeamMembers: 10,
        optionalNotes: 11,
        driveLink: 12,
        backup1: 13,
        backup2: 14
      },
      artFileColumns: {
        assetType: 0,
        assetName: 1,
        assetDescription: 2,
        fileOrFolderName: 3,
        status: 4,
        priority: 5,
        startDate: 6,
        endDate: 7,
        assignedTeamMembers: 8,
        optionalNotes: 9,
        driveLink: 10,
        emailAddress: 11,
        backup1: 12,
        backup2: 13
      }
    }
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSheet = ss.getSheetByName(config.formResponsesSheetName);
  var artFilesSheet = ss.getSheetByName(config.artFilesSheetName);

  // Step 1: Update Form Responses from Art Files
  updateFormResponsesFromArtFiles(formResponsesSheet, artFilesSheet, config);

  // Step 2: Combine Responses in Form Responses 1
  combineFormResponses(formResponsesSheet, config);

  // Step 3: Append or Update Form Responses into Art Files
  appendOrUpdateFormResponses(formResponsesSheet, artFilesSheet, config);
}

function updateFormResponsesFromArtFiles(formResponsesSheet, artFilesSheet, config) {
  var formResponsesData = formResponsesSheet.getDataRange().getValues();
  var artFilesData = artFilesSheet.getDataRange().getValues();
  
  var formFileNames = formResponsesData.slice(1).map(row => row[config.columns.fileName]);
  var artFileNames = artFilesData.slice(1).map(row => row[config.columns.artFileName]);
  
  // Iterate through Art Files and update Form Responses if needed
  artFilesData.slice(1).forEach(artRow => {
    var fileName = artRow[config.columns.artFileName];
    var formRowIndex = formFileNames.indexOf(fileName);
    
    if (formRowIndex !== -1) {
      var formRow = formResponsesData[formRowIndex + 1];
      
      // Update relevant columns from Art Files to Form Responses
      for (var key in config.columns.formColumns) {
        formRow[config.columns.formColumns[key]] = artRow[config.columns.artFileColumns[key]];
      }
      
      // Preserve existing Email Address if available
      if (formRow[config.columns.emailAddress] === '') {
        formRow[config.columns.emailAddress] = artRow[config.columns.artFileColumns.emailAddress];
      }
      
      // Set the updated form row back to Form Responses sheet
      formResponsesSheet.getRange(formRowIndex + 2, 1, 1, formRow.length).setValues([formRow]);
    } else {
      // Append new row to Form Responses if not found
      var newRow = new Array(config.numColumns).fill('');
      newRow[config.columns.timestamp] = new Date();
      for (var key in config.columns.formColumns) {
        newRow[config.columns.formColumns[key]] = artRow[config.columns.artFileColumns[key]];
      }
      newRow[config.columns.emailAddress] = artRow[config.columns.artFileColumns.emailAddress];
      
      formResponsesSheet.appendRow(newRow);
    }
  });
}

function combineFormResponses(formResponsesSheet, config) {
  var data = formResponsesSheet.getDataRange().getValues();
  var recentEntries = {};
  var headers = data[0];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var timestamp = new Date(row[config.columns.timestamp]);
    var fileName = row[config.columns.fileName];
    
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

function appendOrUpdateFormResponses(formResponsesSheet, artFilesSheet, config) {
  var formResponsesData = formResponsesSheet.getRange('C2:O' + formResponsesSheet.getLastRow()).getValues();
  var lastRow = artFilesSheet.getLastRow();
  var artFilesData;
  
  if (lastRow > 1) {
    artFilesData = artFilesSheet.getRange(2, 1, lastRow - 1, config.numColumns).getValues();
  } else {
    artFilesData = [];
  }

  var fileMap = {};
  for (var i = 0; i < artFilesData.length; i++) {
    var fileName = artFilesData[i][config.columns.artFileName];
    fileMap[fileName] = i + 2; // Storing row index, starting from row 2
  }

  for (var j = 0; j < formResponsesData.length; j++) {
    var formData = formResponsesData[j];
    var formFileName = formData[config.columns.fileName];
    var rowIndex = fileMap[formFileName];

    if (rowIndex) {
      artFilesSheet.getRange(rowIndex, 1, 1, formData.length).setValues([formData]);
    } else {
      artFilesSheet.appendRow(formData);
    }
  }
}
