function combineAndUpdateFormResponses() {
  // Configuration
  var config = {
    formResponsesSheetName: 'Form Responses 1',
    targetTableSheetName: 'Art Files',
    numColumns: 14,
    columns: {
      fileName: 5, // Index for "File or Folder Name" in Form Responses (0-based)
      timestamp: 0, // Index for Timestamp in Form Responses (0-based)
      emailAddress: 1, // Index for Email Address in Form Responses (0-based)
      targetFileName: 3, // Index for "File or Folder Name" in Target Table (0-based)
      formResponsesColumns: {
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
      targetTableColumns: {
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
        backup1: 11,
        backup2: 12
      }
    }
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSheet = ss.getSheetByName(config.formResponsesSheetName);
  var targetTableSheet = ss.getSheetByName(config.targetTableSheetName);

  // Step 1: Update Form Responses from Target Table
  updateFormResponsesFromTargetTable(formResponsesSheet, targetTableSheet, config);

  // Step 2: Combine Responses in Form Responses 1
  combineFormResponses(formResponsesSheet, config);

  // Step 3: Append or Update Form Responses into Target Table
  appendOrUpdateFormResponses(formResponsesSheet, targetTableSheet, config);
}

function updateFormResponsesFromTargetTable(formResponsesSheet, targetTableSheet, config) {
  var formResponsesData = formResponsesSheet.getDataRange().getValues();
  var targetTableData = targetTableSheet.getDataRange().getValues();
  
  var formFileNames = formResponsesData.slice(1).map(row => row[config.columns.fileName]);
  var targetFileNames = targetTableData.slice(1).map(row => row[config.columns.targetFileName]);
  
  // Iterate through Target Table and update Form Responses if needed
  targetTableData.slice(1).forEach(targetRow => {
    var fileName = targetRow[config.columns.targetFileName];
    var formRowIndex = formFileNames.indexOf(fileName);
    
    if (formRowIndex !== -1) {
      var formRow = formResponsesData[formRowIndex + 1];
      
      // Update relevant columns from Target Table to Form Responses
      for (var key in config.columns.formResponsesColumns) {
        formRow[config.columns.formResponsesColumns[key]] = targetRow[config.columns.targetTableColumns[key]];
      }
      
      // Preserve existing Email Address if available
      if (formRow[config.columns.emailAddress] === '') {
        formRow[config.columns.emailAddress] = targetRow[config.columns.targetTableColumns.emailAddress];
      }
      
      // Set the updated form row back to Form Responses sheet
      formResponsesSheet.getRange(formRowIndex + 2, 1, 1, formRow.length).setValues([formRow]);
    } else {
      // Append new row to Form Responses if not found
      var newRow = new Array(config.numColumns).fill('');
      newRow[config.columns.timestamp] = new Date();
      for (var key in config.columns.formResponsesColumns) {
        newRow[config.columns.formResponsesColumns[key]] = targetRow[config.columns.targetTableColumns[key]];
      }
      newRow[config.columns.emailAddress] = targetRow[config.columns.targetTableColumns.emailAddress];
      
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

function appendOrUpdateFormResponses(formResponsesSheet, targetTableSheet, config) {
  var formResponsesData = formResponsesSheet.getRange('C2:O' + formResponsesSheet.getLastRow()).getValues();
  var lastRow = targetTableSheet.getLastRow();
  var targetTableData;
  
  if (lastRow > 1) {
    targetTableData = targetTableSheet.getRange(2, 1, lastRow - 1, config.numColumns).getValues();
  } else {
    targetTableData = [];
  }

  var fileMap = {};
  for (var i = 0; i < targetTableData.length; i++) {
    var fileName = targetTableData[i][config.columns.targetFileName];
    fileMap[fileName] = i + 2; // Storing row index, starting from row 2
  }

  for (var j = 0; j < formResponsesData.length; j++) {
    var formData = formResponsesData[j];
    var formFileName = formData[config.columns.fileOrFolderName];
    var rowIndex = fileMap[formFileName];

    if (rowIndex) {
      targetTableSheet.getRange(rowIndex, 1, 1, config.numColumns).setValues([formData]);
    } else {
      targetTableSheet.appendRow(formData);
    }
  }
}
