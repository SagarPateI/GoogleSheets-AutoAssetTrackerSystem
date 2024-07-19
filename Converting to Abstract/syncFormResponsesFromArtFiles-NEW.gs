// Configuration
var config = {
  formResponsesSheetName: 'Form Responses 1',
  targetTableSheetName: 'Art Files',
  numColumns: 14,
  columns: {
    formResponsesColumns: {
      timestamp: 0, 
      emailAddress: 1,
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

function combineAndUpdateFormResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSheet = ss.getSheetByName(config.formResponsesSheetName);
  var artFilesSheet = ss.getSheetByName(config.targetTableSheetName);

  // Step 1: Update Form Responses from Art Files
  updateFormResponsesFromArtFiles(formResponsesSheet, artFilesSheet);

  // Step 2: Combine Responses in Form Responses 1
  combineFormResponses(formResponsesSheet);

  // Step 3: Append or Update Form Responses into Art Files
  appendOrUpdateFormResponses(formResponsesSheet, artFilesSheet);
}

function updateFormResponsesFromArtFiles(formResponsesSheet, artFilesSheet) {
  var formResponsesData = formResponsesSheet.getDataRange().getValues();
  var artFilesData = artFilesSheet.getDataRange().getValues();
  
  Logger.log('Form Responses Data: ' + JSON.stringify(formResponsesData));
  Logger.log('Art Files Data: ' + JSON.stringify(artFilesData));
  
  var formFileNames = formResponsesData.slice(1).map(row => row[config.columns.formResponsesColumns.fileOrFolderName]);
  var artFileNames = artFilesData.slice(1).map(row => row[config.columns.targetTableColumns.fileOrFolderName]);
  
  Logger.log('Form File Names: ' + JSON.stringify(formFileNames));
  Logger.log('Art File Names: ' + JSON.stringify(artFileNames));

  // Iterate through Art Files and update Form Responses if needed
  artFilesData.slice(1).forEach((artRow, index) => {
    var fileName = artRow[config.columns.targetTableColumns.fileOrFolderName];
    var formRowIndex = formFileNames.indexOf(fileName);
    
    Logger.log('Processing Art Row: ' + JSON.stringify(artRow));
    Logger.log('File Name: ' + fileName);
    Logger.log('Art Row Index: ' + (index + 2)); // Adjust for header row
    Logger.log('Form Row Index: ' + (formRowIndex + 2)); // Adjust for header row
    
    if (formRowIndex !== -1) {
      var formRow = formResponsesData[formRowIndex + 1];
      
      Logger.log('Updating Form Row: ' + JSON.stringify(formRow));
      
      // Update relevant columns from Art Files to Form Responses
      formRow[config.columns.formResponsesColumns.assetType] = artRow[config.columns.targetTableColumns.assetType];
      formRow[config.columns.formResponsesColumns.assetName] = artRow[config.columns.targetTableColumns.assetName];
      formRow[config.columns.formResponsesColumns.assetDescription] = artRow[config.columns.targetTableColumns.assetDescription];
      formRow[config.columns.formResponsesColumns.fileOrFolderName] = artRow[config.columns.targetTableColumns.fileOrFolderName];
      formRow[config.columns.formResponsesColumns.status] = artRow[config.columns.targetTableColumns.status];
      formRow[config.columns.formResponsesColumns.priority] = artRow[config.columns.targetTableColumns.priority];
      formRow[config.columns.formResponsesColumns.startDate] = artRow[config.columns.targetTableColumns.startDate];
      formRow[config.columns.formResponsesColumns.endDate] = artRow[config.columns.targetTableColumns.endDate];
      formRow[config.columns.formResponsesColumns.assignedTeamMembers] = artRow[config.columns.targetTableColumns.assignedTeamMembers];
      formRow[config.columns.formResponsesColumns.optionalNotes] = artRow[config.columns.targetTableColumns.optionalNotes];
      formRow[config.columns.formResponsesColumns.driveLink] = artRow[config.columns.targetTableColumns.driveLink];
      formRow[config.columns.formResponsesColumns.backup1] = artRow[config.columns.targetTableColumns.backup1];
      formRow[config.columns.formResponsesColumns.backup2] = artRow[config.columns.targetTableColumns.backup2];
      
      // Preserve existing Email Address if available
      if (formRow[config.columns.formResponsesColumns.emailAddress] === '') {
        formRow[config.columns.formResponsesColumns.emailAddress] = artRow[config.columns.targetTableColumns.emailAddress];
      }
      
      // Set the updated form row back to Form Responses sheet
      formResponsesSheet.getRange(formRowIndex + 2, 1, 1, formRow.length).setValues([formRow]);
    } else {
      // Append new row to Form Responses if not found
      formResponsesSheet.appendRow([
        new Date(), // Timestamp
        artRow[config.columns.targetTableColumns.emailAddress] || '', // Email Address
        artRow[config.columns.targetTableColumns.assetType], // Asset Type
        artRow[config.columns.targetTableColumns.assetName], // Asset Name
        artRow[config.columns.targetTableColumns.assetDescription], // Asset Description
        artRow[config.columns.targetTableColumns.fileOrFolderName], // File or Folder Name
        artRow[config.columns.targetTableColumns.status], // Status
        artRow[config.columns.targetTableColumns.priority], // Priority
        artRow[config.columns.targetTableColumns.startDate], // Start Date
        artRow[config.columns.targetTableColumns.endDate], // End Date
        artRow[config.columns.targetTableColumns.assignedTeamMembers], // Assigned Team Member(s)
        artRow[config.columns.targetTableColumns.optionalNotes], // Issues or Optional Notes?
        artRow[config.columns.targetTableColumns.driveLink], // Agalleius Google Drive Link
        artRow[config.columns.targetTableColumns.backup1], // Backup 1
        artRow[config.columns.targetTableColumns.backup2]  // Backup 2
      ]);
    }
  });
}

function combineFormResponses(formResponsesSheet) {
  var data = formResponsesSheet.getDataRange().getValues();
  var recentEntries = {};
  var headers = data[0];
  
  Logger.log('Form Responses Data: ' + JSON.stringify(data));

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var timestamp = new Date(row[config.columns.formResponsesColumns.timestamp]);
    var fileName = row[config.columns.formResponsesColumns.fileOrFolderName];
    
    Logger.log('Processing Row: ' + JSON.stringify(row));
    Logger.log('Timestamp: ' + timestamp);
    Logger.log('File Name: ' + fileName);
    
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
  var formResponsesData = formResponsesSheet.getRange('C2:O' + formResponsesSheet.getLastRow()).getValues();
  var artFilesData = artFilesSheet.getRange(2, 1, artFilesSheet.getLastRow() - 1, config.numColumns).getValues();

  Logger.log('Form Responses Data: ' + JSON.stringify(formResponsesData));
  Logger.log('Art Files Data: ' + JSON.stringify(artFilesData));

  var fileMap = {};
  for (var i = 0; i < artFilesData.length; i++) {
    var fileName = artFilesData[i][config.columns.targetTableColumns.fileOrFolderName];
    fileMap[fileName] = i + 2; // Storing row index, starting from row 2
  }

  Logger.log('File Map: ' + JSON.stringify(fileMap));

  var updatedRows = new Set();
  
  for (var j = 0; j < formResponsesData.length; j++) {
    var formData = formResponsesData[j];
    var formFileName = formData[config.columns.formResponsesColumns.fileOrFolderName];
    
    Logger.log('Processing Form Data: ' + JSON.stringify(formData));
    Logger.log('File Name: ' + formFileName);
    
    if (fileMap[formFileName] !== undefined) {
      var rowIndex = fileMap[formFileName];
      Logger.log('Mapping File Name: ' + formFileName + ' to Row Index: ' + rowIndex);
      
      // Update the row in Art Files
      artFilesSheet.getRange(rowIndex, 1, 1, formData.length).setValues([formData]);
      updatedRows.add(rowIndex);
    } else {
      // Append new row to Art Files
      artFilesSheet.appendRow(formData);
    }
  }
  
  // Clear any rows in Art Files that were not updated
  var artFilesRows = artFilesSheet.getRange(2, 1, artFilesSheet.getLastRow() - 1, config.numColumns).getValues();
  for (var k = 0; k < artFilesRows.length; k++) {
    if (!updatedRows.has(k + 2)) {
      artFilesSheet.deleteRow(k + 2);
    }
  }
}
