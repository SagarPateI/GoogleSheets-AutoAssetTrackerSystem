// Configuration parameters
var FORM_RESPONSES_SHEET_NAME = 'Form Responses 1';
var ART_FILES_SHEET_NAME = 'Art Files';
var START_ROW_INDEX = 2; // Row where data starts (1-based index, usually 2 if headers are in the first row)

// Column indices for both sheets (0-based index)
var TIMESTAMP_COL = 0;
var EMAIL_COL = 1;
var ASSET_TYPE_COL = 2;
var ASSET_NAME_COL = 3;
var ASSET_DESC_COL = 4;
var FILE_NAME_COL = 5;
var STATUS_COL = 6;
var PRIORITY_COL = 7;
var START_DATE_COL = 8;
var END_DATE_COL = 9;
var TEAM_MEMBERS_COL = 10;
var ISSUES_COL = 11;
var DRIVE_LINK_COL = 12;
var BACKUP1_COL = 13;
var BACKUP2_COL = 14;

function combineAndUpdateFormResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSheet = ss.getSheetByName(FORM_RESPONSES_SHEET_NAME);
  var artFilesSheet = ss.getSheetByName(ART_FILES_SHEET_NAME);

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
  
  var formFileNames = formResponsesData.slice(1).map(row => row[FILE_NAME_COL]);
  var artFileNames = artFilesData.slice(1).map(row => row[FILE_NAME_COL]);
  
  // Iterate through Art Files and update Form Responses if needed
  artFilesData.slice(1).forEach(artRow => {
    var fileName = artRow[FILE_NAME_COL];
    var formRowIndex = formFileNames.indexOf(fileName);
    
    if (formRowIndex !== -1) {
      var formRow = formResponsesData[formRowIndex + 1];
      
      // Update relevant columns from Art Files to Form Responses
      formRow[ASSET_TYPE_COL] = artRow[ASSET_TYPE_COL];
      formRow[ASSET_NAME_COL] = artRow[ASSET_NAME_COL];
      formRow[ASSET_DESC_COL] = artRow[ASSET_DESC_COL];
      formRow[FILE_NAME_COL] = artRow[FILE_NAME_COL];
      formRow[STATUS_COL] = artRow[STATUS_COL];
      formRow[PRIORITY_COL] = artRow[PRIORITY_COL];
      formRow[START_DATE_COL] = artRow[START_DATE_COL];
      formRow[END_DATE_COL] = artRow[END_DATE_COL];
      formRow[TEAM_MEMBERS_COL] = artRow[TEAM_MEMBERS_COL];
      formRow[ISSUES_COL] = artRow[ISSUES_COL];
      formRow[DRIVE_LINK_COL] = artRow[DRIVE_LINK_COL];
      formRow[BACKUP1_COL] = artRow[BACKUP1_COL];
      formRow[BACKUP2_COL] = artRow[BACKUP2_COL];
      
      // Preserve existing Email Address if available
      if (formRow[EMAIL_COL] === '') {
        formRow[EMAIL_COL] = artRow[EMAIL_COL];
      }
      
      // Set the updated form row back to Form Responses sheet
      formResponsesSheet.getRange(formRowIndex + START_ROW_INDEX, 1, 1, formRow.length).setValues([formRow]);
    } else {
      // Append new row to Form Responses if not found
      formResponsesSheet.appendRow([
        new Date(), // Timestamp
        artRow[EMAIL_COL],
        artRow[ASSET_TYPE_COL],
        artRow[ASSET_NAME_COL],
        artRow[ASSET_DESC_COL],
        artRow[FILE_NAME_COL],
        artRow[STATUS_COL],
        artRow[PRIORITY_COL],
        artRow[START_DATE_COL],
        artRow[END_DATE_COL],
        artRow[TEAM_MEMBERS_COL],
        artRow[ISSUES_COL],
        artRow[DRIVE_LINK_COL],
        artRow[BACKUP1_COL],
        artRow[BACKUP2_COL]
      ]);
    }
  });
}

function combineFormResponses(formResponsesSheet) {
  var data = formResponsesSheet.getDataRange().getValues();
  var recentEntries = {};
  var headers = data[0];
  
  for (var i = START_ROW_INDEX - 1; i < data.length; i++) {
    var row = data[i];
    var timestamp = new Date(row[TIMESTAMP_COL]);
    var fileName = row[FILE_NAME_COL];
    
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
  var formResponsesData = formResponsesSheet.getRange(START_ROW_INDEX, ASSET_TYPE_COL + 1, formResponsesSheet.getLastRow() - START_ROW_INDEX + 1, BACKUP2_COL + 1 - ASSET_TYPE_COL).getValues();
  var lastRow = artFilesSheet.getLastRow();
  var artFilesData;
  
  if (lastRow >= START_ROW_INDEX) {
    artFilesData = artFilesSheet.getRange(START_ROW_INDEX, ASSET_TYPE_COL + 1, lastRow - START_ROW_INDEX + 1, BACKUP2_COL + 1 - ASSET_TYPE_COL).getValues();
  } else {
    artFilesData = [];
  }

  var fileMap = {};
  for (var i = 0; i < artFilesData.length; i++) {
    var fileName = artFilesData[i][FILE_NAME_COL - ASSET_TYPE_COL];
    fileMap[fileName] = i + START_ROW_INDEX;
  }

  for (var j = 0; j < formResponsesData.length; j++) {
    var formData = formResponsesData[j];
    var formFileName = formData[FILE_NAME_COL - ASSET_TYPE_COL];
    var rowIndex = fileMap[formFileName];

    if (rowIndex) {
      artFilesSheet.getRange(rowIndex, ASSET_TYPE_COL + 1, 1, formData.length).setValues([formData]);
    } else {
      artFilesSheet.appendRow(formData);
    }
  }
}
