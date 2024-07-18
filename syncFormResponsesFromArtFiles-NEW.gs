function combineAndUpdateFormResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSheet = ss.getSheetByName('Form Responses 1');
  var artFilesSheet = ss.getSheetByName('Art Files');
  
  // Step 0: Synchronize Art Files changes back to Form Responses 1
  synchronizeArtFilesToFormResponses(formResponsesSheet, artFilesSheet);
  
  // Step 1: Combine responses in Form Responses 1 sheet
  combineFormResponses(formResponsesSheet);
  
  // Step 2: Append or update Form Responses into Art Files sheet
  appendOrUpdateFormResponses(formResponsesSheet, artFilesSheet);
}

function synchronizeArtFilesToFormResponses(formResponsesSheet, artFilesSheet) {
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
      var formTimestamp = new Date(formRow[0]);
      var artTimestamp = new Date(artRow[0]);

      // Update Form Responses only if Art Files has a more recent timestamp or if Form Responses has blanks
      if (artTimestamp > formTimestamp || formRow.some(cell => cell === '')) {
        formResponsesSheet.getRange(formRowIndex + 2, 1, 1, formRow.length).setValues([artRow]);
      }
    } else {
      // Append new row to Form Responses if not found
      formResponsesSheet.appendRow(artRow);
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
