function combineAndUpdateFormResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSheet = ss.getSheetByName('Form Responses 1');
  var table1Sheet = ss.getSheetByName('Art Files'); // Sheet name remains "Art Files" but referred to as "Table1"

  // Step 1: Update Form Responses from Table1
  updateFormResponsesFromTable1(formResponsesSheet, table1Sheet);

  // Step 2: Combine Responses in Form Responses 1
  combineFormResponses(formResponsesSheet);

  // Step 3: Append or Update Form Responses into Table1
  appendOrUpdateFormResponses(formResponsesSheet, table1Sheet);
}

function updateFormResponsesFromTable1(formResponsesSheet, table1Sheet) {
  var formResponsesData = formResponsesSheet.getDataRange().getValues();
  var table1Data = table1Sheet.getDataRange().getValues();
  
  var formFileNames = formResponsesData.slice(1).map(row => row[5]); // Assuming "File or Folder Name" is the 6th column (index 5)
  var table1FileNames = table1Data.slice(1).map(row => row[3]); // Assuming "File or Folder Name" is the 4th column (index 3)
  
  // Debug: Log data for troubleshooting
  Logger.log('Form Responses Data: ' + JSON.stringify(formResponsesData));
  Logger.log('Table1 Data: ' + JSON.stringify(table1Data));
  
  // Iterate through Table1 and update Form Responses if needed
  table1Data.slice(1).forEach(table1Row => {
    var fileName = table1Row[3]; // Assuming "File or Folder Name" is the 4th column (index 3)
    var formRowIndex = formFileNames.indexOf(fileName);
    
    if (formRowIndex !== -1) {
      var formRow = formResponsesData[formRowIndex + 1];
      
      // Update relevant columns from Table1 to Form Responses
      formRow[2] = table1Row[0]; // Asset Type
      formRow[3] = table1Row[1]; // Asset Name
      formRow[4] = table1Row[2]; // Asset Description
      formRow[5] = table1Row[3]; // File or Folder Name
      formRow[6] = table1Row[4]; // Status
      formRow[7] = table1Row[5]; // Priority
      formRow[8] = table1Row[6]; // Start Date
      formRow[9] = table1Row[7]; // End Date
      formRow[10] = table1Row[8]; // Assigned Team Member(s)
      formRow[11] = table1Row[9]; // Issues or Optional Notes?
      formRow[12] = table1Row[10]; // Agalleius Google Drive Link
      formRow[13] = table1Row[11]; // Backup 1
      formRow[14] = table1Row[12]; // Backup 2
      
      // Preserve existing Email Address if available
      if (formRow[1] === '') {
        formRow[1] = table1Row[11]; // Use Email Address from Table1 if Email Address in Form Responses is blank
      }
      
      // Set the updated form row back to Form Responses sheet
      formResponsesSheet.getRange(formRowIndex + 2, 1, 1, formRow.length).setValues([formRow]);
    } else {
      // Append new row to Form Responses if not found
      formResponsesSheet.appendRow([
        new Date(), // Timestamp
        table1Row[11], // Email Address
        table1Row[0], // Asset Type
        table1Row[1], // Asset Name
        table1Row[2], // Asset Description
        table1Row[3], // File or Folder Name
        table1Row[4], // Status
        table1Row[5], // Priority
        table1Row[6], // Start Date
        table1Row[7], // End Date
        table1Row[8], // Assigned Team Member(s)
        table1Row[9], // Issues or Optional Notes?
        table1Row[10], // Agalleius Google Drive Link
        table1Row[11], // Backup 1
        table1Row[12]  // Backup 2
      ]);
    }
  });
}

function combineFormResponses(formResponsesSheet) {
  var data = formResponsesSheet.getDataRange().getValues();
  var recentEntries = {};
  var headers = data[0];
  
  // Debug: Log data for troubleshooting
  Logger.log('Form Responses Data: ' + JSON.stringify(data));
  
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

function appendOrUpdateFormResponses(formResponsesSheet, table1Sheet) {
  var formResponsesData = formResponsesSheet.getRange('C2:O' + formResponsesSheet.getLastRow()).getValues();
  var lastRow = table1Sheet.getLastRow();
  var table1Data;
  
  if (lastRow > 1) {
    table1Data = table1Sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  } else {
    table1Data = [];
  }

  var fileMap = {};
  for (var i = 0; i < table1Data.length; i++) {
    var fileName = table1Data[i][3]; // Assuming "File or Folder Name" is the 4th column (index 3)
    fileMap[fileName] = i + 2; // Storing row index, starting from row 2
  }

  // Debug: Log file map for troubleshooting
  Logger.log('File Map: ' + JSON.stringify(fileMap));

  for (var j = 0; j < formResponsesData.length; j++) {
    var formData = formResponsesData[j];
    var formFileName = formData[3]; // Assuming "File or Folder Name" is the 4th column (index 3)
    var rowIndex = fileMap[formFileName];

    if (rowIndex) {
      table1Sheet.getRange(rowIndex, 1, 1, formData.length).setValues([formData]);
    } else {
      table1Sheet.appendRow(formData);
    }
  }
}
