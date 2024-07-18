function appendOrUpdateFormResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSheet = ss.getSheetByName('Form Responses 1');
  var artFilesSheet = ss.getSheetByName('Art Files');

  // Get the data from Form Responses 1, excluding the header row
  var formResponsesData = formResponsesSheet.getRange('B2:L' + formResponsesSheet.getLastRow()).getValues();

  // Get existing data in Art Files
  var artFilesData = artFilesSheet.getRange(2, 1, artFilesSheet.getLastRow() - 1, 12).getValues();

  // Create a map of File or Folder Name to row index in Art Files
  var fileMap = {};
  for (var i = 0; i < artFilesData.length; i++) {
    var fileName = artFilesData[i][2]; // Assuming "File or Folder Name" is the 3rd column (index 2)
    fileMap[fileName] = i + 2; // Storing row index, starting from row 2
  }

  // Iterate through form responses and append or update rows in Art Files
  for (var j = 0; j < formResponsesData.length; j++) {
    var formData = formResponsesData[j];
    var formFileName = formData[2]; // Assuming "File or Folder Name" is the 3rd column (index 2)
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
