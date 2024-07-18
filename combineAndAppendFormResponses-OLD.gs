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
      
      // Preserve existing Email Address if available
      if (formRow[1] === '') {
        formRow[1] = artRow[11]; // Use Email Address from Art Files if Email Address in Form Responses is blank
      }
      
      // Set the updated form row back to Form Responses sheet
      formResponsesSheet.getRange(formRowIndex + 2, 1, 1, formRow.length).setValues([formRow]);
    } else {
      // Append new row to Form Responses if not found
      formResponsesSheet.appendRow([
        new Date(), // Timestamp
        artRow[11], // Email Address
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
