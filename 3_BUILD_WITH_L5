function useBatchIDToBuildQCodeFrame() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OpenEndCoding');
  
    // Get the current spreadsheet file and its parent folder
  var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getName().replace(/\s+/g, '_');
  var parentFolder = spreadsheetFile.getParents().next();
  const GPT_API = "ADD API KEY";

  //We need to download the file to the right folder 

  ///CHECK IF BATCH IS FINISHED

  // Check if there is text in L5
  var batchId = sheet.getRange('L5').getValue();
  if (!batchId || batchId.trim() === '') {
    SpreadsheetApp.getUi().alert('Error: There is no Batch ID in L5');
    throw new Error('Error: There is no Batch ID in L5');
  }

  // Check the batch status and download the output file once it's ready
  while (true) {
    // Create the options for the UrlFetchApp request
    const statusOptions = {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + GPT_API
      }
    };

    // Make the request to the OpenAI API to check the batch status
    const statusResponse = UrlFetchApp.fetch('https://api.openai.com/v1/batches/' + batchId, statusOptions);
    const statusResponseData = JSON.parse(statusResponse.getContentText());
     
    // Log the batch status
    Logger.log(statusResponseData);
   
    // Check if the batch is processed
    if (statusResponseData.status === 'completed') {
      const outputFileId = statusResponseData.output_file_id;
      
      // Create the options for the UrlFetchApp request to download the file content
      const downloadOptions = {
        method: 'get',
        headers: {
          'Authorization': 'Bearer ' + GPT_API
        }
       }
      
      

      // Make the request to the OpenAI API to download the file content
      const downloadResponse = UrlFetchApp.fetch('https://api.openai.com/v1/files/' + outputFileId + '/content', downloadOptions);
      const outputBlob = downloadResponse.getBlob();



      // Define dailyFolder

      var now = new Date();
      var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HHmm');
      var folderName = 'Created_' + formattedDate;

      // Check if the folder already exists
      var folders = parentFolder.getFoldersByName(folderName);
      var dailyFolder;
      if (folders.hasNext()) {
        // Folder exists
        dailyFolder = folders.next();
      } else {
        // Folder does not exist, create it
        dailyFolder = parentFolder.createFolder(folderName);
      }
      //////////


      // Save the file to Google Drive and store the file name in the global variable
      generatedFileName = `${sheetName}_${formattedDate}_${formattedTime}_OUTPUT.jsonl`;
      dailyFolder.createFile(outputBlob).setName(generatedFileName);

      // Log the success message
     Logger.log('Output file downloaded and saved to Google Drive');
     Logger.log('Generated file name: ' + generatedFileName);
     break;
    } else {
         SpreadsheetApp.getUi().alert('The batch process is still running, this can take around 30 minutes\n\nPlease check again later');
         throw new Error('API batch process still running');
      }
    
  }

  //// ADD THE NEW JSONL FILE NAME HERE

  var files = dailyFolder.getFilesByName(generatedFileName); 
  
  if (!files.hasNext()) {
    SpreadsheetApp.getUi().alert('Error: File with this name is not saved in todays folder');
    throw new Error('Error: File with this name is not saved in todays folder');
    return;
  }

  var jsonlFile = files.next();
  var jsonlContent = jsonlFile.getBlob().getDataAsString().split('\n');

  // Filter out empty lines and parse JSON
  var jsonObjects = jsonlContent
    .filter(line => line.trim() !== '') // Remove empty lines
    .map(line => {
      try {
        return JSON.parse(line);
      } catch (e) {
        Logger.log('Error parsing line: ' + line);
        return null;
      }
    })
    .filter(obj => obj !== null); // Remove null entries from failed parsing

  // Get all custom_id values from column A
  var customIdsRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  var customIds = customIdsRange.getValues().flat();

  // Create a map for quick lookup of row index by custom_id
  var customIdToRowIndexMap = {};
  customIds.forEach((id, index) => {
    customIdToRowIndexMap[id] = index + 2; // +2 because sheet rows are 1-indexed and we start from row 2
  });

  jsonObjects.forEach(obj => {
    var customId = parseInt(obj.custom_id);
    var message = obj.response.body.choices[0].message.content;
    var rowIndex = customIdToRowIndexMap[customId];
    
    if (rowIndex) {
      sheet.getRange(rowIndex, 3).setValue(message); // Column C is the 3rd column
    }
  });


 /////// NOW WE WILL BUILD THE Q CODE FRAME


  // Open the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OpenEndCoding');
  var sheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var cellE4Value = sheetSettings.getRange('E4').getValue();
  
  // Get all data from the sheet
  var data = sheet.getDataRange().getValues();
  
  // Read themes from cell K4 - since this includes the "Other","Nonsense" options
  var themesText = sheet.getRange('K4').getValue();
  var themes = themesText.split(/\d+\. /).slice(1).map(function(item, index) {
    return { value: (index + 1).toString(), name: item.trim() };
  });

  // Override the value for the last two themes, if there are at least 2 themes
  if (themes.length >= 2) {
    themes[themes.length - 2].value = "98";
    themes[themes.length - 1].value = "99";
  }

  // Initialize the XML output structure based on the content of cell E4
  var output;

  if (cellE4Value === "Overlapping (Many Categories per Response)") {
    // If cell E4 contains "Overlapping (Many Categories per Response)", use this output
    output = [
      '<?xml version="1.0" encoding="utf-8"?>',
      '<!--This file was exported by Q, and contains the data to code a text variable.-->',
      '<!--Contact Support if you would like more information about this file format.-->',
      '<!--\'type\' may be SingleResponse or MultipleResponse-->',
      '<QCodes type="MultipleResponse">',
      '  <!--For MultipleResponse: \'value\' should be an integer, starting at 1.  When back coding it will be used to match the variable in the back-coding question.-->',
      '  <!--For SingleResponse: \'value\' can be any number.  When back coding it will be used to match the value in the back-coding question.-->',
      '<Code name="Missing Data" missingValues="true">',
      '<Notes />',
      '<!--The text in each Match must be lower case, have leading and trailing spaces stripped.-->',
      '<Matches>',
      '  <Match></Match>',
      '</Matches>',
      '</Code>'
    ]; 
  } else if (cellE4Value === "Mutually Exclusive (One Category per Response)") {
    // If cell E4 contains "Mutually Exclusive (One Category per Response)", use this output
    output = [
      '<?xml version="1.0" encoding="utf-8"?>',
      '<!--This file was exported by Q, and contains the data to code a text variable.-->',
      '<!--Contact Support if you would like more information about this file format.-->',
      '<!--\'type\' may be SingleResponse or MultipleResponse-->',
      '<QCodes type="SingleResponse">',
      '  <!--For MultipleResponse: \'value\' should be an integer, starting at 1.  When back coding it will be used to match the variable in the back-coding question.-->',
      '  <!--For SingleResponse: \'value\' can be any number.  When back coding it will be used to match the value in the back-coding question.-->',
      '<Code name="Missing Data" missingValues="true">',
      '<Notes />',
      '<!--The text in each Match must be lower case, have leading and trailing spaces stripped.-->',
      '<Matches>',
      '  <Match></Match>',
      '</Matches>',
      '</Code>'
    ]; 
  } else {
    // Optionally handle the case where E4 contains an unexpected value
    Logger.log('Unexpected value in cell E4: ' + cellE4Value);
  }
  

  // Loop through each theme
    themes.forEach(function(theme) {
    var matches = [];

    // Loop through the rows starting from row 2
    for (var j = 1; j < data.length; j++) {
      var response = data[j][1]; // Column B (responses)
      var codesText = data[j][3]; // Column D (codes)
      
      // Use a regular expression to match the exact theme value
      var regex = new RegExp('\\b' + theme.value + '\\b');
      
      if (typeof codesText === 'string' && regex.test(codesText)) {
        if (typeof response === 'string') {
          // Sanitize the response for XML
          var sanitizedResponse = response.toLowerCase().replace(/&/g, "&amp;");
          matches.push("<Match>" + sanitizedResponse + "</Match>");
        }
      }
    }

    if (matches.length > 0) {
      output.push("<Code name=\"" + theme.name + "\" value=\"" + theme.value + "\">");
      output.push("<Matches>");
      output = output.concat(matches);
      output.push("</Matches>");
      output.push("</Code>");
    }
  });


  // Add the closing tag for the XML
  output.push('</QCodes>');

  // Combine all parts into the final XML
  var finalOutput = output.join("\n");

  // Create the output file name with date and time
  var outputFileName = `QThemeOutput_${formattedDate}_${formattedTime}.QCodes`;

  // Create a new text file in Google Drive with the XML output
  var file = dailyFolder.createFile(outputFileName, finalOutput, MimeType.PLAIN_TEXT);
  Logger.log('File URL: ' + file.getUrl());
  SpreadsheetApp.getUi().alert('The Q file ( ' + outputFileName + ' ) has been saved.\n\nPlease check this folder: G:\Shared drives\Team Drive\2. Projects\1. Administration\Open-End Coding - GPT\ '+ dailyFolder.getName());

  Logger.log('useBatchFileToBuildQCodeFrame completed.');
}
