function justBatch() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OpenEndCoding');
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2);
  const data = dataRange.getValues();
  sheet.getRange('F1').setValue("");
  sheet.getRange('F2').setValue("Please wait...");
  sheet.getRange('L5').setValue("");
  sheet.getRange('C2:C10900').setValue("");
  var range7 = sheet.getRange("B2:B10900");
  range7.setBackground("#dbead4");
  
  ///Check if there is a question in G2
  var cellValue = sheet.getRange('G2').getValue();
  
  if (!cellValue || cellValue.trim().split(/\s+/).length <= 3) {
    SpreadsheetApp.getUi().alert('Please enter a question into G2');
    throw new Error('Please enter a question into G2');
  }


  // Check for forbidden characters in B2:B10898
  var characterCheckRange = sheet.getRange('B2:B10898');
  var characterCheckValues = characterCheckRange.getValues();

  for (var i = 0; i < characterCheckValues.length; i++) {
    var characterCheckText = characterCheckValues[i][0];
    if (characterCheckText && /["<>]/.test(characterCheckText)) {
      SpreadsheetApp.getUi().alert('You have one of these symbols ( "  or  <  or  > ) in your responses.\n\nCheck and remove from cell B' + (i + 2));
      throw new Error('You have one of these symbols ( "  or  <  or  > ) in your responses.\n\nCheck and remove from cell B' + (i + 2));
    }
  }


  ///Check if there are numbers in A2
  var cellValue = sheet.getRange('A2').getValue();
  
  if (isNaN(cellValue) || cellValue === '') {
    SpreadsheetApp.getUi().alert('Please enter response numbers starting from cell A2');
    throw new Error('Please enter response numbers in column A');
  }

  ///Check if there is text in B2
  var cellValue = sheet.getRange('B2').getValue();
  
  if (!cellValue || cellValue.trim() === '') {
    SpreadsheetApp.getUi().alert('Please enter open end responses starting from cell B2');
    throw new Error('Please enter the open end responses starting from cell B2');
  }


  const GPT_API = "ADD GPT API KEY HERE";

  // Get the data in column X, starting from x2
  var wdataRange = sheet.getRange(2, 24, sheet.getLastRow() - 1, 1);
  var wdata = wdataRange.getValues();
  
  // Create an array to hold the JSONL lines
  var jsonlLines = [];
  
  // Loop through the data and add non-empty rows to the jsonlLines array
  for (var i = 0; i < wdata.length; i++) {
    if (wdata[i][0]) {
      jsonlLines.push(wdata[i][0]);
    }
  }
  
  // Join the lines with newline characters
  var jsonlContent = jsonlLines.join("\n");
  
  // Get the current date and time
  var now = new Date();
  var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HHmm');

  // Get the sheet file name and replace spaces with underscores
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getName().replace(/\s+/g, '_');

  // Create the file name with sheet name, date, and time
  var fileName = `${sheetName}_${formattedDate}_${formattedTime}.jsonl`;

  // Get the current spreadsheet file and its parent folder
  var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parentFolder = spreadsheetFile.getParents().next();

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



  // Create the new file in the same folder as the spreadsheet
  var file = dailyFolder.createFile(fileName, jsonlContent, MimeType.PLAIN_TEXT);

  Logger.log('File created with ID: ' + file.getId());

  // Details - Update File Name
  const FILE_NAME = fileName;

  // Create the payload for the file upload request
  const files = dailyFolder.getFilesByName(FILE_NAME);
  
  // Check if the file exists
  if (!files.hasNext()) {
    Logger.log('File not found');
    return;
  }
  
  const fileBlob = files.next().getBlob();

  // Create the payload for the file upload request
  const uploadPayload = {
    purpose: 'batch',
    file: fileBlob
  };

  // Create the options for the UrlFetchApp file upload request
  const uploadOptions = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + GPT_API
    },
    payload: uploadPayload
  };

  // Make the request to the OpenAI API to upload the file
  const uploadResponse = UrlFetchApp.fetch('https://api.openai.com/v1/files', uploadOptions);
  const uploadResponseData = JSON.parse(uploadResponse.getContentText());
  
  // Log the response and get the file ID
  Logger.log(uploadResponseData);
  const fileId = uploadResponseData.id;
  
  if (fileId) {
    // Create the payload for the batch creation request
    const batchPayload = {
      input_file_id: fileId,
      endpoint: '/v1/chat/completions',
      completion_window: '24h'
    };

    // Create the options for the UrlFetchApp batch creation request
    const batchOptions = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + GPT_API
      },
      payload: JSON.stringify(batchPayload)
    };

    // Make the request to the OpenAI API to create the batch
    const batchResponse = UrlFetchApp.fetch('https://api.openai.com/v1/batches', batchOptions);
    const batchResponseData = JSON.parse(batchResponse.getContentText());
    
    // Log the batch response and get the batch ID
    Logger.log(batchResponseData);
    const batchId = batchResponseData.id;
    sheet.getRange('L5').setValue(batchId);

    const startTime = new Date().getTime();
    const TIMEOUT_LIMIT = 1.2 * 60 * 1000; // 1.2 minutes in milliseconds (Default is 1.2 * 60...)

    // Check the batch status and download the output file once it's ready
    while (true) {
      const currentTime = new Date().getTime();
      const elapsedTime = currentTime - startTime;
 
      // Time-out Monitor
      if (elapsedTime > TIMEOUT_LIMIT) {
        SpreadsheetApp.getUi().alert('The GPT API needs longer to run, this process normally takes about 30 minutes.\n\nAfter 30 minutes has passed, please select the CHECK IF THE CODING PROCESS IS COMPLETE button.\n\nYou do not need to run the script again, it will be processing');
        throw new Error('The GPT API needs longer to run. Check back in 30 mins');
      }
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



      // PROGRESS TRACKER START Extract the request counts
      var requestCounts = statusResponseData.request_counts || { total: 0, completed: 0, failed: 0 };

      var progressSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OpenEndCoding');
      var total = requestCounts.total;
      var completed = requestCounts.completed;

      // Update the progress text in F1
      var progressText = "Number of responses: " + total + "\nNumber completed: " + completed;
      progressSheet.getRange('F1').setValue(progressText);

      // Calculate the percentage of completion
      var progressPercentage = total > 0 ? (completed / total) * 100 : 0;

      // Fill F2 with green blocks based on progress
      var cell = progressSheet.getRange('F2');
      var progressBlocks = Math.round(progressPercentage / 5);
      var progressBar = "▓".repeat(progressBlocks) + "░".repeat(20 - progressBlocks); // 20 blocks total
      cell.setValue(progressBar);

      // Force the spreadsheet to update
      SpreadsheetApp.flush();

      // PROGRESS TRACKER START Extract the request counts

      
      // Check if the batch is processed
      if (statusResponseData.status === 'completed') {
        const outputFileId = statusResponseData.output_file_id;
        
        // Create the options for the UrlFetchApp request to download the file content
        const downloadOptions = {
          method: 'get',
          headers: {
            'Authorization': 'Bearer ' + GPT_API
          }
        };

        // Make the request to the OpenAI API to download the file content
        const downloadResponse = UrlFetchApp.fetch('https://api.openai.com/v1/files/' + outputFileId + '/content', downloadOptions);
        const outputBlob = downloadResponse.getBlob();

        // Save the file to Google Drive and store the file name in the global variable
        generatedFileName = FILE_NAME.replace('.jsonl', '_OUTPUT.jsonl');
        dailyFolder.createFile(outputBlob).setName(generatedFileName);
  
        // Log the success message
        Logger.log('Output file downloaded and saved to Google Drive');
        Logger.log('Generated file name: ' + generatedFileName);
        break;
      } else {
        // Wait for 5 seconds before checking again
        Utilities.sleep(5000);
      }
    }
  }

  Logger.log('justBatch completed.');
  // Call the second function at the end of the first function
  useBatchFileToBuildQCodeFrame();
}

function useBatchFileToBuildQCodeFrame() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OpenEndCoding');
  
    // Get the current spreadsheet file and its parent folder
  var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parentFolder = spreadsheetFile.getParents().next();

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

  //// ADD THE NEW JSONL FILE NAME HERE

  var files = dailyFolder.getFilesByName(generatedFileName); 
  
  if (!files.hasNext()) {
    Logger.log('File not found.');
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
  
  // Read themes from cell K4 - Since this includes the "Nonsense" option
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
    SpreadsheetApp.getUi().alert('There is unexpected text in the Settings sheet, cell E4.\n\nPlease select either:\n\nOverlapping (Many Categories per Response)\n\nOR\n\nMutually Exclusive (One Category per Response)');
    throw new Error('Uexpected text in the Settings sheet, cell E4');
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


function existingThemeProcess() {
  justBatch();
}
