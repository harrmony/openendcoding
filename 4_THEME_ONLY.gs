function themeBuild() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OpenEndCoding');
  const sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2);
  const data = dataRange.getValues();

///Check if there is a question in G2

  var cellValue = sheet.getRange('G2').getValue();
  
  if (!cellValue || cellValue.trim().split(/\s+/).length <= 3) {
    SpreadsheetApp.getUi().alert('Please enter a question into G2');
    throw new Error('Please enter a question into G2');
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

  // Sort data by length of response text (Column B) without modifying the sheet
  const sortedData = data.slice().sort((a, b) => b[1].length - a[1].length);

  // Extract top longest responses
  var numberOfResponsesForThemes = sheet3.getRange('C4').getValue();
  const topResponses = sortedData.slice(0, numberOfResponsesForThemes);

  // Concatenate the top responses
  const concatenatedResponses = topResponses.map(row => row[1]).join('\n');

  //Place the concatenated text into cell E2
  sheet.getRange('E2').setValue(concatenatedResponses);
  
  /////// NOW RANGE IS SORTED - WE SEND THE TEXT TO GPT

  // Get the GPT Propmt (user content) from cell L2
  const userContent = sheet.getRange('L2').getValue();

  const GPT_API = "ADD API KEY";
  const BASE_URL = "https://api.openai.com/v1/chat/completions";

  const headers = {
    "Content-Type": "application/json",
    "Authorization": `Bearer ${GPT_API}`
  };

  const payload = {
    "model": "gpt-4o-mini",
    "messages": [
      {
        "role": "system",
        "content": ""
      },
      {
        "role": "user",
        "content": userContent
      }
    ],
    "temperature": 0.3,
    "max_tokens": 2000,
    "top_p": 1,
    "frequency_penalty": 0,
    "presence_penalty": 0
  };

  const options = {
    headers: headers,
    method: "POST", // Use POST method to send the payload
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(BASE_URL, options);
  const jsonResponse = JSON.parse(response.getContentText());
  const result = jsonResponse.choices[0].message.content;

  // Log the result and place it into cell H2
  console.log(result);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Themes').getRange('A2').setValue(result);


  Logger.log('Themes created');

}
