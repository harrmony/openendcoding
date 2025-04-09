function clearCells() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OpenEndCoding');
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Themes');
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var range1 = sheet.getRange("B2:C10900");
  var range2 = sheet.getRange("G2");
  var range3 = sheet.getRange("E2");
  var range4 = sheet.getRange("F1:F2");
  var range5 = sheet.getRange("L5");
  var range6 = sheet1.getRange("A2");
  var range7 = sheet.getRange("B2:B10900");
  var range8 = sheet2.getRange("A4")
  var range9 = sheet2.getRange("B4")
  var range10 = sheet2.getRange("C4")
  var range11 = sheet2.getRange("D4")


  // Clear or Update the content of the cells
  range1.clearContent();
  range2.clearContent();
  range2.setBackground("#dbead4");
  range3.clearContent();
  range4.clearContent();
  range5.clearContent();
  range6.clearContent();
  range6.setBackground("#dbead4");
  range7.setBackground("#dbead4");
  range8.setValue("12");
  range9.setValue("5");
  range10.setValue("150");
  range11.clearContent();

  SpreadsheetApp.getUi().alert('Open-End Text, Survey Question and Themes will be removed.\n\nYou can always undo this.');
  Logger.log('Cells from B2 to C10900, G2 and A2 on Themes have been cleared.');
}
