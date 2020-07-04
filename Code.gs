/**
*
* Collect all cells that are bold and print their binary value to a new sheet.
*
* References
* https://developers.google.com/apps-script/reference/slides/text-style#getfontweight
* https://www.reddit.com/r/googlesheets/comments/hdt9ua/column_where_bold_figures_0_and_regular_figures_1/
*
*/

function updateBoldCells() {
  
  //  Declare variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  
  var sheet = spreadsheet.getActiveSheet();  
  var dataRange = sheet.getDataRange();
  var dataRangeBoldFormat = dataRange.getFontWeights();
  var newSheetValues = dataRangeBoldFormat;
  var newSheetName = "Text Weight Values";
  
  //  Create new sheet or empty sheet if it already exists
  try {
    var newSheet = spreadsheet.insertSheet(newSheetName);
  } catch (e) {
    var newSheet = spreadsheet.getSheetByName(newSheetName).clear();  
  }
  
  //  Update our new array so every value that is bolded becomes 1, other values become 0
  for (var x = 0; x < dataRangeBoldFormat.length; x++){
    for (var y = 0; y < dataRangeBoldFormat[x].length; y++){
      if (dataRangeBoldFormat[x][y] === "bold"){
        newSheetValues[x][y] = 1;
      } else {
        newSheetValues[x][y] = 0;
      }
    }
  }
  
  //  Add new array to new sheet
  newSheet.getRange(1, 1, newSheetValues.length, newSheetValues[0].length).setValues(newSheetValues);  
}

/** 
* 
* Create a menu option for script functions. Either run this function of reload your spreadsheet to use.
*
* References
* https://developers.google.com/apps-script/reference/document/document-app#getui
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
  .addItem('Create New Sheet with Bolded Cell Values from Current Sheet', 'updateBoldCells')
  .addToUi();
}
