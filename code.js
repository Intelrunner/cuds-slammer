var ss = SpreadsheetApp.getActiveSpreadsheet();

function findRows() {
  range = SpreadsheetApp.getActiveSheet().getLastRow();
  console.log(range)
  return range;
}

function getSheetName() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  Logger.log(sheet.getSheetName());
}

function renameSheet(newName) {
  // Get a reference to the sheet using its existing name
  // and then rename it using the setName() method.
  var sheet = ss.getSheets()[0];
  getSheetName();
  sheet.setName(newName);
  getSheetName();
  return 200
}


function createMain() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet(0);
  sheet = ss.getSheets()[1]
  sheet.setName("data")
  sheet = ss.getSheets()[0]
  sheet.setName("Main")
  return 200
};


function insertAverage() {
  var sheet = ss.getSheetByName("Main");
  sheet.insertColumns(2);
  sheet.getRange("B1").setValue("Average");
  return 200
}

function fillAverage() {
  var data = ss.getSheetByName("data")
  var tgt = ss.getSheetByName("Main").getRange("B2")

  tgt.setFormula("=AVERAGE('data'!2:2)")
  var x = findRows()
  range = (x)
  var sourceRange = tgt;
  sourceRange.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  return 200

}

function copyData() {
  var sheetFrom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data");
  var sheetTo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");

  // Copy from 17th row, 4th column, all rows for one column 
  var valuesToCopy = sheetFrom.getRange(1, 1, sheetFrom.getLastRow(), 1).getValues();

  //Paste to another sheet from first cell onwards
  sheetTo.getRange(1, sheetTo.getLastColumn() + 1, valuesToCopy.length, 1).setValues(valuesToCopy);
  return 200
};

function insertBreakdown() {
  //set sheet
  sheet = ss.getSheetByName("Main")
  sheet.insertColumns(3, 3)
  sheet.getRange("C1").setValue("On Demand")
  sheet.getRange("D1").setValue("1YR CUD")
  sheet.getRange("E1").setValue("3YR CUD")
  sheet.getRange("C2").setFormula("=ROUND(SUM(B2*.15),2)")
  var miliseconds = 1000
  Utilities.sleep(miliseconds)
  sheet.getRange("D2").setFormula("=SUM(ROUND(B2*.25),2)")
  Utilities.sleep(miliseconds)
  sheet.getRange("E2").setFormula("=SUM(ROUND(B2*.6),2)")
  tgt = sheet.getRange("C2:E2")
  tgt.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES)
  //add new column
}

function Main() {
  newName = 'data'
  findRows()
  createMain()
  copyData()
  //renameSheet(newName)
  insertAverage()
  fillAverage()
  insertBreakdown()
}






