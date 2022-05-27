// Create Variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var newName = "data";

//Clasp Test

function findRows() {
  range = SpreadsheetApp.getActiveSheet().getLastRow();
  console.log(range);
  return range;
}

function getSheetName() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  Logger.log(sheet.getSheetName());
}

function createMain() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet(0);
  sheet = ss.getSheets()[1];
  sheet.setName("data");
  sheet = ss.getSheets()[0];
  sheet.setName("Main");
  return 200;
}

function copyData() {
  var sheetFrom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data");
  var sheetTo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");

  // Copy 1st row, 1st column, all rows for one column
  var valuesToCopy = sheetFrom
    .getRange(1, 1, sheetFrom.getLastRow(), 1)
    .getValues();

  //Paste to another sheet from first cell onwards
  sheetTo
    .getRange(1, sheetTo.getLastColumn() + 1, valuesToCopy.length, 1)
    .setValues(valuesToCopy);
  return 200;
}

function insertAverage() {
  var sheet = ss.getSheetByName("Main");
  sheet.insertColumns(2);
  sheet.getRange("B1").setValue("Average");
  return 200;
}

function fillAverage() {
  var data = ss.getSheetByName("data");
  var tgt = ss.getSheetByName("Main").getRange("B2");

  tgt.setFormula("=ROUND(AVERAGE('data'!2:2),2)");
  var x = findRows();
  range = x;
  var sourceRange = tgt;
  sourceRange.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  return 200;
}

//  This func:
// Sets the title for 4 columns
// Sets formulas in the newly titled columns
// Extends those formulas down to the nearest neighbor sets the title for 4 columns
// it
function insertBreakdown() {
  //set sheet
  sheet = ss.getSheetByName("Main");
  sheet.insertColumns(3, 3);
  sheet.getRange("C1").setValue("On Demand (15%)");
  sheet.getRange("D1").setValue("1YR CUD (25%)");
  sheet.getRange("E1").setValue("3YR CUD (60%)");
  sheet.getRange("F1").setValue("OD Price(hr)");
  sheet.getRange("G1").setValue("100% OD (mo)");
  //[TODO]sheet.getRange("H1").setValue("15/25/60 Plan")
  //[TODO]sheet.getRange("I1").setValue("Monthly Savings")
  //[TODO]sheet.getRange("J1").setValue("Annual Savings")

  sheet.getRange("C2").setFormula("=ROUND(SUM(B2*.15),2)");
  var miliseconds = 500;
  Utilities.sleep(miliseconds);
  sheet.getRange("D2").setFormula("=SUM(ROUND(B2*.25),2)");
  Utilities.sleep(miliseconds);
  sheet.getRange("E2").setFormula("=SUM(ROUND(B2*.6),2)");
  tgt = sheet.getRange("C2:E2");
  tgt.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

//time to add in the get prices function
function getPrices(sdesc) {
  var url = encodeURI(
    `https://us-central1-eric-playground-298616.cloudfunctions.net/grabber?desc=${sdesc}`
  );
  console.log(url);
  var response = UrlFetchApp.fetch(url);
  return response;
}

function writePrices() {
  var sheet = ss.getSheetByName("Main");
  var startRange = "A1";
  var endRange = "F" + findRows();
  var range = sheet.getRange(`${startRange}:${endRange}`);
  const numRows = findRows();
  i = 1;
  while ((i <= numRows, i++)) {
    cell = range.getCell(i, 1);
    desc = cell.getValue();
    sprice = getPrices(desc);
    price = range.getCell(i, 6);
    price.setValue(sprice);
    Utilities.sleep(200);
  }
  return 200;
}
function Main() {
  findRows();
  createMain();
  copyData();
  insertAverage();
  fillAverage();
  insertBreakdown();
  writePrices();
}
