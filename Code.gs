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
  sheet.getRange("H1").setValue("Plan_OD");
  sheet.getRange("I1").setValue("Plan_1YR")
  sheet.getRange("J1").setValue("Plan_3YR");
  sheet.getRange("K1").setValue("SUDs")

  sheet.getRange("C2").setFormula("=ROUND(SUM(B2*.15))");
  var miliseconds = 500;
  Utilities.sleep(miliseconds);
  sheet.getRange("D2").setFormula("=SUM(ROUND(B2*.25))");
  Utilities.sleep(miliseconds);
  sheet.getRange("E2").setFormula("=SUM(ROUND(B2*.60))");
  tgt = sheet.getRange("C2:E2");
  tgt.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

/* TODO turn this into a batch process
* 
*/
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
  i = 2;
  while (i <= numRows) {
    cell = range.getCell(i, 1);
    desc = cell.getValue();
    sprice = getPrices(desc);
    price = range.getCell(i, 6);
    price.setValue(sprice);
    Utilities.sleep(200);
  }
  return 200;
}

/* 
  TODO: Write function to calculate 100% OD prices
   price * usage * 730 (hours) = OD cost
*/

function allOD() {
  const numRows = findRows();
  var sheet = ss.getSheetByName("Main");
  var tgt = sheet.getRange("G2").setFormula("=ROUND(SUM(B2*F2*730),2)");
  tgt.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

function planOD() {
  var sheet = ss.getSheetByName("Main");
  var tgt = sheet.getRange("H2").setFormula("=ROUND(SUM(C2*F2*730),2")
tgt.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES)
};

function plan1YR(){
  var sheet = ss.getSheetByName("Main");
  var tgt = sheet.getRange("I2").setFormula("=ROUND(SUM(D2*F2*730)*(1-.37),2")
tgt.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
};

function plan3YR(){
  const numRows = findRows();
  var sheet = ss.getSheetByName("Main");
  var tgt = sheet.getRange("I2").setFormula("=ROUND(SUM(E2*F2*730)*(1-.53),2")
tgt.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
};

function Main() {
  findRows();
  createMain();
  copyData();
  insertAverage();
  fillAverage();
  insertBreakdown();
  writePrices();
  allOD();
  planOD();
  plan1YR();
  plan3YR();
}
