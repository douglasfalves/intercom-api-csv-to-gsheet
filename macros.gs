function apagaColunas() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("CSV");

  var range = sheet.getRange("A:B");
  range.clear();
  range = sheet.getRange("D:D");
  range.clear();
  range = sheet.getRange("J:O");
  range.clear();
};

function ocultaColunas(){
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("CSV");

  sheet.hideColumns(1, 2);// A:B
  sheet.hideColumns(4, 1);// D
  sheet.hideColumns(10, 6);// J:N
};

function mostraColunas() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("CSV");

  sheet.showColumns(1, 2);// A:B
  sheet.showColumns(4, 1);// D
  sheet.showColumns(10, 6);// J:N 
};