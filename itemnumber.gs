var databaseSsId="1nwJE0i3qTvjO8KW8BhneMYOCOVf74hVgDKoE7mx9wmE"; //original inventory list

function updateitemnumber() {
 var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();

  
  var ss2=SpreadsheetApp.openById(databaseSsId);
  var sheetInvDb=ss2.getSheetByName("Inventory List"); 
  
  var sku = sheet.getRange(row, 3).getValue();
  var targetRow = lookup(sku, sheetInvDb, 5, 6, "row");  
  var itemnumber = sheet.getRange(row, col).getValue();

   if (targetRow != null) {

  sheetInvDb.getRange(targetRow, 6).setValue(itemnumber);
  
  
  Browser.msgBox("Database Updated!");
} else {
  Browser.msgBox("SKU not found in Database");
  return 0;
}
}


var databaseSsId="1nwJE0i3qTvjO8KW8BhneMYOCOVf74hVgDKoE7mx9wmE"; //original inventory list

function updatevariation() {
 var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();

  
  var ss2=SpreadsheetApp.openById(databaseSsId);
  var sheetInvDb=ss2.getSheetByName("Inventory List"); 
  
  var sku = sheet.getRange(row, 3).getValue();
  var targetRow = lookup(sku, sheetInvDb, 5, 6, "row");  
  var variation = sheet.getRange(row, col).getValue();

   if (targetRow != null) {

  sheetInvDb.getRange(targetRow, 7).setValue(variation);
  
  
  Browser.msgBox("Database Updated!");
} else {
  Browser.msgBox("SKU not found in Database");
  return 0;
}
}

