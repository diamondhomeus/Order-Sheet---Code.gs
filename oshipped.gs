var resSsId="1WKIBMnbG-iZ57B76g0pBpO1DRWCYJ4nq3BoSG-f0ypM";


function shippedFRPO ()
{
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var orderId=sheet.getRange(row, 4).getValue();
  
  
  var sheet2=SpreadsheetApp.openById(resSsId);
  var sheetOOS=sheet2.getSheetByName("OOS");
  
  var rowO=lookup(orderId,sheetOOS,4, 6,"row");
  
  
  
  sheetOOS.getRange(rowO, 3).setValue("O-SHIPPED");
  sheetOOS.getRange(rowO, 11).setValue("Shipped Order");
  

  

  
}


