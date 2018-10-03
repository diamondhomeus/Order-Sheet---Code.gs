
var returnSheetId ="1ErDlyC2lC0o1Z5lfiMNW-Fcrb0ciHgeOYe55hT8VvOA";
var orderSheetId ="1CFFhqYpKu3_2WdmNMkrjZt4u7MmowxJRTpkB-S00ByA";
var resSsId="1WKIBMnbG-iZ57B76g0pBpO1DRWCYJ4nq3BoSG-f0ypM";
var trackingSheetId="1JtU1zOMK8y8WsDLhElg1eZPDi76hGaA2kTFRQFKt_Ro";
var invId="1nwJE0i3qTvjO8KW8BhneMYOCOVf74hVgDKoE7mx9wmE";
var dedId="14i28_rwfKqyzG_kkX-m9mtuXN16h-nyqZJTeuQeq_no";



var amandaSheetId="1A8RqRPdAjq7F7wTSR_ywfqD-lyqlJmDqw77jy_bdHSE";
var shamsSheetId="1k3t6s2mO7LFvkjg5tIm8VoGq7ahaBx0_OkKRjQKUJ-Q";
var nazmusSheetId="1CnaeLDtyfylwhRPLvmjDbOkibTS6KOTleIIIOtwwuis";
var jeremySheetId="17oRbfebN25sY3SluQcxIbCdCZoaAtPMlKPoKcjf_Jp8";
var bradSheetId="1Uw6cn-x__Y2kwIIRik64HIiiUuHwA804z4wDsLtJS1U";
var rianneSheetId="12iQwvtZubUyGE8DKq4xW9DQnMlKIsXUnal4KymznCAQ";
var saraSheetId="10Z3GDhPnY0NyepwC24RKVZsCwtfUEyB5a7_Fr6BC-AM";
var matthewSheetId="1wtlzNqzczzpf02PvgoCTCLsoYDL03OfEcZGGATZPSPs";
var gageSheetId="1e35eb0jfvsLHQZQq_hR0Ec32ox_2QNCSuPsF6XZXALg";
var trevSheetId="1sGKwXVIsSq4oTKstW9ZaIEvdgq8sixn_PIzmgto6pj0";
var steveSheetId="1Gq1X8n-F8cSoNjlrF_58mWVIAEmf5KBC7E0a9EiteWQ";
var daveSheetId="1O_n5QKFx6shs3Ex3ifwCqGQb9st7P0MK1evCQwqO4yo";
var mmSheetId="18NxGFqvIKw1P7pYorFRVrquRp4ItY9LyPipaZqrMZwM";
var rohitSheetId="1QCoNwFambfuzo5gIf9FVL2wzkOFuwHnFm9Z3tb61sF8";
var reillySheetId="1vouGwVfb7pQjB-hohYHFXFOQqQ8defV68WlRoR1XlW8";
var domSheetId="1p4hICtGj-R59BjScOSvOZDLiZunNCPqzzNtXrqIPlNs";
var zainabSheetId="1jNmvOkr63LwiLXCfR4c31Tfy7IyP0ga0AEB2xxBtD-Q";



function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  
  ui.createMenu('Sign-In')
  
  .addItem('Updates', 'newupdates')
  .addSeparator()
  .addItem('Kris', 'Kris')
  .addItem('Gage', 'Gage')
  .addItem('Brad', 'Brad')
  .addItem('Amanda', 'Amanda')
  .addItem('Brodie', 'Brodie')
  .addItem('Rianne', 'Rianne')
  .addItem('Trevor', 'Trevor')
  .addItem('Steve', 'Steve')
  .addSeparator()
  .addItem('Signout User', 'Signout')
  .addSeparator()
  .addItem('Sign-in to All Sites', 'link3')
  .addToUi(); 
  
  ui.createMenu('Shipping')
  .addItem('Jump to Orders', 'Jumpdown')
  .addSeparator()
  .addItem('Process CAD', 'exportCADorder')
  .addItem('Process UK', 'exportUKorder')
  .addSeparator()
  .addItem('OOS', 'oostransfer')
  .addItem('LP', 'lptransfer')
  .addItem('PO BOX', 'potransfer')
  .addItem('FR', 'frtransfer')
  .addItem('OI', 'oitransfer')
  .addItem('DUPLICATE', 'duptransfer')
  .addItem('INVALID ADDRESS', 'invalidtransfer')
  .addItem('SHIPPING COSTS ERROR', 'shippingcostserrortransfer')
  .addSeparator()
  .addItem('SORT ORDERS - BEFORE 5PM', 'sortorders')
  .addItem('SORT ORDERS - AFTER 5PM', 'sortordersafter5')
  .addSeparator()
  .addItem('NEW COUPON', 'newcoupon')
  .addItem('CLOSED ASIN', 'closedasin')
  .addSeparator()
  .addItem('PROMO - OI', 'promotransfer')
  .addItem('PROMO - OOS', 'promotransferoos')
  .addSeparator()
  .addItem('O-SHIPPED', 'shippedFRPO')
  .addSeparator()
  .addItem('Highlight Duplicates - Order ID', 'highLightDuplicates')
  .addItem('Reset Highlights', 'resetDuplicates')
  .addSeparator()
  .addItem('Item Number - Update Database', 'updateitemnumber')
  .addItem('Variation - Update Database', 'updatevariation')
  .addToUi(); 
  
  ui.createMenu('RES')
  .addItem('ZERO/ALL', 'zeroout')
  .addItem('ZERO/LEAVE COG', 'zerooutcog')
  .addItem('ZERO/COG', 'zerooutonlycog')
  .addSeparator()
  .addItem('ADV - ZERO/ALL', 'zerooutalladv')
  .addItem('ADV - ZERO/LEAVE COG', 'zerooutalladvcog')
  .addSeparator()
  .addItem('RF - ZERO/ALL', 'zerooutallRF')
  .addSeparator()
  .addItem('PARTIAL REFUND - Dollar', 'partialdollar')
  .addItem('PARTIAL REFUND - Percentage', 'partialpercent')
  .addSeparator()
  .addItem('PROCESS - RESET ORDER', 'resetorder')
  .addItem('PROCESS - SHIP OOS/FR/PO', 'oshippedorder')
  .addSeparator()
  .addItem('DIVIDE - 1 of 2 Qty - Adj Loss', 'qtyloss')
  .addItem('MULTIPLY - 1 of 2 Qty - Adj Loss', 'qtyloss2')
  
  .addSeparator()
  .addItem('OPEN ORDER - SUPPLIER', 'opensupplierorder')
  
  
  .addSeparator()
  .addItem('Transfer to PENDING', 'transfertoPENDING')
  .addItem('Transfer to OI', 'transfertoOI')
  .addItem('Transfer to LS', 'transfertoLS')
  .addItem('Transfer to ADV', 'transfertoADV')
  .addItem('Transfer to ESC', 'transfertoESC')
  .addItem('Transfer to RTN - Pending', 'transfertoReturns')
  .addItem('Transfer to RTN - Partial', 'transfertoReturnsPartial')
  .addItem('Transfer to RP', 'transfertoRP')
  .addItem('Transfer to RF', 'transfertoRF')
  .addItem('Transfer to DED - LOSS + COG', 'transfertoDEDcog')
  .addItem('Transfer to DED - LOSS ONLY', 'transfertoDEDloss')
  .addToUi(); 
  
  
  ui.createMenu('Scripts')
  .addItem('Check - Stock', 'importVariation')
  .addItem('Check - OnSale', 'checkSales')
  .addItem('Reset', 'formulastock')
  .addSeparator()
  .addItem('Update - OnSale', 'refreshOnSale')
  .addItem('OS - Add Coupon', 'addcoupon')
  .addItem('OS - Add Sale', 'addsale')
  
  .addSeparator()
  .addItem('Check PO/FR', 'updateOrderList')
  .addItem('Download UnShipped','updateOrderIds')
  .addSeparator()
  .addItem('OOS - Brodie', 'boostransfer')
  .addItem('LP - Brodie', 'blptransfer')
  .addItem('PO BOX - Brodie', 'bpotransfer')
  .addItem('FR - Brodie', 'bfrtransfer')
  .addItem('OI - Brodie', 'boitransfer')
  .addItem('DUPLICATE - CX - Brodie', 'bduptransfer')
  .addItem('INVALID ADDRESS', 'binvalidtransfer')
  .addItem('SHIPPING COSTS ERROR', 'bshippingcostserrortransfer')
  .addSeparator()
  .addItem('Pause OS Orders', 'pauseorders')
  .addItem('Pause WM Orders', 'pauseorderswm')
  .addItem('Pause AE Orders', 'pauseordersae')
  .addItem('Pause WF Orders', 'pauseorderswf')
  .addItem('Pause SG Orders', 'pauseorderssg')
  .addSeparator()
  .addItem('Run onEdit Manually', 'onEdit2')
  .addItem('Fix ConditionalFormatting', 'fixConditionalFormatting')
  //   .addItem('Delete - Skugrid', 'cleanup2') 
  
  .addToUi(); 
  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentDate=new Date();
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  var mysheet=ss.getSheetByName(sheetName);
  
  
  
  var lastrow = last_row(mysheet, 1);
  
  mysheet.setActiveCell(mysheet.getDataRange().offset(lastrow-1, 0, 1,         1));  
  
}



function exportUKorder() {
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet= rng.getSheet();
  
  
  var orderID = sheet.getRange(row, 4).getValue();
  var comments = sheet.getRange(row, 17).getValue();
  
  
  
  
  var cadtab = ss.getSheetByName("INT-Track");
  
  var lr = cadtab.getLastRow();
  
  cadtab.getRange(lr+1, 2).setValue(orderID);
  sheet.getRange(row, 17).setValue(comments+" INT - CAD").setDataValidation(null)
  
  
  
  
}



function exportCADorder() {
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet= rng.getSheet();
  
  
  var orderID = sheet.getRange(row, 4).getValue();
  var comments = sheet.getRange(row, 17).getValue();
  
  
  
  
  var cadtab = ss.getSheetByName("INT-Track");
  
  var lr = cadtab.getLastRow();
  
  cadtab.getRange(lr+1, 1).setValue(orderID);
  sheet.getRange(row, 17).setValue(comments+" INT-CAD").setDataValidation(null)
  
  
  
  
}



function exportCADordertotracking() {
  
  
  var ss2= SpreadsheetApp.openById("1JtU1zOMK8y8WsDLhElg1eZPDi76hGaA2kTFRQFKt_Ro");
  var cadsheet = ss2.getSheetByName("INT-Track")
  
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var cadtab = ss.getSheetByName("CAD");
  var orderID = cadtab.getRange("A2:A").getValues();
  
  var lr = cadsheet.getLastRow();
  
  cadsheet.getRange(lr+1, 1, orderID.length, orderID[0].length).setValues(orderID);
  
  
  
  
  
}

function partialdollar() {
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheetName=sheet.getName();
  
  
  var dollaramount = Browser.inputBox("Input - Refund Amount - Dollars");
  
  var SoldPrice = sheet.getRange(row, 9).getValue();
  
  
  var Diff = SoldPrice-dollaramount;
  
  var AmazonFees = Diff*0.15;
  
  sheet.getRange(row, 9).setValue(SoldPrice-dollaramount).setFontColor("orange");
  sheet.getRange(row, 10).setValue(AmazonFees).setFontColor("orange"); 
  sheet.getRange(row, 26).setValue("ISSUE").setFontColor("orange");
  
  
  
}

function partialpercent() {
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheetName=sheet.getName();
  
  
  var dollarpercent = Browser.inputBox("Input - Refund Amount - Percentage");
  dollarpercent="0."+dollarpercent;
  
  var SoldPrice = sheet.getRange(row, 9).getValue();
  var discount = SoldPrice*dollarpercent;
  
  
  
  var Diff = SoldPrice-discount
  
  var AmazonFees = Diff*0.15;
  
  
  
  sheet.getRange(row, 9).setValue(SoldPrice-discount).setFontColor("orange");
  sheet.getRange(row, 10).setValue(AmazonFees).setFontColor("orange"); 
  
  sheet.getRange(row, 26).setValue("ISSUE").setFontColor("orange");
  
  var amazonrefund = Utilities.formatString('%11.2f', SoldPrice-Diff-0.01);
  
  amazonrefund=amazonrefund
  
  
  
  
  Browser.msgBox('Refund Amount on Amazon : '+amazonrefund);
}






function Jumpdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mysheet = ss.getActiveSheet();
  
  var lastrow = mysheet.getLastRow();
  
  mysheet.setActiveCell(mysheet.getDataRange().offset(lastrow-1, 0, 1,         1));  
};





function transferFromPending()
{
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName("Pending")
  var lr=sheet.getLastRow();
  var values=sheet.getRange("A8:AN"+lr).getValues();
  var formats=sheet.getRange("A8:AN"+lr).getBackgrounds();
  var formulas=sheet.getRange("A8:AN"+lr).getFormulasR1C1()
  var fonts=sheet.getRange("A8:AN"+lr).getFontColors();
  
  
  
  var a=10
  
  var currentDate=new Date();
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  var sheet1=ss.getSheetByName(sheetName);
  var month1=currentDate.getMonth();
  
  
  var firstDayPrevMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 15);
  sheetName=Utilities.formatDate(firstDayPrevMonth, ss.getSpreadsheetTimeZone(), "MMMM-yyyy");
  var sheet2=ss.getSheetByName(sheetName);
  
  var arr1=[];
  var bg1=[];
  var arr2=[];
  var bg2=[];
  
  var colD1=[];
  var colG1=[];
  var colH1=[];
  
  var colD2=[];
  var colG2=[];
  var colH2=[];
  
  
  for (var i=0; i<values.length; i++)
  {
    for (var j=0; j<values[i].length; j++)
    {
      if(formulas[i][j]!=""){values[i][j]=formulas[i][j]};
      
    }
    
    
    
    
    var monthTemp=values[i][0].getMonth(); //current month
    if(monthTemp==month1)
    {
      arr1.push(values[i]);
      bg1.push(formats[i]);
      
      var fc=fonts[i][4-1];
      if(fonts[i][4-1]=="#000000"){
        fc="#1155cc"
      }
      colD1.push(fc);
      
      
      var fc=fonts[i][7-1];
      if(fonts[i][7-1]=="#000000"){
        fc="#1155cc"
      }
      colG1.push(fc);
      
      var fc=fonts[i][8-1];
      if(fonts[i][8-1]=="#000000"){
        fc="#1155cc"
      }
      colH1.push(fc);
      
      
    }
    
    else
    {
      arr2.push(values[i]);
      bg2.push(formats[i]);
      
      var fc=fonts[i][4-1];
      if(fonts[i][4-1]=="#000000"){
        fc="#1155cc"
      }
      colD2.push(fc);
      
      
      var fc=fonts[i][7-1];
      if(fonts[i][7-1]=="#000000"){
        fc="#1155cc"
      }
      colG2.push(fc);
      
      var fc=fonts[i][8-1];
      if(fonts[i][8-1]=="#000000"){
        fc="#1155cc"
      }
      colH2.push(fc);
      
    }
    
    
    
    
  }
  
  if(arr1.length>1)
  {
    var lr1=sheet1.getLastRow();
    sheet1.getRange(lr1+1,1, arr1.length, arr1[0].length).setValues(arr1);
    sheet1.getRange(lr1+1,1, arr1.length, arr1[0].length).setBackgrounds(bg1);
    sheet1.getRange(lr1+1, 3,arr1.length, 6).setBorder(true, true, true, true, false, true);
    sheet1.getRange(lr1+1, 4, arr1.length, 1).setFontColors(colD1);
    sheet1.getRange(lr1+1, 7, arr1.length, 1).setFontColors(colG1);
    sheet1.getRange(lr1+1, 8, arr1.length, 1).setFontColors(colH1);
    
  }
  
  if( arr2.length>1)
  {
    var lr2=sheet2.getLastRow();
    sheet2.getRange(lr2+1,1, arr2.length, arr2[0].length).setValues(arr2);
    sheet2.getRange(lr2+1,1, arr2.length, arr2[0].length).setBackgrounds(bg2);
    sheet2.getRange(lr2+1, 3,arr2.length, 6).setBorder(true, true, true, true, false, true);
    sheet2.getRange(lr2+1, 4, arr2.length, 1).setFontColors(colD2);
    sheet2.getRange(lr2+1, 7, arr2.length, 1).setFontColors(colG2);         
    sheet2.getRange(lr2+1, 8, arr2.length, 1).setFontColors(colH2);
  }
  
  sheet.getRange("A8:AN"+lr).clearContent()
  var format=ss.getSheetByName("Template").getRange("A5:AN5").getBackgrounds();
  
  if (values.length>=1)
  {
    var formats2=[];
    
    for (var i=1; i<=values.length; i++)
    {
      formats2.push(format[0]);
    }
    sheet.getRange(8, 1, formats.length, formats[0].length).setBackgrounds(formats2);
    sheet.getRange(8, 1, formats.length, formats[0].length).setBorder(false, false, false, false, false, false);
    
  }
  
  
  
  
}






function transferFromPendingShipper()
{
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName("Pending")
  var lr=sheet.getLastRow();
  var values=sheet.getRange("A5:AN"+lr).getValues();
  var formats=sheet.getRange("A5:AN"+lr).getBackgrounds();
  var formulas=sheet.getRange("A5:AN"+lr).getFormulasR1C1()
  
  
  
  var a=10
  
  var currentDate=new Date();
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  var sheet1=ss.getSheetByName(sheetName);
  var month1=currentDate.getMonth();
  
  
  var firstDayPrevMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 15);
  sheetName=Utilities.formatDate(firstDayPrevMonth, ss.getSpreadsheetTimeZone(), "MMMM-yyyy");
  var sheet2=ss.getSheetByName(sheetName);
  
  var arr1=[];
  var bg1=[];
  var arr2=[];
  var bg2=[];
  
  for (var i=0; i<values.length; i++)
  {
    for (var j=0; j<values[i].length; j++)
    {
      if(formulas[i][j]!=""){values[i][j]=formulas[i][j]};
      
    }
    if(values[i][0]==""){continue}
    var monthTemp=values[i][0].getMonth(); //current month
    if(monthTemp==month1)
    {
      arr1.push(values[i]);
      bg1.push(formats[i]);
      
    }
    
    else
    {
      arr2.push(values[i]);
      bg2.push(formats[i]);
      
    }
    
    
    
    
  }
  
  if(arr1.length>1)
  {
    var lr1=sheet1.getLastRow();
    sheet1.getRange(lr1+1,1, arr1.length, arr1[0].length).setValues(arr1);
    sheet1.getRange(lr1+1,1, arr1.length, arr1[0].length).setBackgrounds(bg1);
    sheet1.getRange(lr1+1, 3,arr1.length, 6).setBorder(true, true, true, true, false, true);
  }
  
  if( arr2.length>1)
  {
    var lr2=sheet2.getLastRow();
    sheet2.getRange(lr2+1,1, arr2.length, arr2[0].length).setValues(arr2);
    sheet2.getRange(lr2+1,1, arr2.length, arr2[0].length).setBackgrounds(bg2);
    sheet2.getRange(lr2+1, 3,arr2.length, 6).setBorder(true, true, true, true, false, true);
  }
  
  sheet.getRange("A5:AN"+lr).clearContent()
  var format=ss.getSheetByName("Template").getRange("A5:AN5").getBackgrounds();
  
  if (values.length>=1)
  {
    var formats2=[];
    
    for (var i=1; i<=values.length; i++)
    {
      formats2.push(format[0]);
    }
    sheet.getRange(5, 1, formats.length, formats[0].length).setBackgrounds(formats2);
    sheet.getRange(5, 1, formats.length, formats[0].length).setBorder(false, false, false, false, false, false);
    
  }
  
  
  
  
}




















function fixConditionalFormatting()
{
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var currentDate=new Date();
  
  
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  var sheet=ss.getSheetByName(sheetName);  
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("AM4:AM");
  var rules = [];
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberBetween(1, 19)
  .setBackground("#f4c7c3")
  .setRanges([range])
  .build();
  rules.push(rule);
  
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberGreaterThanOrEqualTo(20)
  .setBackground("white").setFontColor("white")
  .setRanges([range])
  .build();    
  rules.push(rule);
  
  var range = sheet.getRange("N4:O");
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberGreaterThan(0)
  .setBold(true)
  .setFontColor("#008000")
  .setRanges([range])
  .build();
  rules.push(rule);
  
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberLessThan(0)
  .setBold(true)
  .setFontColor("#c53929")
  .setRanges([range])
  .build();
  rules.push(rule);
  
  var range = sheet.getRange("C4:C");
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberGreaterThan(1)
  .setBackground("#6d9eeb").setUnderline(true).setFontColor("black")
  .setRanges([range])
  .build(); 
  rules.push(rule);
  
  
  var range = sheet.getRange("AN4:AN");     
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenTextEqualTo("PO")
  .setBackground("#f4c7c3")
  .setRanges([range])
  .build();
  rules.push(rule);
  
  
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenTextEqualTo("FR")
  .setBackground("#f4c7c3")
  .setRanges([range])
  .build();
  rules.push(rule);
  
  var range = sheet.getRange("L4:L");     
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied("=if(isnumber(K4),(L4)>30)")
  .setBackground("red")
  .setFontColor("white")
  .setRanges([range])
  .build();
  rules.push(rule);
  
  var range = sheet.getRange("L4:L");     
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberBetween(-40,-5000)
  .setBackground("red")
  .setFontColor("white")
  .setRanges([range])
  .build();
  rules.push(rule);
  
  var range = sheet.getRange("K4:K");     
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberBetween(0.01,3)
  .setBackground("red")
  .setFontColor("white")
  .setRanges([range])
  .build();
  rules.push(rule);     
  
  
  var range = sheet.getRange("M5:M");     
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('=if(AND(len(M5)<9,len(M5)>0),(ISERROR(SEARCH("sportsmansguide",S5))))')
  .setBackground("red")
  .setFontColor("white")
  .setRanges([range])
  .build();
  rules.push(rule);   
  
  var range = sheet.getRange("M5:M");     
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('=exact(M5,E5)')
  .setBackground("red")
  .setFontColor("white")
  .setRanges([range])
  .build();
  rules.push(rule);   
  
  
  /*     
  var range = sheet.getRange("Z4:Z");     
  var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenTextContains("AEG")
  .setBackground("#cccccc")
  .setRanges([range])
  .build();
  rules.push(rule);     
  */   
  
  
  sheet.setConditionalFormatRules(rules);
  
}



















function qtyloss2()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();//e.range;
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  
  
  var sold = sheet.getRange(row, 9).getValue();
  sold=sold*2;
  
  var fees = sheet.getRange(row, 10).getValue();
  fees=fees*2;
  
  var qty = sheet.getRange(row, 3).getValue();
  qty=qty*2;
  
  var cog = sheet.getRange(row, 11).getValue();
  cog=cog*2;
  
  
  sheet.getRange(row, 9).setValue(sold).setFontColor("orange");
  sheet.getRange(row, 10).setValue(fees).setFontColor("orange");
  sheet.getRange(row, 11).setValue(cog).setFontColor("orange");
  sheet.getRange(row, 3).setValue(qty).setFontColor("orange").setBackground("grey");
  
  
  sheet.getRange(row, 17).setValue("ADJ/QTY");
  
  
  
  
}      


function qtyloss()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();//e.range;
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  
  
  var sold = sheet.getRange(row, 9).getValue();
  sold=sold/2;
  
  var fees = sheet.getRange(row, 10).getValue();
  fees=fees/2;
  
  var qty = sheet.getRange(row, 3).getValue();
  qty=qty/2;
  
  var cog = sheet.getRange(row, 11).getValue();
  cog=cog/2;
  
  
  sheet.getRange(row, 9).setValue(sold).setFontColor("orange");
  sheet.getRange(row, 10).setValue(fees).setFontColor("orange");
  sheet.getRange(row, 11).setValue(cog).setFontColor("orange");
  sheet.getRange(row, 3).setValue(qty).setFontColor("orange").setBackground("grey");
  
  
  sheet.getRange(row, 17).setValue("ADJ/QTY");
  
  
  
  
}      


function zeroout() {
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheetName=sheet.getName();
  
  sheet.getRange(row, 3).setValue("0").setFontColor("orange");
  sheet.getRange(row, 9).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 10).setValue("0.00").setFontColor("orange");
  
  sheet.getRange(row, 11).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 12).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 26).setValue("ISSUE").setFontColor("orange");
  sheet.getRange(row, 27).clearContent();
}


function zerooutcog() {
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheetName=sheet.getName();
  
  sheet.getRange(row, 3).setValue("0").setFontColor("orange");
  sheet.getRange(row, 9).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 10).setValue("0.00").setFontColor("orange");
  
  
  sheet.getRange(row, 12).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 26).setValue("ISSUE").setFontColor("orange");
  sheet.getRange(row, 27).clearContent();
}



function zerooutonlycog() {
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheetName=sheet.getName();
  
  
  sheet.getRange(row, 11).setValue("0.00").setFontColor("orange");
  
  
  sheet.getRange(row, 26).setValue("ISSUE").setFontColor("orange");
  sheet.getRange(row, 27).clearContent();
}






function newupdates(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  
  var h1 = '<html>';
  
  var uploadamz ='https://drive.google.com/drive/u/0/folders/0B7gkCg_86yMwX2VlN3RGZUZuWUk';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+uploadamz+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );
  
}

function openchat(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  
  var h1 = '<html>';
  
  var uploadamz ='https://sellercentral.amazon.com/messaging/inbox/ref=ag_cmin_wper_home?cs=-1406852990&ct=2596819448769383330&fi=RESPONSE_NEEDED&pn=1';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+uploadamz+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );
  
}



function pauseorders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  
  var currentDate=new Date();
  
  var timeStamp = Utilities.formatDate(currentDate, "America/Toronto", "HH:mm")
  
  
  
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  var sheet2=ss.getSheetByName(sheetName);
  
  
  
  
  
  sheet2.getRange("C2").setValue("Time: "+timeStamp+" - PAUSE Overstock Orders - Brodie - Continue Once MSG Deleted");
}


function pauseorderswm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  
  var currentDate=new Date();
  
  var timeStamp = Utilities.formatDate(currentDate, "America/Toronto", "HH:mm")
  
  
  
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  var sheet2=ss.getSheetByName(sheetName);
  
  
  
  
  
  sheet2.getRange("C2").setValue("Time: "+timeStamp+" - PAUSE Walmart Orders - Brodie - Continue Once MSG Deleted");
}


function pauseordersae() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  
  var currentDate=new Date();
  
  var timeStamp = Utilities.formatDate(currentDate, "America/Toronto", "HH:mm")
  
  
  
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  var sheet2=ss.getSheetByName(sheetName);
  
  
  
  
  
  sheet2.getRange("C2").setValue("Time: "+timeStamp+" - PAUSE Aliexpress Orders - Brodie - Continue Once MSG Deleted");
}


function pauseorderssg() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  
  var currentDate=new Date();
  
  var timeStamp = Utilities.formatDate(currentDate, "America/Toronto", "HH:mm")
  
  
  
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  var sheet2=ss.getSheetByName(sheetName);
  
  
  
  
  
  sheet2.getRange("C2").setValue("Time: "+timeStamp+" - PAUSE Sportsman Orders - Brodie - Continue Once MSG Deleted");
}


function pauseorderswf() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  
  var currentDate=new Date();
  
  var timeStamp = Utilities.formatDate(currentDate, "America/Toronto", "HH:mm")
  
  
  
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  var sheet2=ss.getSheetByName(sheetName);
  
  
  
  
  
  sheet2.getRange("C2").setValue("Time: "+timeStamp+" - PAUSE Wayfair Orders - Brodie - Continue Once MSG Deleted");
}



















function replaceAll(string, find, replace) {
  return string.replace(new RegExp(escapeRegExp(find), 'g'), replace);
}

function escapeRegExp(string) {
  return string.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}




function newcoupon(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  
  var h1 = '<html>';
  
  
  var couponButton=retrieveCouponUrl2();
  
  
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+couponButton+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );
  
}



function retrieveCouponUrl2()
{
  var ss=SpreadsheetApp.openById(couponSsId);
  var sheet=ss.getSheetByName("TEMPLATE");
  var lr=last_row(sheet,1);
  
  if(lr>3)
  {    
    var cUrl=sheet.getRange(lr,1).getValue();
    
    
    sheet.getRange(lr,2).setValue(sheet.getRange(lr,1).getValue());//move the url to column B
    sheet.getRange(lr,1).clearContent();
    
    return cUrl;
  }
  
  
  else if(lr<=3)  //take in 10%
  {
    
    var lr=last_row(sheet,3);
    
    if(lr<=3){return "N/A"};
    var cUrl=sheet.getRange(lr,3).getValue();
    var cFrm='=HYPERLINK("'+cUrl+'","10%")';
    
    sheet.getRange(lr,4).setValue(sheet.getRange(lr,3).getValue());//move the url to column D
    sheet.getRange(lr,3).clearContent();
    
    
    
    return cUrl;
    
    
    
    
    
  }
  
  
}






function addcoupon() {
  
  
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheetName=sheet.getName();
  
  
  
  
  
  var backg = sheet.getRange(row, 5).getBackground();
  var lastrow = rng.getLastRow();
  
  for (var i=row; i<=lastrow; i++)
  {
    
    var orderId1 = sheet.getRange(i,4).getValue();
    var couponButton=retrieveCouponUrl(orderId1)
    var producturl = sheet.getRange(i, 19).getValue();
    var OSbutton ='=HYPERLINK("'+producturl+'","OS")';
    
    
    sheet.getRange(i, 7).setValue(OSbutton).setFontColor("#1155cc").setBackground(backg);
    sheet.getRange(i, 8).setValue(couponButton).setFontColor("#1155cc").setBackground(backg);
    
    
    
  }
}



function addsale() {
  
  
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  
  var col=rng.getColumn();
  var sheetName=sheet.getName();
  
  
  
  var lastrow = rng.getLastRow();
  
  for (var i=row; i<=lastrow; i++)
  {
    
    var OSbutton ='=HYPERLINK("https://www.topcashback.com/EarnCashback.aspx?mpurl=overstock&mpID=1004891","OS")'; 
    var backg = sheet.getRange(i, 5).getBackground();
    
    
    
    sheet.getRange(i, 7).setValue(OSbutton).setFontColor("Red").setBackground(backg);
    sheet.getRange(i, 8).clearContent().setBackground(backg);
    
  }
  
  
}








function onEdit2(e) {
  
  //var rng=e.range;
  // Browser.msgBox(u)
  
  //var a=ScriptApp.getProjectTriggers();
  
  //[0].getTriggerSourceId();
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var value=sheet.getRange(row, 17).getValue();
  var sheetName=sheet.getName();
  
  
  
  var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
  
  if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
  
  
  
  
  
  
  
  
  
  if(row>4 && sheetName!="CL")//if1
  {
    
    
    if(col==14 || col==15 || col==17)
    {
      
      if(value=="ClubO Applied")
      {
        var cb=sheet.getRange(row, 14).getValue();
        var os=sheet.getRange(row, 15).getValue();
        
        if(cb==""){cb=0};
        if(os==""){os=0};
        
        var val=cb+os;
        val=val-10;
        if(val<3.19)
        {
          
          sheet.getRange(row, 17).setFontColor("#ff0000")
          
        }
        
      }  
      
      
    }  //end of col if
    
    
    
    //when profilt clumn is edited
    
    
    else if(col==12)
    {
      
      var valueL=sheet.getRange(row, 12).getValue();
      
      
      if(valueL=="OI")
      {
        var ssRes=SpreadsheetApp.openById(resSsId);
        
        var sheetRes=ssRes.getSheetByName("OI");
        if(sheetRes==null)
        {
          var templateSheet=ssRes.getSheetByName("TEMPLATE");
          sheetRes=ssRes.insertSheet(sheetName, {template: templateSheet});
          if(sheetRes.isSheetHidden()==true)
          {
            sheetRes.showSheet();  //makes the sheet visible 
          }
          
        }
        var values=sheet.getRange(row,1,1,26).getValues();
        var date=values[0][0];
        var code=valueL;
        var orderId=sheet.getRange(row,4).getFormula();
        var supId=values[0][13-1];
        if(supId==""){supId="N/A"};
        var supId2=addHypToSupId(supId);
        
        var employeeName=values[0][16-1];
        var lr=sheetRes.getLastRow();
        var msg='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+values[0][4-1]+'","MSG")';
        
        var arr=[date,code,"PENDING",orderId,supId2,employeeName,msg]
        sheetRes.getRange(lr+1,1,1,7).setValues([arr]);
        var trackingDetails=trackingCalc(sheet.getRange(row,13).getValue().toString());
        var trackingNo=trackingDetails[0];
        if(trackingNo!=null)
        {sheetRes.getRange(lr+1, 12).setFormula(trackingNo)}
        
        var rtrns=values[0]//sheet.getRange(row, 1,1,sheet.getLastColumn()).getValues();
        rtrns[7-1]=sheet.getRange(row, 7).getFormula();
        var vals=[rtrns[2-1], rtrns[1-1],   rtrns[14-1+3]+rtrns[15-1],  rtrns[5-1],  rtrns[6-1],   rtrns[7-1], rtrns[19-1],  rtrns[25-1],  rtrns[24-1], rtrns[16-1],  rtrns[11-1],  rtrns[26-1]]  
        sheetRes.getRange(lr+1, 13,1,vals.length).setValues([vals]);
        
        
        
        
      }
      
      else if(valueL=="OOS" || valueL=="PO" || valueL=="FR" || valueL=="LP")
      {
        var ssRes=SpreadsheetApp.openById(resSsId);
        var sheetRes=ssRes.getSheetByName("OOS");
        
        var values=sheet.getRange(row,1,1,27).getValues();
        var date=values[0][0];
        var code=valueL;
        var orderId=sheet.getRange(row,4).getFormula();
        var supId=values[0][13-1];
        if(supId==""){supId="N/A"};
        var supId2=addHypToSupId(supId);
        var employeeName=values[0][16-1];
        var lr=sheetRes.getLastRow();
        var msg='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+values[0][4-1]+'","MSG")';
        var variation=values[0][6-1];
        var source1=values[0][19-1];
        var itemNo=values[0][5-1];
        var shipBy=values[0][27-1];
        var comments="";
        var productName=values[0][2-1];
        var soldPrice=values[0][9-1];
        var ASIN=values[0][24-1];
        
        var arr=[date,code,"PENDING",orderId,supId2,employeeName,msg,"","",shipBy,"",variation, source1, itemNo,comments,productName,soldPrice,"","","","",ASIN];
        sheetRes.getRange(lr+1,1,1,17).setValues([arr]);
        
        
        //trasnfer to Tracking feed
        
        //commented out for now///////////
        
        /* 
        var supId=sheet.getRange(row, 13).getValue();
        
        var sheetTrackingFeed=ssTracking.getSheetByName("Tracking Feed");
        var lrTF=sheetTrackingFeed.getLastRow();
        
        
        if(supId==""){return 0};
        var values=sheet.getRange(row, 1,1,27).getValues();
        var colA=values[0][0]; //date
        var colB="PENDING"; 
        var colC=hypOrderId(values[0][4-1]); //hyperlinked order id
        var colD=supId;
        var trackingDetails=trackingCalc(sheet.getRange(row,13).getValue().toString());
        var trackingNo=trackingDetails[0];
        var colE=trackingNo;
        var colF=values[0][27-1];
        var colG="";
        sheetTrackingFeed.getRange(lrTF+1, 1,1,7).setValues([[colA,colB,colC,colD,colE,colF,colG]]);
        */
        
        
        
        
        
      }
      
      
      
      
      
      
    }//end of column is 12
    
    
    
    else if(col==13)
    {
      
      var issuecode = sheet.getRange(row, 12).getValue();
      var markupform ='=R[0]C[-3]-R[0]C[-2]-R[0]C[-1]';
      
      
      if(issuecode == "OOS" | issuecode == "PO" | issuecode == "FR"){
        
        sheet.getRange(row, 12).setValue(markupform);
        sheet.getRange(row, 12).setBackground("#E9E9E9");
        sheet.getRange(row, 12).setFontColor("#197319");
        
        sheet.getRange(row, 16).setValue("Rianne");
        
        
        
        
        
        var orderId=sheet.getRange(row, 4).getValue();
        
        
        var sheet6=SpreadsheetApp.openById(resSsId);
        var sheetOOS=sheet6.getSheetByName("OOS");
        
        var rowO=lookup(orderId,sheetOOS,4, 6,"row");
        
        
        
        sheetOOS.getRange(rowO, 3).setValue("O-SHIPPED");
        sheetOOS.getRange(rowO, 11).setValue("Shipped Order");
        
      }
      
      
      
      
      
      
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
  }//end of if1
  
  
  var flag=0;
  
  if(row>=11 && sheetName=="CL" && col==1)//if
  {
    
    
    var asin=sheet.getRange(row, 1).getValue();
    var ssInv=SpreadsheetApp.openById(invId)
    var sheet2=ssInv.getSheetByName("Inventory List");
    
    
    
    if(asin==""){return 0}
    
    var currentDate=new Date();
    
    
    
    var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
    var sheet2=ss.getSheetByName(sheetName);
    var rowM=lookup(asin,sheet2,25, 27,"row");
    
    
    if(rowM!=null)
    {                                        
      
      
      var sku=sheet2.getRange(rowM, 25).getValue();
      var variation=sheet2.getRange(rowM, 6).getValue();
      var src1=sheet2.getRange(rowM, 19).getValue();
      
      var sheetCA=ss.getSheetByName("CL");
      
      var rowT=lookup(asin,sheetCA,1, 3,"row");
      if(rowT!=row){return 0};
      
      
      sheetCA.getRange(row,1,1,3).setValues([[sku,variation,src1]]);
    }
    return 0;         
  }
  
  
  
  if(row>4 && col==17)
  {
    
    var values=sheet.getRange(row, 1,1,26).getValues();
    var values1=values[0];
    var initials=values[0][26-1];
    var comment=values[0][17-1];
    
    
    if(comment=="Blue-Corey" || comment=="Orange-Jeremy" ||comment=="Green-Brad" || comment=="Red-Gage")
    {
      
      if(comment=="Blue-Corey" && initials.toLowerCase().indexOf("cc")>=0)
      {var ssC=SpreadsheetApp.openById("1J6q9z2JDKHDB5Mj2ntdsv0jTMmCBq8572F0vUaRofNY");}
      else if(comment=="Orange-Jeremy")
      {var ssC=SpreadsheetApp.openById("17oRbfebN25sY3SluQcxIbCdCZoaAtPMlKPoKcjf_Jp8");}
      else if(comment=="Green-Brad")       
      {var ssC=SpreadsheetApp.openById("1Uw6cn-x__Y2kwIIRik64HIiiUuHwA804z4wDsLtJS1U");}
      else if(comment=="Red-Gage")
      {var ssC=SpreadsheetApp.openById("1Cb7llmNr-zsUnTJUB1BOe4z57wJSrKUchsFT-miY4aY");}
      else {return 0}  
      
      var sheetC=ssC.getSheetByName(sheet.getName());
      
      
      if(sheetC==null)
      {
        
        var templateSheet=ssC.getSheetByName("TEMPLATE");
        sheetC=ssC.insertSheet(sheetName, {template: templateSheet});
        if(sheetC.isSheetHidden()==true)
        {
          sheetC.showSheet();  //makes the sheet visible 
          
        }
        
      }
      
      
      var values=sheet.getRange(row, 1,1,26).getValues();
      var valuesF=sheet.getRange(row, 1,1,26).getFormulas();
      
      var date=values[0][1-1];
      var orderId=valuesF[0][4-1];
      var asin=valuesF[0][24-1];
      var markUp=values[0][12-1];
      var title=values[0][2-1];
      
      var lr_comSheet=last_row(sheetC,1);
      sheetC.getRange(lr_comSheet+1, 1,1,4).setValues([[date,orderId, asin, markUp]]);
      sheetC.getRange(lr_comSheet+1, 7).setValue(title); 
      
      
      
      
    }//if blue corey   
    
    
    /*       else if(comment=="Closed ASIN" || comment=="Closed ASIN/ClubO")
    {
    
    
    var sku=sheet.getRange(row, 25).getValue();
    var variation=sheet.getRange(row, 6).getValue();
    var src1=sheet.getRange(row, 19).getValue();
    var status=""//sheet.getRange(row, 19).getValue();
    
    var sheetCA=ss.getSheetByName("CL");
    var rowT=lookup(sku,sheetCA,1, 3,"row");
    if(rowT!=null){return 0}; //stops the repeat
    
    var lrCa=sheetCA.getLastRow();
    var asin=sheet.getRange(row, 24).getValue();
    var prodName=sheet.getRange(row, 2).getValue();
    var colF='=HYPERLINK("https://www.amazon.com/gp/product/'+ asin + '","AMZ AD")';
    var colG='=HYPERLINK("https://catalog.amazon.com/abis/edit/RelistProduct.amzn?marketplaceID=ATVPDKIKX0DER&ref=xx_myirelis_cont_myifba&sku='+sku+'&asin='+asin+'&productType=HOME","RELIST")';
    var colH="";
    var colHTemp=sheet.getRange(row, 12).getValue();;
    var dateclosed=sheet.getRange(row, 1).getValue();
    
    if(colHTemp=="OOS" || colHTemp=="LP" ||colHTemp=="OI")
    {
    colH=colHTemp;
    
    }
    
    
    
    var finalProfit=sheet.getRange(row, 14).getValue();
    if(finalProfit==""){finalProfit=sheet.getRange(row, 15).getValue()}
    
    var soldPrice=sheet.getRange(row, 9).getValue();
    
    sheetCA.getRange(lrCa+1,1,1,17).setValues([["",variation,src1,status,"",colF, colG,colH,asin,prodName,finalProfit,soldPrice,"","",sku,"",dateclosed]]);
    
    }
    
    
    */
    
  }// if col=17
  
  
  
  var instocksheet = ss.getSheetByName("InStock");
  var employeelog = instocksheet.getRange("A1").getValue();
  
  var issuecode = sheet.getRange(row, 12).getValue();
  var issueformula = sheet.getRange(row, 12).getFormula();
  var markupform ='=R[0]C[-3]-R[0]C[-2]-R[0]C[-1]';
  var initialsship = sheet.getRange(row, 16).getValue();
  
  
  
  if(col==13 && issuecode == "OOS" | issuecode == "PO" | issuecode == "FR"){
    
    sheet.getRange(row, 12).setValue(markupform);
    sheet.getRange(row, 12).setBackground("#E9E9E9");
    sheet.getRange(row, 12).setFontColor("#197319");
    sheet.getRange(row, 16).setValue("Rianne");
  }
  
  
  else if(col==13 && initialsship=="")
  {
    
    sheet.getRange(row,16).setValue(employeelog);
    
    
  } 
  
  
  
  
  
  if(col==13)  //when column M or Z is edited, transfer to commision sheets
  {
    
    
    
    
    
    
    
    var supId=sheet.getRange(row, 13).getValue();
    var initial=sheet.getRange(row, 26).getValue();
    
    
    if(supId!="" && initial!="")
    {
      
      if(initial.indexOf("Shams")>=0)
      {
        var commissionSheetId=shamsSheetId;
        
      }
      else if(initial.indexOf("JECU")>=0)
      {
        var commissionSheetId=jeremySheetId;
        
      }
      else if(initial.indexOf("BPCU")>=0)
      {
        var commissionSheetId=bradSheetId;
        
      }
      
      else if(initial.indexOf("SGCU")>=0)
      {
        var commissionSheetId=saraSheetId;
        
      }
      else if(initial.indexOf("GLCU")>=0)
      {
        var commissionSheetId=gageSheetId;
        
      }
      else if(initial.indexOf("MGCU")>=0)
      {
        var commissionSheetId=matthewSheetId;
        
      }
      else if(initial.indexOf("DSCU")>=0)
      {
        var commissionSheetId=daveSheetId;
        
      }
      else if(initial.indexOf("SHCU")>=0)
      {
        var commissionSheetId=steveSheetId;
        
      }
      
      else if(initial.indexOf("TPCU")>=0)
      {
        var commissionSheetId=trevSheetId;
        
      }
      
      else if(initial.indexOf("ZACU")>=0)
      {
        var commissionSheetId=zainabSheetId;
        
      }
      
      else if(initial.indexOf("DGCU")>=0)
      {
        var commissionSheetId=domSheetId;
        
      }
      else if(initial.indexOf("RNCU")>=0)
      {
        var commissionSheetId=reillySheetId;
        
      }
      
      
      
      else if(initial.indexOf("MMCU")>=0)
      {
        var commissionSheetId=mmSheetId;
        
      }
      
      else if(initial.indexOf("RHCU")>=0)
      {
        var commissionSheetId=rohitSheetId;
        
      }
      
      
      
      
      else
      { return 0}
      
      var comSs=SpreadsheetApp.openById(commissionSheetId);
      var comSheet=comSs.getSheetByName(sheet.getName());
      if(comSheet==null){
        
        var comTemplateSheet=comSs.getSheetByName("TEMPLATE");
        comSheet=comSs.insertSheet(sheet.getName(),{template: comTemplateSheet});
        comSheet.showSheet();
        
      }
      
      var values=sheet.getRange(row, 1,1,26).getValues();
      var valuesF=sheet.getRange(row, 1,1,26).getFormulas();
      
      var date=values[0][1-1];
      var orderId=valuesF[0][4-1];
      var asin=valuesF[0][24-1];
      var markUp=values[0][12-1];
      var title=values[0][2-1];
      
      
      var lr_comSheet=last_row(comSheet,1);
      
      lr_comSheet<2?2:lr_comSheet;
      
      comSheet.getRange(lr_comSheet+1, 1,1,4).setValues([[date,orderId, asin, markUp]]);
      comSheet.getRange(lr_comSheet+1, 7).setValue(title);
      
      
      var initials=initial;
      var salePrice=values[0][9-1];
      var sku=valuesF[0][25-1];
      var finalProfit=values[0][14-1]+values[0][15-1];
      var supId=values[0][13-1];
      comSheet.getRange(lr_comSheet+1, 8,1,4).setValues([[initials,salePrice, sku, finalProfit]]);
      comSheet.getRange(lr_comSheet+1, 14).setValue(supId.toString());
      
      
      
      
    }
    
    
    
    
    
    
    
    
    
  }//end of if col is M and Z
  
  
  
  
  
  
}









function lookup(l_value,sheet2,lookup_col, pick_up_col,value_or_row)
{
  
  
  
  var last_row2=sheet2.getLastRow();
  
  if (last_row2<2){last_row2=2}
  
  var ar=sheet2.getRange(2,lookup_col,last_row2-1, pick_up_col-lookup_col+1).getValues()
  
  var flag=0;
  for (var i=0; i<last_row2-1; i++)
  {
    var temp1 = ar[i]
    var temp=temp1[0]
    if(temp==l_value)
    {
      flag=1;
      if (value_or_row=="value")   
      {return ar[i][temp1.length-1]
      break;  }
      
      if (value_or_row=="row")   
      {return i+2
      break;  }
      
      
    }
    
  }
  if(flag==0){return null}
}






function last_row(sheet, col)
{
  //var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  //col=1
  var values=sheet.getRange(1, col,sheet.getLastRow(),1).getValues();
  
  
  for(var i=values.length-1; i>=0; i--)
  {
    if (values[i][0] != "")
    {break}
    
  }
  
  
  return i+1
  
}
























function trackingCalc(id)
{
  var ss=SpreadsheetApp.openById(trackingSheetId);
  
  var currentDate=new Date();
  
  
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy");
  var sheet=ss.getSheetByName(sheetName);
  
  var row=lookup(id,sheet,3, 6,"row");
  
  if(row==null)
  {
    var firstDayPrevMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 15);
    sheetName=Utilities.formatDate(firstDayPrevMonth, ss.getSpreadsheetTimeZone(), "MMMM-yyyy");
    sheet=ss.getSheetByName(sheetName);
    var row=lookup(id,sheet,3, 6,"row");
    
    
    
    
    
  }
  
  if(row==null){return [null,null]}
  
  var trackingNo=sheet.getRange(row, 8).getFormula();
  
  
  
  
  var supId="";  //useless at the moment
  
  return [trackingNo,supId];
  
  
  
  
  
}








function lossCalc(id)
{
  var ss=SpreadsheetApp.openById(orderSheetId)
  
  var currentDate=new Date();
  
  
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy");
  var sheet=ss.getSheetByName(sheetName);
  
  var row=lookup(id,sheet,4, 6,"row");
  
  if(row==null)
  {
    var firstDayPrevMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 15);
    sheetName=Utilities.formatDate(firstDayPrevMonth, ss.getSpreadsheetTimeZone(), "MMMM-yyyy");
    sheet=ss.getSheetByName(sheetName);
    var row=lookup(id,sheet,4, 6,"row");
    
    
    
    
    
  }
  
  if(row==null){return [null,null]}
  
  var profit=sheet.getRange(row, 12).getValue();
  var loss=0-profit;
  
  var supId=sheet.getRange(row, 13).getValue();
  var prodTitle=sheet.getRange(row, 2).getValue();
  
  var values1=[loss,supId,prodTitle]
  var values2=sheet.getRange(row, 1, 1,26).getValues();
  
  var frm=sheet.getRange(row, 7).getFormula();
  values2[0][7-1]=frm;
  
  var v=values1.concat(values2[0]);
  
  
  
  
  return v;
  
  
  
  
  
}









function addHypToSupId(supId)
{       
  
  var supId=supId.toString();
  var n=supId.toString().length
  
  var hSupId=supId;
  if(n==9)
  {
    var baseUrl="https://www.overstock.com/myaccount/#/orders/details/";
    hSupId='=HYPERLINK("'+baseUrl+supId+'","'+supId+'")';
  }
  
  
  if(n==13)
  {
    var baseUrl="https://www.walmart.com/account/order/";
    hSupId='=HYPERLINK("'+baseUrl+supId+'","'+supId+'")';
  }
  
  else if(n==10)
  {
    var baseUrl="https://www.wayfair.com/v/account/order/details?order_id=";
    hSupId='=HYPERLINK("'+baseUrl+supId+'","'+supId+'")';
  }
  
  else if( n==8 && supId.indexOf("559") >= 0 )
  {
    
    hSupId='=HYPERLINK("https://diamondhomeusa.com/='+supId+'","'+supId+'")';
  }
  
  
  else if(n==8 )
  {
    var baseUrl="https://www.sportsmansguide.com/account/accountorderlist";
    hSupId='=HYPERLINK("'+baseUrl+'","'+supId+'")';
  }
  
  
  
  else if(n==19)
  {
    var baseUrl="https://www.amazon.com/gp/your-account/order-details/ref=oh_aui_or_o00_?ie=UTF8&orderID=";
    hSupId='=HYPERLINK("'+baseUrl+supId+'","'+supId+'")';
  }
  
  else if(n==7)
  {
    
    hSupId='=HYPERLINK("https://www.atgstores.com/account/OrderView.aspx?order='+supId+'&customer=1KUd%2fA04IpjFZVhJcN997g%3d%3d","'+supId+'")';
  }
  
  else if(n==14)
  {
    
    hSupId='=HYPERLINK("https://trade.aliexpress.com/order_detail.htm?spm=a2g0s.9042311.0.0.azuaRK&orderId='+supId+'","'+supId+'")';
  }
  
  
  
  return hSupId;
  
  
}




function hypOrderId(newOrderId)
{
  var h='=HYPERLINK("https://sellercentral.amazon.ca/gp/orders-v2/details?ie=UTF8&orderID='+newOrderId+'","'+newOrderId+'")';
  return h;
}




































function grabUrl() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var rng=ss.getActiveRange();
  var sheet=rng.getSheet();
  
  var values=rng.getFormulas();
  
  var sR=rng.getRow();
  var lR=rng.getLastRow();
  
  
  var sC=rng.getColumn();
  var lC=rng.getLastColumn();
  
  var urls=[];
  
  for (var i=0; i<values.length; i++)
  {
    for (var j=0; j<values[0].length; j++)
    {
      var value=values[i][j];
      var n1=value.indexOf('=HYPERLINK(');
      if(n1<0){continue};
      
      var n2=value.indexOf('\"',n1+12);
      var url=value.slice(n1+12,n2);
      urls.push(url);
      
      
      
    }
    
    
    
  }
  
  if(urls.length==0)
  {
    Browser.msgBox("No URL found in selected cells!")
    return 0;
    
  }
  
  
  else 
  {
    var urlString=urls.join("<br>");
    var html='<p>'+urlString+'</p>'
    
    var html = HtmlService.createHtmlOutput(html)
    .setTitle('URLs')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    SpreadsheetApp.getUi().showSidebar(html);
    
    
  }
  
  
  
  
}

function cleanup()
{
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName("InStock");
  
  
  sheet.getRange("AA5:AC").clearContent();
  
  
  
}


function cleanup2()
{
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName("Data");
  
  
  sheet.getRange("A2:C").clearContent();
  
  
  
}

function grabUrl() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var rng=ss.getActiveRange();
  var sheet=rng.getSheet();
  
  var values=rng.getFormulas();
  
  var sR=rng.getRow();
  var lR=rng.getLastRow();
  
  
  var sC=rng.getColumn();
  var lC=rng.getLastColumn();
  
  var urls=[];
  
  for (var i=0; i<values.length; i++)
  {
    for (var j=0; j<values[0].length; j++)
    {
      var value=values[i][j];
      var n1=value.indexOf('=HYPERLINK(');
      if(n1<0){continue};
      
      var n2=value.indexOf('\"',n1+12);
      var url=value.slice(n1+12,n2);
      urls.push(url);
      
      
      
    }
    
    
    
  }
  
  if(urls.length==0)
  {
    Browser.msgBox("No URL found in selected cells!")
    return 0;
    
  }
  
  
  else 
  {
    var urlString=urls.join("<br>");
    var html='<p>'+urlString+'</p>'
    
    var html = HtmlService.createHtmlOutput(html)
    .setTitle('URLs')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    SpreadsheetApp.getUi().showSidebar(html);
    
    
  }
  
  
  
  
}

function link(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var data = sheet.getRange(cell.getRow(), 1, 1, 25);
  var col = data.getNumColumns();
  var row = data.getNumRows(); 
  
  var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
    if(column == 4 || column == 7){
      var cell=data.getCell(row,column);
      var url = /"(.*?)"/.exec(cell.getFormulaR1C1());
      if(url != null){
        var data1 = /"(.*?)"/.exec(cell.getFormulaR1C1())[1];
        var h1 = h1.concat('<script>'
                           +'var a = document.createElement("a"); a.href="'+data1+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}' 
                           +'</script>');
      }
    }
  }
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
    if(column == 8){
      var cell=data.getCell(row,column);
      var url = /"(.*?)"/.exec(cell.getFormulaR1C1());
      if(url != null){
        var data1 = /"(.*?)"/.exec(cell.getFormulaR1C1())[1];
        var h1 = h1.concat('<script>'
                           +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(1400);var a = document.createElement("a"); a.href="'+data1+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                           +'</script>');
      }
    }
  }
  
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
    if(column == 24){
      var cell=data.getCell(row,column);
      var url = /"(.*?)"/.exec(cell.getFormulaR1C1());
      if(url != null){
        var data1 = /"(.*?)"/.exec(cell.getFormulaR1C1())[1];
        var h1 = h1.concat('<script>'
                           +'var a = document.createElement("a"); a.href="'+data1+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}' 
                           +'</script>');
      }
    }
  }
  /*
  var joke ='https://drive.google.com/open?id=1In7dEgVh93rnzvgkH1KkHW9C-mfufW3p';
  
  
  var h1 = h1.concat('<script>'
  +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+joke+'"; a.target="_blank";'
  +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
  +'</script>');                    
  
  
  
  */
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );
  
}









function link2(){
  var selection=SpreadsheetApp.getActiveSheet().getActiveRange()
  var columns=selection.getNumColumns();
  var rows=selection.getNumRows();
  var h1 = '<html>';
  for (var column=1; column < columns; column++) {
    
    if(column == 4 || column == 7 || column == 8){
      //Browser.msgBox(column);
      var cell=selection.getCell(rows,column);
      var url = /"(.*?)"/.exec(cell.getFormulaR1C1());
      if(url != null){
        var data = /"(.*?)"/.exec(cell.getFormulaR1C1())[1];
        var h1 = h1.concat('<script>'
                           +'var a = document.createElement("a"); a.href="'+data+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                           +'google.script.host.close();'
                           +'</script>');
      }
    }
  }
  var h1 = h1.concat('<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  
  var html = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
}



function link3(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var data = sheet.getRange(cell.getRow(), 1, 1, 9);
  var col = data.getNumColumns();
  var row = data.getNumRows(); 
  
  var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
  var amazon ='https://sellercentral.amazon.ca/gp/homepage.html';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+amazon+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  var walmart ='https://www.walmart.com/cp/wedding-registry/1229486';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+walmart+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  var overstock ='https://www.overstock.com/myaccount/#/orders';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+overstock+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  var aliexpress ='https://trade.aliexpress.com/order_list.htm?spm=a2g0s.9042311.0.0.8rmKk7';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+aliexpress+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');                      
  
  var wayfair ='https://www.wayfair.com/session/secure/account/order_search.php';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+wayfair+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');                      
  
  
  var sportsman ='https://www.sportsmansguide.com/account/accountorderlist';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+sportsman+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');           
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );
  
}



function newsheet()
{
  
  
  var ssRes=SpreadsheetApp.openById("1CFFhqYpKu3_2WdmNMkrjZt4u7MmowxJRTpkB-S00ByA");
  
  
  var currentDate =new Date();
  
  
  var sheetRes=ssRes.getSheetByName(Utilities.formatDate(currentDate, ssRes.getSpreadsheetTimeZone(), "MMMM-YYYY"))
  
  
  var  sheetname = Utilities.formatDate(currentDate, ssRes.getSpreadsheetTimeZone(), "MMMM-YYYY");
  
  if(sheetRes==null)
  {
    var templateSheet=ssRes.getSheetByName("TEMPLATE");
    sheetRes=ssRes.insertSheet(sheetname, {template: templateSheet});
    if(sheetRes.isSheetHidden()==true)
    {
      sheetRes.showSheet();  //makes the sheet visible 
    }
    
  }
}


function colorblank()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var secondmonth = new Date(
    new Date().getFullYear(),
    new Date().getMonth() - 1, 
    new Date().getDate()
  );
  
  var second = Utilities.formatDate(secondmonth, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  
  var first = ss.getSheetByName(second);
  first.setTabColor(null); // Set the color to red.
  
  
}
