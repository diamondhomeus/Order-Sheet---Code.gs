

function oostransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="OOS";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var sheetName=sheet.getName();
    
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
{
  
                                  var ssRes=SpreadsheetApp.openById(resSsId);
                                  var sheetRes=ssRes.getSheetByName("OOS");
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var values=sheet.getRange(row,1,1,27).getValues();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var orderIdtext=sheet.getRange(row,4).getValue();
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
                                  var SKU=values[0][25-1];

                                  
                                  var arr=[date,code,"PENDING",orderId,supId2,employeeName,msg,"","",shipBy,"",variation, source1, itemNo,comments,productName,soldPrice];
                                  sheetRes.getRange(lr+1,1,1,17).setValues([arr]);


 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getActiveSheet();
 var cell = sheet.getActiveCell();
 var data = sheet.getRange(cell.getRow(), 1, 1, 26);
 var col = data.getNumColumns();
 var row = data.getNumRows(); 
  
 var h1 = '<html>';
 for (var column=1; column < col; column++) {
     //Browser.msgBox(column); 
}
         var removestock ='https://catalog.amazon.com/abis/edit/RelistProduct.amzn?marketplaceID=ATVPDKIKX0DER&ref=xx_myirelis_cont_myifba&sku='+SKU+'&asin='+ASIN+'&productType=HOME';


         var h1 = h1.concat('<script>'
                           +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+removestock+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                           +'</script>');


    var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';


         var h1 = h1.concat('<script>'
                           +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                           +'</script>');


     
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );
  
}



                                  
                                  
                                  
                                  
                    
           
}


function lptransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="LP";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var sheetName=sheet.getName();
    
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
{
  
                                  var ssRes=SpreadsheetApp.openById(resSsId);
                                  var sheetRes=ssRes.getSheetByName("OOS");
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var values=sheet.getRange(row,1,1,27).getValues();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var orderIdtext=sheet.getRange(row,4).getValue();
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
                                  var SKU=values[0][25-1];




 var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var data = sheet.getRange(cell.getRow(), 1, 1, 26);
  var col = data.getNumColumns();
  var row = data.getNumRows(); 
  
  var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
      var closetheasin ='https://sellercentral.amazon.ca/hz/inventory/ref=ag_invmgr_tnav_cmin_?tbla_myitable=sort:%7BsortOrder%3ADESCENDING%2CsortedColumnId%3Adate%7D;search:'+SKU+';pagination:1;';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+closetheasin+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');


    var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';


         var h1 = h1.concat('<script>'
                           +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                           +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );




                                  
                                  var arr=[date,code,"PENDING",orderId,supId2,employeeName,msg,"","",shipBy,"",variation, source1, itemNo,comments,productName,soldPrice];
                                  sheetRes.getRange(lr+1,1,1,17).setValues([arr]);



    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();







                                               var sku=sheet.getRange(row, 25).getValue();
                                                 var variation=sheet.getRange(row, 6).getValue();
                                                 var src1=sheet.getRange(row, 19).getValue();
                                                 var status="";
                                                 var closedsku=sheet.getRange(row, 17).setValue("Closed ASIN");
                                                 var sheetCA=ss.getSheetByName("CL");
                                                 var rowT=lookup(sku,sheetCA,1, 3,"row");
                                                    if(rowT!=null){return 0}; //stops the repeat
                                                    
                                                 var lrCa=sheetCA.getLastRow();
                                                 var asin=sheet.getRange(row, 24).getValue();
                                                 var prodName=sheet.getRange(row, 2).getValue();
                                                 var colF='=HYPERLINK("https://www.amazon.com/gp/product/'+ asin + '","AMZ AD")';
                                                 var colG='=HYPERLINK("https://catalog.amazon.com/abis/edit/RelistProduct.amzn?marketplaceID=ATVPDKIKX0DER&ref=xx_myirelis_cont_myifba&sku='+sku+'&asin='+asin+'&productType=HOME","RELIST")';
                                                 var colH="LP";
                                               
                                                 var dateclosed=sheet.getRange(row, 1).getValue();
                           var ms = new Date(dateclosed).getTime() + (5*86400000);
                           var relistdate = Utilities.formatDate(new Date(ms),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),"MM/dd/yyyy");
                                                 
                                   
                                                 
                                                 
                                                 
                                                 var finalProfit=sheet.getRange(row, 14).getValue();
                                                 if(finalProfit==""){finalProfit=sheet.getRange(row, 15).getValue()}
                                                 
                                                 var soldPrice=sheet.getRange(row, 10).getValue();
                                                 
                                                 sheetCA.getRange(lrCa+1,1,1,18).setValues([["",variation,src1,status,"",colF, colG,colH,asin,prodName,finalProfit,soldPrice,"","",sku,"",dateclosed,relistdate]]);


                                  
 
  

                                  
                                  
                    
                    }
}


function potransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="PO";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var sheetName=sheet.getName();
    
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
{
  
                                  var ssRes=SpreadsheetApp.openById(resSsId);
                                  var sheetRes=ssRes.getSheetByName("OOS");
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var values=sheet.getRange(row,1,1,27).getValues();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var orderIdtext=sheet.getRange(row,4).getValue();
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
                                  
                                  var arr=[date,code,"PENDING",orderId,supId2,employeeName,msg,"","",shipBy,"",variation, source1, itemNo,comments,productName,soldPrice];
                                  sheetRes.getRange(lr+1,1,1,17).setValues([arr]);

 var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var data = sheet.getRange(cell.getRow(), 1, 1, 26);
  var col = data.getNumColumns();
  var row = data.getNumRows(); 
  
  var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }



    var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';


         var h1 = h1.concat('<script>'
                           +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                           +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );

                                  
                                  
                                  
                                  
                    
                    }
}


function frtransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="FR";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var sheetName=sheet.getName();

    
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
{
  
                                  var ssRes=SpreadsheetApp.openById(resSsId);
                                  var sheetRes=ssRes.getSheetByName("OOS");
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var values=sheet.getRange(row,1,1,27).getValues();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var orderIdtext=sheet.getRange(row,4).getValue();
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
                                  
                                  var arr=[date,code,"PENDING",orderId,supId2,employeeName,msg,"","",shipBy,"",variation, source1, itemNo,comments,productName,soldPrice];
                                  sheetRes.getRange(lr+1,1,1,17).setValues([arr]);

                                  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var data = sheet.getRange(cell.getRow(), 1, 1, 26);
  var col = data.getNumColumns();
  var row = data.getNumRows(); 
  
  var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
  
  
  
  var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );      
                                  
                                  
                    
                    }
}



function oitransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="OI";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var sheetName=sheet.getName();
    
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
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
                                  var valueL=sheet.getRange(row, 12).getValue();
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
}








function duptransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="OI";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var employee = ss.getSheetByName("InStock").getRange("A1").getValue();
    var name=sheet.getRange(row, 16).setValue(employee);
    var sheetName=sheet.getName();
    var orderIdtext=sheet.getRange(row, 4).getValue();
  
    var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
  
  
  
  var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );      
                                  
  
  
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
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
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var supId=values[0][13-1];
                                  if(supId==""){supId="N/A"};
                            //      var supId2=addHypToSupId(supId);
                                  
                                  var employeeName=values[0][16-1];
                                  var lr=sheetRes.getLastRow();
                                  var msg='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+values[0][4-1]+'","MSG")';
                                  
                                  var arr=[date,code,"CUSTOMER",orderId,"N/A",employee,msg,"Duplicate Order"]
                                  sheetRes.getRange(lr+1,1,1,8).setValues([arr]);
                             
                                  var rtrns=values[0]//sheet.getRange(row, 1,1,sheet.getLastColumn()).getValues();
                                  rtrns[7-1]=sheet.getRange(row, 7).getFormula();
                                  var vals=[rtrns[2-1], rtrns[1-1],   rtrns[14-1+3]+rtrns[15-1],  rtrns[5-1],  rtrns[6-1],   rtrns[7-1], rtrns[19-1],  rtrns[25-1],  rtrns[24-1], rtrns[16-1],  rtrns[11-1],  rtrns[26-1]]  
                                  sheetRes.getRange(lr+1, 13,1,vals.length).setValues([vals]);
  
  var month=Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MMM");
  month.toString();
  month = month.slice(0,-2);
  
  var futuredate=new Date();       
  futuredate.setDate(futuredate.getDate() + 1);
  
  var dayof=Utilities.formatDate(futuredate, ss.getSpreadsheetTimeZone(), "dd");
  
  sheetRes.getRange(lr+1, 25).setValue(month+dayof);
                                  
                    
                    }
}



function shippingcostserrortransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="OI";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var employee = ss.getSheetByName("InStock").getRange("A1").getValue();
    var name=sheet.getRange(row, 16).setValue(employee);
    var sheetName=sheet.getName();
    var orderIdtext=sheet.getRange(row, 4).getValue();
  
    var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
  
  
  
  var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );      
                                  
  
  
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
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
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var supId=values[0][13-1];
                                  if(supId==""){supId="N/A"};
                            //      var supId2=addHypToSupId(supId);
                                  
                                  var employeeName=values[0][16-1];
                                  var lr=sheetRes.getLastRow();
                                  var msg='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+values[0][4-1]+'","MSG")';
                                  
                                  var arr=[date,code,"CUSTOMER",orderId,"N/A",employee,msg,"1000 Shipping",""]
                                  sheetRes.getRange(lr+1,1,1,9).setValues([arr]);
                             
                                  var rtrns=values[0]//sheet.getRange(row, 1,1,sheet.getLastColumn()).getValues();
                                  rtrns[7-1]=sheet.getRange(row, 7).getFormula();
                                  var vals=[rtrns[2-1], rtrns[1-1],   rtrns[14-1+3]+rtrns[15-1],  rtrns[5-1],  rtrns[6-1],   rtrns[7-1], rtrns[19-1],  rtrns[25-1],  rtrns[24-1], rtrns[16-1],  rtrns[11-1],  rtrns[26-1]]  
                                  sheetRes.getRange(lr+1, 13,1,vals.length).setValues([vals]);
  
  var month=Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MMM");
  month.toString();
  month = month.slice(0,-2);
  
  var futuredate=new Date();       
  futuredate.setDate(futuredate.getDate() + 1);
  
  var dayof=Utilities.formatDate(futuredate, ss.getSpreadsheetTimeZone(), "dd");
  
  sheetRes.getRange(lr+1, 25).setValue(month+dayof);
                                  
                    
                    }
}




function invalidtransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="OI";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var employee = ss.getSheetByName("InStock").getRange("A1").getValue();
    var name=sheet.getRange(row, 16).setValue(employee);
    var sheetName=sheet.getName();
    var orderIdtext=sheet.getRange(row, 4).getValue();
  
    var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
  
  
  
  var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );      
                                  
  
  
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
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
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var supId=values[0][13-1];
                                  if(supId==""){supId="N/A"};
                            //      var supId2=addHypToSupId(supId);
                                  
                                  var employeeName=values[0][16-1];
                                  var lr=sheetRes.getLastRow();
                                  var msg='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+values[0][4-1]+'","MSG")';
                                  
                                  var arr=[date,code,"CUSTOMER",orderId,"N/A",employee,msg,"Invalid Address"]
                                  sheetRes.getRange(lr+1,1,1,8).setValues([arr]);
                             
                                  var rtrns=values[0]//sheet.getRange(row, 1,1,sheet.getLastColumn()).getValues();
                                  rtrns[7-1]=sheet.getRange(row, 7).getFormula();
                                  var vals=[rtrns[2-1], rtrns[1-1],   rtrns[14-1+3]+rtrns[15-1],  rtrns[5-1],  rtrns[6-1],   rtrns[7-1], rtrns[19-1],  rtrns[25-1],  rtrns[24-1], rtrns[16-1],  rtrns[11-1],  rtrns[26-1]]  
                                  sheetRes.getRange(lr+1, 13,1,vals.length).setValues([vals]);
  
  var month=Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MMM");
  month.toString();
  month = month.slice(0,-2);
  
  var futuredate=new Date();       
  futuredate.setDate(futuredate.getDate() + 1);
  
  var dayof=Utilities.formatDate(futuredate, ss.getSpreadsheetTimeZone(), "dd");
  
  sheetRes.getRange(lr+1, 25).setValue(month+dayof);
                                  
                    
                    }
}





function closedasin()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();


                                                    var sku=sheet.getRange(row, 25).getValue();
                                                 var variation=sheet.getRange(row, 6).getValue();
                                                 var src1=sheet.getRange(row, 19).getValue();
                                                 var status="";
                                                 var closedsku=sheet.getRange(row, 17).setValue("Closed ASIN");
                                                 var sheetCA=ss.getSheetByName("CL");
                                                 var rowT=lookup(sku,sheetCA,1, 3,"row");
                                                    if(rowT!=null){return 0}; //stops the repeat
                                                    
                                                 var lrCa=sheetCA.getLastRow();
                                                 var asin=sheet.getRange(row, 24).getValue();
                                                 var prodName=sheet.getRange(row, 2).getValue();
                                                 var colF='=HYPERLINK("https://www.amazon.com/gp/product/'+ asin + '","AMZ AD")';
                                                 var colG='=HYPERLINK("https://catalog.amazon.com/abis/edit/RelistProduct.amzn?marketplaceID=ATVPDKIKX0DER&ref=xx_myirelis_cont_myifba&sku='+sku+'&asin='+asin+'&productType=HOME","RELIST")';
                                                 var colH="LP";
                                               
                                                 var dateclosed=sheet.getRange(row, 1).getValue();
                           var ms = new Date(dateclosed).getTime() + (5*86400000);
                           var relistdate = Utilities.formatDate(new Date(ms),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),"MM/dd/yyyy");
                                                 
                                   
                                                 
                                                 
                                                 
                                                 var finalProfit=sheet.getRange(row, 14).getValue();
                                                 if(finalProfit==""){finalProfit=sheet.getRange(row, 15).getValue()}
                                                 
                                                 var soldPrice=sheet.getRange(row, 10).getValue();
                                                 
                                                 sheetCA.getRange(lrCa+1,1,1,18).setValues([["",variation,src1,status,"",colF, colG,colH,asin,prodName,finalProfit,soldPrice,"","",sku,"",dateclosed,relistdate]]);




                   var values=sheet.getRange(row,1,1,27).getValues();


                                  var SKU=values[0][25-1];



 var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var data = sheet.getRange(cell.getRow(), 1, 1, 26);
  var col = data.getNumColumns();
  var row = data.getNumRows(); 
  
  var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
      var closetheasin ='https://sellercentral.amazon.ca/hz/inventory/ref=ag_invmgr_tnav_cmin_?tbla_myitable=sort:%7BsortOrder%3ADESCENDING%2CsortedColumnId%3Adate%7D;search:'+SKU+';pagination:1;';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+closetheasin+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );


}







function promotransfer()
{
 
    closingasin();


    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="OI";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var sheetName=sheet.getName();
    
    

    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
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
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var supId=values[0][13-1];
                                  if(supId==""){supId="N/A"};
                                  var supId2=addHypToSupId(supId);
                                  
                                  var employeeName=values[0][16-1];
                                  var lr=sheetRes.getLastRow();
                                  var msg='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+values[0][4-1]+'","MSG")';
                                  
                                  var arr=[date,code,"PENDING",orderId,supId2,employeeName,msg,"Promotional"]
                                  sheetRes.getRange(lr+1,1,1,8).setValues([arr]);
                                  var trackingDetails=trackingCalc(sheet.getRange(row,13).getValue().toString());
                                  var trackingNo=trackingDetails[0];
                                   if(trackingNo!=null)
                                      {sheetRes.getRange(lr+1, 12).setFormula(trackingNo)}
                                  
                                  var rtrns=values[0]//sheet.getRange(row, 1,1,sheet.getLastColumn()).getValues();
                                  rtrns[7-1]=sheet.getRange(row, 7).getFormula();
                                  var vals=[rtrns[2-1], rtrns[1-1],   rtrns[14-1+3]+rtrns[15-1],  rtrns[5-1],  rtrns[6-1],   rtrns[7-1], rtrns[19-1],  rtrns[25-1],  rtrns[24-1], rtrns[16-1],  rtrns[11-1],  rtrns[26-1]]  
                                  sheetRes.getRange(lr+1, 13,1,vals.length).setValues([vals]);
                                  
                              
                                  
                    
                    }




                                                 var sku=sheet.getRange(row, 25).getValue();
                                                 var variation=sheet.getRange(row, 6).getValue();
                                                 var src1=sheet.getRange(row, 19).getValue();
                                                 var status="";
                                                 var closedsku=sheet.getRange(row, 17).setValue("Closed ASIN");
                                                 var sheetCA=ss.getSheetByName("CL");
                                                 var rowT=lookup(sku,sheetCA,1, 3,"row");
                                                    if(rowT!=null){return 0}; //stops the repeat
                                                    
                                                 var lrCa=sheetCA.getLastRow();
                                                 var asin=sheet.getRange(row, 24).getValue();
                                                 var prodName=sheet.getRange(row, 2).getValue();
                                                 var colF='=HYPERLINK("https://www.amazon.com/gp/product/'+ asin + '","AMZ AD")';
                                                 var colG='=HYPERLINK("https://catalog.amazon.com/abis/edit/RelistProduct.amzn?marketplaceID=ATVPDKIKX0DER&ref=xx_myirelis_cont_myifba&sku='+sku+'&asin='+asin+'&productType=HOME","RELIST")';
                                                 var colH="PROMO";
                                               
                                                 var dateclosed=sheet.getRange(row, 1).getValue();
                           var ms = new Date(dateclosed).getTime() + (5*86400000);
                           var relistdate = Utilities.formatDate(new Date(ms),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),"MM/dd/yyyy");
                                                 
                                   
                                                 
                                                 
                                                 
                                                 var finalProfit=sheet.getRange(row, 14).getValue();
                                                 if(finalProfit==""){finalProfit=sheet.getRange(row, 15).getValue()}
                                                 
                                                 var soldPrice=sheet.getRange(row, 10).getValue();
                                                 
                                                 sheetCA.getRange(lrCa+1,1,1,18).setValues([["",variation,src1,status,"",colF, colG,colH,asin,prodName,finalProfit,soldPrice,"","",sku,"",dateclosed,relistdate]]);


                 
}





function promotransferoos()
{
 
    closingasin();

   oospromo();


    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
  
    var sheetName=sheet.getName();
    
    

    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
{
  
 
                              
                                  
                    
                    }




                                                 var sku=sheet.getRange(row, 25).getValue();
                                                 var variation=sheet.getRange(row, 6).getValue();
                                                 var src1=sheet.getRange(row, 19).getValue();
                                                 var status="";
                                                 var closedsku=sheet.getRange(row, 17).setValue("Closed ASIN");
                                                 var sheetCA=ss.getSheetByName("CL");
                                                 var rowT=lookup(sku,sheetCA,1, 3,"row");
                                                    if(rowT!=null){return 0}; //stops the repeat
                                                    
                                                 var lrCa=sheetCA.getLastRow();
                                                 var asin=sheet.getRange(row, 24).getValue();
                                                 var prodName=sheet.getRange(row, 2).getValue();
                                                 var colF='=HYPERLINK("https://www.amazon.com/gp/product/'+ asin + '","AMZ AD")';
                                                 var colG='=HYPERLINK("https://catalog.amazon.com/abis/edit/RelistProduct.amzn?marketplaceID=ATVPDKIKX0DER&ref=xx_myirelis_cont_myifba&sku='+sku+'&asin='+asin+'&productType=HOME","RELIST")';
                                                 var colH="PROMO";
                                               
                                                 var dateclosed=sheet.getRange(row, 1).getValue();
                           var ms = new Date(dateclosed).getTime() + (5*86400000);
                           var relistdate = Utilities.formatDate(new Date(ms),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),"MM/dd/yyyy");
                                                 
                                   
                                                 
                                                 
                                                 
                                                 var finalProfit=sheet.getRange(row, 14).getValue();
                                                 if(finalProfit==""){finalProfit=sheet.getRange(row, 15).getValue()}
                                                 
                                                 var soldPrice=sheet.getRange(row, 10).getValue();
                                                 
                                                 sheetCA.getRange(lrCa+1,1,1,18).setValues([["",variation,src1,status,"",colF, colG,colH,asin,prodName,finalProfit,soldPrice,"","",sku,"",dateclosed,relistdate]]);


                 
}


function closingasin ()
{

    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
  
    

  var values=sheet.getRange(row,1,1,27).getValues();


                                  var SKU=values[0][25-1];



 var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var data = sheet.getRange(cell.getRow(), 1, 1, 26);
  var col = data.getNumColumns();
  var row = data.getNumRows(); 
  
  var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
      var closetheasin ='https://sellercentral.amazon.ca/hz/inventory/ref=ag_invmgr_tnav_cmin_?tbla_myitable=sort:%7BsortOrder%3ADESCENDING%2CsortedColumnId%3Adate%7D;search:'+SKU+';pagination:1;';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+closetheasin+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );


}





function oospromo()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="OOS";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var sheetName=sheet.getName();
    
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
{
  
                                  var ssRes=SpreadsheetApp.openById(resSsId);
                                  var sheetRes=ssRes.getSheetByName("OOS");
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var values=sheet.getRange(row,1,1,27).getValues();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var orderIdtext=sheet.getRange(row,4).getValue();
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
                                  var comments="Promotional";
                                  var productName=values[0][2-1];
                                  var soldPrice=values[0][9-1];
                                  var ASIN=values[0][24-1];
                                  var SKU=values[0][25-1];

                                  
                                  var arr=[date,code,"PENDING",orderId,supId2,employeeName,msg,"","",shipBy,"",variation, source1, itemNo,comments,productName,soldPrice];
                                  sheetRes.getRange(lr+1,1,1,17).setValues([arr]);

}










}





function boostransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="OOS";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
        var name=sheet.getRange(row, 16).setValue("Brodie");
    var sheetName=sheet.getName();
    
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
{
  
                                  var ssRes=SpreadsheetApp.openById(resSsId);
                                  var sheetRes=ssRes.getSheetByName("OOS");
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var values=sheet.getRange(row,1,1,27).getValues();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var orderIdtext=sheet.getRange(row,4).getValue();
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
                                  var SKU=values[0][25-1];

                                  
                                  var arr=[date,code,"PENDING",orderId,supId2,employeeName,msg,"1","",shipBy,"",variation, source1, itemNo,comments,productName,soldPrice];
                                  sheetRes.getRange(lr+1,1,1,17).setValues([arr]);


 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getActiveSheet();
 var cell = sheet.getActiveCell();
 var data = sheet.getRange(cell.getRow(), 1, 1, 26);
 var col = data.getNumColumns();
 var row = data.getNumRows(); 
  
 var h1 = '<html>';
 for (var column=1; column < col; column++) {
     //Browser.msgBox(column); 
}
         var removestock ='https://catalog.amazon.com/abis/edit/RelistProduct.amzn?marketplaceID=ATVPDKIKX0DER&ref=xx_myirelis_cont_myifba&sku='+SKU+'&asin='+ASIN+'&productType=HOME';


         var h1 = h1.concat('<script>'
                           +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+removestock+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                           +'</script>');


    var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';


         var h1 = h1.concat('<script>'
                           +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                           +'</script>');


     
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );
  
}
}





function blptransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="LP";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var name=sheet.getRange(row, 16).setValue("Brodie");
    var sheetName=sheet.getName();
    
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
{
  
                                  var ssRes=SpreadsheetApp.openById(resSsId);
                                  var sheetRes=ssRes.getSheetByName("OOS");
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var values=sheet.getRange(row,1,1,27).getValues();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var orderIdtext=sheet.getRange(row,4).getValue();
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
                                  var SKU=values[0][25-1];




 var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var data = sheet.getRange(cell.getRow(), 1, 1, 26);
  var col = data.getNumColumns();
  var row = data.getNumRows(); 
  
  var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
      var closetheasin ='https://sellercentral.amazon.ca/hz/inventory/ref=ag_invmgr_tnav_cmin_?tbla_myitable=sort:%7BsortOrder%3ADESCENDING%2CsortedColumnId%3Adate%7D;search:'+SKU+';pagination:1;';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+closetheasin+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');


    var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';


         var h1 = h1.concat('<script>'
                           +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                           +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );




                                  
                                  var arr=[date,code,"PENDING",orderId,supId2,employeeName,msg,"1","",shipBy,"",variation, source1, itemNo,comments,productName,soldPrice];
                                  sheetRes.getRange(lr+1,1,1,17).setValues([arr]);



    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();







                                               var sku=sheet.getRange(row, 25).getValue();
                                                 var variation=sheet.getRange(row, 6).getValue();
                                                 var src1=sheet.getRange(row, 19).getValue();
                                                 var status="";
                                                 var closedsku=sheet.getRange(row, 17).setValue("Closed ASIN");
                                                 var sheetCA=ss.getSheetByName("CL");
                                                 var rowT=lookup(sku,sheetCA,1, 3,"row");
                                                    if(rowT!=null){return 0}; //stops the repeat
                                                    
                                                 var lrCa=sheetCA.getLastRow();
                                                 var asin=sheet.getRange(row, 24).getValue();
                                                 var prodName=sheet.getRange(row, 2).getValue();
                                                 var colF='=HYPERLINK("https://www.amazon.com/gp/product/'+ asin + '","AMZ AD")';
                                                 var colG='=HYPERLINK("https://catalog.amazon.com/abis/edit/RelistProduct.amzn?marketplaceID=ATVPDKIKX0DER&ref=xx_myirelis_cont_myifba&sku='+sku+'&asin='+asin+'&productType=HOME","RELIST")';
                                                 var colH="LP";
                                               
                                                 var dateclosed=sheet.getRange(row, 1).getValue();
                           var ms = new Date(dateclosed).getTime() + (5*86400000);
                           var relistdate = Utilities.formatDate(new Date(ms),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),"MM/dd/yyyy");
                                                 
                                   
                                                 
                                                 
                                                 
                                                 var finalProfit=sheet.getRange(row, 14).getValue();
                                                 if(finalProfit==""){finalProfit=sheet.getRange(row, 15).getValue()}
                                                 
                                                 var soldPrice=sheet.getRange(row, 10).getValue();
                                                 
                                                 sheetCA.getRange(lrCa+1,1,1,18).setValues([["",variation,src1,status,"",colF, colG,colH,asin,prodName,finalProfit,soldPrice,"","",sku,"",dateclosed,relistdate]]);


                                  
 
  

                                  
                                  
                    
                    }
}


function bpotransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="PO";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var name=sheet.getRange(row, 16).setValue("Brodie");
    var sheetName=sheet.getName();
    
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
{
  
                                  var ssRes=SpreadsheetApp.openById(resSsId);
                                  var sheetRes=ssRes.getSheetByName("OOS");
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var values=sheet.getRange(row,1,1,27).getValues();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var orderIdtext=sheet.getRange(row,4).getValue();
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
                                  
                                  var arr=[date,code,"PENDING",orderId,supId2,employeeName,msg,"1","",shipBy,"",variation, source1, itemNo,comments,productName,soldPrice];
                                  sheetRes.getRange(lr+1,1,1,17).setValues([arr]);

 var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var data = sheet.getRange(cell.getRow(), 1, 1, 26);
  var col = data.getNumColumns();
  var row = data.getNumRows(); 
  
  var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }



    var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';


         var h1 = h1.concat('<script>'
                           +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                           +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );

                                  
                                  
                                  
                                  
                    
                    }
}


function bfrtransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="FR";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var name=sheet.getRange(row, 16).setValue("Brodie");
    var sheetName=sheet.getName();

    
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
{
  
                                  var ssRes=SpreadsheetApp.openById(resSsId);
                                  var sheetRes=ssRes.getSheetByName("OOS");
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var values=sheet.getRange(row,1,1,27).getValues();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var orderIdtext=sheet.getRange(row,4).getValue();
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
                                  
                                  var arr=[date,code,"PENDING",orderId,supId2,employeeName,msg,"1","",shipBy,"",variation, source1, itemNo,comments,productName,soldPrice];
                                  sheetRes.getRange(lr+1,1,1,17).setValues([arr]);

                                  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell = sheet.getActiveCell();
  var data = sheet.getRange(cell.getRow(), 1, 1, 26);
  var col = data.getNumColumns();
  var row = data.getNumRows(); 
  
  var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
  
  
  
  var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );      
                                  
                                  
                    
                    }
}



function boitransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="OI";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var name=sheet.getRange(row, 16).setValue("Brodie");
    var sheetName=sheet.getName();
    
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
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
                                  var valueL=sheet.getRange(row, 12).getValue();
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
}








function bduptransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="OI";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var name=sheet.getRange(row, 16).setValue("Brodie");
    var sheetName=sheet.getName();
    var orderIdtext=sheet.getRange(row, 4).getValue();
  
    var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
  
  
  
  var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );      
                                  
  
  
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
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
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                              
                                  
                                  var employeeName=values[0][16-1];
                                  var lr=sheetRes.getLastRow();
                                  var msg='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+values[0][4-1]+'","MSG")';
                                  
                                  var arr=[date,code,"CUSTOMER",orderId,"N/A","Brodie",msg,"Duplicate Order"]
                                  sheetRes.getRange(lr+1,1,1,8).setValues([arr]);
                                 
                                  
                                  var rtrns=values[0]//sheet.getRange(row, 1,1,sheet.getLastColumn()).getValues();
                                  rtrns[7-1]=sheet.getRange(row, 7).getFormula();
                                  var vals=[rtrns[2-1], rtrns[1-1],   rtrns[14-1+3]+rtrns[15-1],  rtrns[5-1],  rtrns[6-1],   rtrns[7-1], rtrns[19-1],  rtrns[25-1],  rtrns[24-1], rtrns[16-1],  rtrns[11-1],  rtrns[26-1]]  
                                  sheetRes.getRange(lr+1, 13,1,vals.length).setValues([vals]);
  
  var month=Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MMM");
  month.toString();
  month = month.slice(0,-2);
  
  var futuredate=new Date();       
  futuredate.setDate(futuredate.getDate() + 1);
  
  var dayof=Utilities.formatDate(futuredate, ss.getSpreadsheetTimeZone(), "dd");
  
  sheetRes.getRange(lr+1, 25).setValue(month+dayof);
                                  
                    
                    }
}



function bshippingcostserrortransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="OI";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var employee = "Brodie"
    var name=sheet.getRange(row, 16).setValue(employee);
    var sheetName=sheet.getName();
    var orderIdtext=sheet.getRange(row, 4).getValue();
  
    var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
  
  
  
  var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );      
                                  
  
  
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
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
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var supId=values[0][13-1];
                                  if(supId==""){supId="N/A"};
                            //      var supId2=addHypToSupId(supId);
                                  
                                  var employeeName=values[0][16-1];
                                  var lr=sheetRes.getLastRow();
                                  var msg='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+values[0][4-1]+'","MSG")';
                                  
                                  var arr=[date,code,"CUSTOMER",orderId,"N/A","Brodie",msg,"1000 Shipping",""]
                                  sheetRes.getRange(lr+1,1,1,9).setValues([arr]);
                             
                                  var rtrns=values[0]//sheet.getRange(row, 1,1,sheet.getLastColumn()).getValues();
                                  rtrns[7-1]=sheet.getRange(row, 7).getFormula();
                                  var vals=[rtrns[2-1], rtrns[1-1],   rtrns[14-1+3]+rtrns[15-1],  rtrns[5-1],  rtrns[6-1],   rtrns[7-1], rtrns[19-1],  rtrns[25-1],  rtrns[24-1], rtrns[16-1],  rtrns[11-1],  rtrns[26-1]]  
                                  sheetRes.getRange(lr+1, 13,1,vals.length).setValues([vals]);
  
  var month=Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MMM");
  month.toString();
  month = month.slice(0,-2);
  
  var futuredate=new Date();       
  futuredate.setDate(futuredate.getDate() + 1);
  
  var dayof=Utilities.formatDate(futuredate, ss.getSpreadsheetTimeZone(), "dd");
  
  sheetRes.getRange(lr+1, 25).setValue(month+dayof);
                                  
                    
                    }
}




function binvalidtransfer()
{
 
    
    var ss=SpreadsheetApp.getActiveSpreadsheet();

    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var oos="OI";
    var color = "#F4C4C2"
    var font = "#00021B"
    var value=sheet.getRange(row, 12).setValue(oos).setBackground(color).setFontColor(font);
    var employee = "Brodie"
    var name=sheet.getRange(row, 16).setValue(employee);
    var sheetName=sheet.getName();
    var orderIdtext=sheet.getRange(row, 4).getValue();
  
    var h1 = '<html>';
  for (var column=1; column < col; column++) {
    //Browser.msgBox(column); 
  }
  
  
  
  var msgcustomer ='https://sellercentral.amazon.com/gp/help/contact/contact.html?orderID='+orderIdtext+'&marketplaceID=ATVPDKIKX0DER';
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+msgcustomer+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );      
                                  
  
  
    
    //do not trigger from these sheets
    var neglectSheets=["Intermediate leaderboard","Initals Stats","Monthly Stats","Yearly Stats","Leaderboard"];
    
    if(neglectSheets.indexOf(sheet.getName())>=0){return 0};
    
  
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
                                  var valueL=sheet.getRange(row, 12).getValue();
                                  var date=values[0][0];
                                  var code=valueL;
                                  var orderId=sheet.getRange(row,4).getFormula();
                                  var supId=values[0][13-1];
                                  if(supId==""){supId="N/A"};
                            //      var supId2=addHypToSupId(supId);
                                  
                                  var employeeName=values[0][16-1];
                                  var lr=sheetRes.getLastRow();
                                  var msg='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+values[0][4-1]+'","MSG")';
                                  
                                  var arr=[date,code,"CUSTOMER",orderId,"N/A","Brodie",msg,"Invalid Address"]
                                  sheetRes.getRange(lr+1,1,1,8).setValues([arr]);
                             
                                  var rtrns=values[0]//sheet.getRange(row, 1,1,sheet.getLastColumn()).getValues();
                                  rtrns[7-1]=sheet.getRange(row, 7).getFormula();
                                  var vals=[rtrns[2-1], rtrns[1-1],   rtrns[14-1+3]+rtrns[15-1],  rtrns[5-1],  rtrns[6-1],   rtrns[7-1], rtrns[19-1],  rtrns[25-1],  rtrns[24-1], rtrns[16-1],  rtrns[11-1],  rtrns[26-1]]  
                                  sheetRes.getRange(lr+1, 13,1,vals.length).setValues([vals]);
  
  var month=Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MMM");
  month.toString();
  month = month.slice(0,-2);
  
  var futuredate=new Date();       
  futuredate.setDate(futuredate.getDate() + 1);
  
  var dayof=Utilities.formatDate(futuredate, ss.getSpreadsheetTimeZone(), "dd");
  
  sheetRes.getRange(lr+1, 25).setValue(month+dayof);
                                  
                    
                    }
}






