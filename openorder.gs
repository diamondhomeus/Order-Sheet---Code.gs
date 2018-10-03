function opensupplierorder() {
  
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();//e.range;
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  
  var supplierID = sheet.getRange(row, 13).getValue();
  
  var supplierIDlen = sheet.getRange(row, 13).getValue();
      supplierIDlen = supplierIDlen.toString().length

  
  
  if(supplierIDlen == 9 )
  {
    var supplier = "https://www.overstock.com/myaccount/#/orders/details/"+supplierID;
  }
  
  
  else if(supplierIDlen == 8 )
  {
    var supplier = "https://www.sportsmansguide.com/account/accountorderlist";
  }
  
  else if(supplierIDlen == 10 )
  {
    var supplier = "https://www.overstock.com/myaccount/#/orders/details/"+supplierID;
  }
  
  else if(supplierIDlen == 13 )
  {
    var supplier = "https://www.wayfair.com/session/secure/account/order_search.php";
  }
  else if(supplierIDlen == 14 )
  {
    var supplier = "https://trade.aliexpress.com/order_detail.htm?spm=a2g0s.9042311.0.0.33d44c4dnfCXZn&orderId="+supplierID;
  }
  
  else if(supplierIDlen == 15 )
  {
    var supplier = "https://trade.aliexpress.com/order_detail.htm?spm=a2g0s.9042311.0.0.33d44c4dnfCXZn&orderId="+supplierID;
  }
  
  
  
  
  var h1 = '<html>';
  
  
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+supplier+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );
  
  
}
