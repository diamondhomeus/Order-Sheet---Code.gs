

function onEdit4(e) {

  var ss=SpreadsheetApp.getActiveSpreadsheet();


    var rng=ss.getActiveSheet().getActiveRange();
    var sheet=rng.getSheet();
    var row=rng.getRow();
    var col=rng.getColumn();
    var value=sheet.getRange(row, 17).getValue();
    var sheetName=sheet.getName();
    

if(col==5)
{

linkopen();


}














function linkopen(){




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
                           +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(1200);var a = document.createElement("a"); a.href="'+data1+'"; a.target="_blank";'
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



  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );
  
}
}
