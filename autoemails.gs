
function reportissue() {
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();//e.range;
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  
  
    var emailAddress ="brodskies.store@gmail.com";


var ss = SpreadsheetApp.getActiveSpreadsheet();
var date = new Date();

var today = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "MMMM dd YYYY");

  var name = Browser.inputBox("Employee Name");
  
  
       
       
 var h1 = '<html>';



var messagetemplate ='https://mail.google.com/mail/u/0/?view=cm&fs=1&to='+emailAddress+'&tf=1&su=Diamond%20Home%20Scripts%20:%20%20'+today+' - '+name+'&body=%0a%0aDate%20:%20'+today+',%0a%0aXXXAPP/SPREADSHEETXXX%0a%0aXXXDESCRIPTIONXXX%0a%0aProvide%20full%20description%20and%20images%20if%20applicable%0a%0a'+name+'%0aDiamond%20Home';


         var h1 = h1.concat('<script>'
                           +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+messagetemplate+'"; a.target="_blank";'
                           +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                           +'</script>');
                           

         
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );
  



  
  
  
}