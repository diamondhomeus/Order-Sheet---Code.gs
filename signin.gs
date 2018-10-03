

function Kris() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInv = ss.getSheetByName("InStock");
  sheetInv.getRange("A1").setValue("Kris");
  var activesheet = ss.getActiveSheet();
  var startrow = ss.getActiveRange().getRow();
  var popup = Browser.msgBox('1) Are you on the Month Tab To Ship On? 2) Have your Mouse on the Start Row, Select "Yes" If Correct', 'Otherwise select "No" to Cancel', Browser.Buttons.YES_NO);
  
  if(popup == "yes"){
    activesheet.getRange(2, 24).setValue("Kris");
           activesheet.getRange(2, 25).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")&" ➶ "&left(Z1/280*100,1)&"%"');
    activesheet.getRange(1, 26).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")');
  }
  if(popup == "no"){
    return 0;
  }
}



function Gage() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInv = ss.getSheetByName("InStock");
  sheetInv.getRange("A1").setValue("Gage");
  var activesheet = ss.getActiveSheet();
  var startrow = ss.getActiveRange().getRow();
  var popup = Browser.msgBox('1) Are you on the Month Tab To Ship On? 2) Have your Mouse on the Start Row, Select "Yes" If Correct', 'Otherwise select "No" to Cancel', Browser.Buttons.YES_NO);
  
  if(popup == "yes"){
    activesheet.getRange(2, 24).setValue("Gage");
          activesheet.getRange(2, 25).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")&" ➶ "&left(Z1/280*100,1)&"%"');
    activesheet.getRange(1, 26).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")');
  }
  if(popup == "no"){
    return 0;
  }
}


function Brad() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInv = ss.getSheetByName("InStock");
  sheetInv.getRange("A1").setValue("Brad");
  var activesheet = ss.getActiveSheet();
  var startrow = ss.getActiveRange().getRow();
  var popup = Browser.msgBox('1) Are you on the Month Tab To Ship On? 2) Have your Mouse on the Start Row, Select "Yes" If Correct', 'Otherwise select "No" to Cancel', Browser.Buttons.YES_NO);
  
  if(popup == "yes"){
    activesheet.getRange(2, 24).setValue("Brad");
       activesheet.getRange(2, 25).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")&" ➶ "&left(Z1/280*100,1)&"%"');
    activesheet.getRange(1, 26).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")');
  }
  if(popup == "no"){
    return 0;
  }
}


function Amanda() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInv = ss.getSheetByName("InStock");
  sheetInv.getRange("A1").setValue("Amanda");
  var activesheet = ss.getActiveSheet();
  var startrow = ss.getActiveRange().getRow();
  var popup = Browser.msgBox('1) Are you on the Month Tab To Ship On? 2) Have your Mouse on the Start Row, Select "Yes" If Correct', 'Otherwise select "No" to Cancel', Browser.Buttons.YES_NO);
  
  if(popup == "yes"){
    activesheet.getRange(2, 24).setValue("Amanda");
          activesheet.getRange(2, 25).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")&" ➶ "&left(Z1/280*100,1)&"%"');
    activesheet.getRange(1, 26).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")');
  }
  if(popup == "no"){
    return 0;
  }
}



function Brodie() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInv = ss.getSheetByName("InStock");
  sheetInv.getRange("A1").setValue("Brodie");
  var activesheet = ss.getActiveSheet();
  var startrow = ss.getActiveRange().getRow();
  var popup = Browser.msgBox('1) Are you on the Month Tab To Ship On? 2) Have your Mouse on the Start Row, Select "Yes" If Correct', 'Otherwise select "No" to Cancel', Browser.Buttons.YES_NO);
  
  if(popup == "yes"){
    activesheet.getRange(2, 24).setValue("Brodie");
        activesheet.getRange(2, 25).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")&" ➶ "&left(Z1/280*100,1)&"%"');
    activesheet.getRange(1, 26).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")');
  }
  if(popup == "no"){
    return 0;
  }
}


function Rianne() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInv = ss.getSheetByName("InStock");
  sheetInv.getRange("A1").setValue("Rianne");
  var activesheet = ss.getActiveSheet();
  var startrow = ss.getActiveRange().getRow();
  var popup = Browser.msgBox('1) Are you on the Month Tab To Ship On? 2) Have your Mouse on the Start Row, Select "Yes" If Correct', 'Otherwise select "No" to Cancel', Browser.Buttons.YES_NO);
  
  if(popup == "yes"){
    activesheet.getRange(2, 24).setValue("Rianne");
           activesheet.getRange(2, 25).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")&" ➶ "&left(Z1/280*100,1)&"%"');
    activesheet.getRange(1, 26).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")');
  }
  if(popup == "no"){
    return 0;
  }
}


function Trevor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInv = ss.getSheetByName("InStock");
  sheetInv.getRange("A1").setValue("Trev");
  var activesheet = ss.getActiveSheet();
  var startrow = ss.getActiveRange().getRow();
  var popup = Browser.msgBox('1) Are you on the Month Tab To Ship On? 2) Have your Mouse on the Start Row, Select "Yes" If Correct', 'Otherwise select "No" to Cancel', Browser.Buttons.YES_NO);
  
  if(popup == "yes"){
    activesheet.getRange(2, 24).setValue("Trev");
      activesheet.getRange(2, 25).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")&" ➶ "&left(Z1/280*100,1)&"%"');
    activesheet.getRange(1, 26).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")');
  }
  if(popup == "no"){
    return 0;
  }
}


function Steve() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInv = ss.getSheetByName("InStock");
  sheetInv.getRange("A1").setValue("Steve");
  var activesheet = ss.getActiveSheet();
  var startrow = ss.getActiveRange().getRow();
  var popup = Browser.msgBox('1) Are you on the Month Tab To Ship On? 2) Have your Mouse on the Start Row, Select "Yes" If Correct', 'Otherwise select "No" to Cancel', Browser.Buttons.YES_NO);
  
  if(popup == "yes"){
    activesheet.getRange(2, 24).setValue("Steve");
    activesheet.getRange(2, 25).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")&" ➶ "&left(Z1/280*100,1)&"%"');
    activesheet.getRange(1, 26).setFormula('COUNTIFS(P'+startrow+':P,X2,M'+startrow+':M,"<>"&"")+COUNTIFS(P'+startrow+':P,X2,L'+startrow+':L,"*")');



  }
  if(popup == "no"){
    return 0;
  }
}





function Signout() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInv = ss.getSheetByName("InStock"); 
   var activesheet = ss.getActiveSheet()
activesheet.getRange("Y2").setValue("OFF");
activesheet.getRange("X2").setValue("OFF");





  sheetInv.getRange("A1").clearContent();
}