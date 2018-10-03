
function sortordersafter5() 
{
  
  var currentTime = new Date().getHours();
  
  
  if(currentTime > 16)
  {
    Browser.msgBox("This Script is for Running on Orders after 5pm ONLY");
    
  }
  
  var popup = Browser.msgBox('1) Are you on the Month Tab To Ship On? 2) Have your Mouse on the Start Row, Select "Yes" If Correct, Wait for Script to Finish', 'Otherwise select "No" to Cancel', Browser.Buttons.YES_NO);
  
  if(popup == "yes"){
    
    
    
    
    // TCB  / {column: sortFirst, ascending: sortFirstAsc}
    
    var sortFirst = 7;
    var sortFirstAsc = false; 
    
    // Coupon   / {column: sortSecond, ascending: sortSecondAsc}
    
    var sortSecond = 8; 
    var sortSecondAsc = false; 
    
    // Order ID   / {column: sortThird, ascending: sortThirdAsc}
    
    var sortThird = 4; 
    var sortThirdAsc = false; 
    
    // Stock   / {column: sortFour, ascending: sortFourAsc}
    
    var sortFour = 39; 
    var sortFourAsc = true; 
    
    //  Multis  / {column: sortFifth, ascending: sortFifthAsc}
    
    var sortFifth = 36; 
    var sortFifthAsc = true; 
    
    // Quantity  / {column: sortSix, ascending: sortSixAsc}
    
    var sortSix = 3; 
    var sortSixAsc = false; 
    
    // Product Title / {column: sortSeven, ascending: sortSevenAsc}
    
    var sortSeven = 2; 
    var sortSevenAsc = true;
    
    // Sold Price / {column: sortEight, ascending: sortEightAsc}
    
    var sortEight = 9; 
    var sortEightAsc = false;
    
    
    // Processed By Shipper / {column: sortEight, ascending: sortEightAsc}
    
    var sortNine = 13; 
    var sortNineAsc = true;
    
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var headerRows = sheet.getActiveRange().getRow();
    var currentTime = new Date().getHours();
    
    
    if(currentTime > 16)
    {
      var sheetlast = sheet.getMaxRows();
    }
    
    else if(currentTime < 16)
    {
      Browser.msgBox("Wrong Script! Cancelling.."); 
      return 0;
    }
    
    var rangeprocessed = sheet.getRange(headerRows, 1, sheetlast-headerRows, sheet.getLastColumn());
    
    // SORT PROCESSED
    rangeprocessed.sort([ {column: sortNine, ascending: sortNineAsc}  ]);
    
    
    sheet.getRange(headerRows, 36, sheet.getLastRow()-headerRows, 1).clearContent();
    
    
    var lastprocessed = last_row(sheet, 13);
    
    
    if(lastprocessed < headerRows)
    {
      var lastprocessed = headerRows;
    }
    
    
    var range = sheet.getRange(lastprocessed, 1, sheetlast-lastprocessed, sheet.getLastColumn());
    var values=sheet.getRange(lastprocessed, 4, sheet.getLastRow(), 1).getValues();
    var backgrounds=sheet.getRange(lastprocessed, 4, sheet.getLastRow(), 1).getBackgrounds();
    
    sheet.getRange(lastprocessed-1, 36).setValue("Y");
    
    for (var i=0; i<values.length; i++)// start from row 5
    {
      
      if(backgrounds[i][0]=="#ead1dc")
      {
        sheet.getRange(i+lastprocessed, 36).setValue("M");
      }
    } 
    
    //SORT MULTIS, SORT ORDER ID, SORT HIGH PRICE
    range.sort([ {column: sortFifth, ascending: sortFifthAsc}, {column: sortThird, ascending: sortThirdAsc} , {column: sortEight, ascending: sortEightAsc}  ]);
    
    
    
    var lastm = last_row(sheet, 36);
    
    
    if(lastm < headerRows)
    {
      var lastm = headerRows;
    }
    
    
    var range2 = sheet.getRange(lastm, 1, sheetlast-lastm, sheet.getLastColumn());
    
    //SORT LOW STOCK
    range2.sort([ {column: sortFour, ascending: sortFourAsc} ]);
    
    
    
    if(lastm < headerRows)
    {
      var lastm = headerRows;
    }
    
    var values=sheet.getRange(lastm, 39, sheet.getLastRow(), 1).getValues();
    var backgrounds=sheet.getRange(lastm, 39, sheet.getLastRow(), 1).getBackgrounds();
    
    
    for (var i=0; i<values.length; i++)
    {
      
      if(backgrounds[i][0]=="#f4c7c3")
      {
        sheet.getRange(i+lastm, 36).setValue("S");
      }
      
      
    } 
    
    
    var lastaz = last_row(sheet, 36);
    
    
    if(lastaz < headerRows)
    {
      var lastaz = headerRows;
    }
    
    
    var diff = lastm-lastaz
    
    if(diff == 0)
    {
      
      var junk = "";
      
    }
    else if(diff >0 )
    {
      
      var range8 = sheet.getRange(lastm, 1, lastaz-lastm, sheet.getLastColumn());
      
      
      range8.sort([ {column: sortSeven, ascending: sortSevenAsc} ]);
      
    }
    
    
    var lasts = last_row(sheet, 36);



    var range3 = sheet.getRange(lasts, 1, sheetlast-lasts, sheet.getLastColumn());
    
    // SORT HIGH QTY ORDERS
    range3.sort([ {column: sortSix, ascending: sortSixAsc} ]);
    
    
    
    var lasts = last_row(sheet, 36);
    var values=sheet.getRange(lasts, 3, sheet.getLastRow(), 1).getValues();
    
    for (var i=0; i<values.length; i++)
    {
      
      if(values[i][0] > 1)
      {
        sheet.getRange(i+lasts, 36).setValue("Q");
      }
    } 
    
    
    var lastqty = last_row(sheet, 36);
    
    var diff = lastqty-lastaz
    
    if(diff == 0)
    {
      
      var junk = "";
      
    }
    else if(diff >0 )
    {
      var range9 = sheet.getRange(lastaz, 1, lastqty-lastaz, sheet.getLastColumn());
      
      
      range9.sort([ {column: sortSeven, ascending: sortSevenAsc} ]);
    }
    
    
    
    
    
    
    var lastq = last_row(sheet, 36);
    var range4 = sheet.getRange(lastq, 1, sheetlast-lastq, sheet.getLastColumn());
    
    
    // SORT SUPPLIERS / COUPON,SALE
    range4.sort([ {column: sortFirst, ascending: sortFirstAsc}, {column: sortSecond, ascending: sortSecondAsc}, {column: sortSeven, ascending: sortSevenAsc}  ]);
    
  }


  if(popup == "no"){
    return 0;
  }
  
  
  Browser.msgBox("Complete! Wait 10 More Seconds to Finish Loading - Double Check for Errors"); 
  
}








  
function sortorders() 
{

    var currentTime = new Date().getHours();


if(currentTime > 16)
{
  Browser.msgBox("If After 5PM, Please get the Last Row Number for your Shift - Will Need it In This Script");

}

  

  
  var popup = Browser.msgBox('1) Are you on the Month Tab To Ship On? 2) Have your Mouse on the Start Row, Select "Yes" If Correct, Wait for Script to Finish', 'Otherwise select "No" to Cancel', Browser.Buttons.YES_NO);
  
  if(popup == "yes"){
    
    
    
    
    
    // TCB  / {column: sortFirst, ascending: sortFirstAsc}
    
    var sortFirst = 7; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
    var sortFirstAsc = false; //Set to false to sort descending
    
    // Coupon   / {column: sortSecond, ascending: sortSecondAsc}
    
    var sortSecond = 8; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
    var sortSecondAsc = false; //Set to false to sort descending
    
    // Order ID   / {column: sortThird, ascending: sortThirdAsc}
    
    var sortThird = 4; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
    var sortThirdAsc = false; //Set to false to sort descending
    
    // Stock   / {column: sortFour, ascending: sortFourAsc}
    
    var sortFour = 39; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
    var sortFourAsc = true; //Set to false to sort descending
    
    //  Multis  / {column: sortFifth, ascending: sortFifthAsc}
    
    var sortFifth = 36; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
    var sortFifthAsc = true; //Set to false to sort descending
    
    // Quantity  / {column: sortSix, ascending: sortSixAsc}
    
    var sortSix = 3; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
    var sortSixAsc = false; //Set to false to sort descending
    
    // Product Title / {column: sortSeven, ascending: sortSevenAsc}
    
    var sortSeven = 2; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
    var sortSevenAsc = true; //Set to false to sort descending
    
    // Sold Price / {column: sortEight, ascending: sortEightAsc}
    
    var sortEight = 9; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
    var sortEightAsc = false; //Set to false to sort descending

    // Processed By Shipper / {column: sortEight, ascending: sortEightAsc}
    
    var sortNine = 13; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
    var sortNineAsc = true; //Set to false to sort descending
    
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var headerRows = sheet.getActiveRange().getRow();
     var currentTime = new Date().getHours();


if(currentTime < 17)
{
var sheetlast = sheet.getMaxRows();
}
else if(currentTime > 16)
{
var sheetlast = Browser.inputBox("Input Last Row Number of Last Order/Shift (Before 5PM)");

}




 
    var rangeprocessed = sheet.getRange(headerRows, 1, sheetlast-headerRows, sheet.getLastColumn());
    
    // SORT PROCESSED
   rangeprocessed.sort([ {column: sortNine, ascending: sortNineAsc}  ]);
    
    
    
    sheet.getRange(headerRows, 36, sheet.getLastRow()-headerRows, 1).clearContent();



 var lastprocessed = last_row(sheet, 13);
 
 
 if(lastprocessed < headerRows)
{
var lastprocessed = headerRows;
}

    
    var range = sheet.getRange(lastprocessed, 1, sheetlast-lastprocessed, sheet.getLastColumn());
    var values=sheet.getRange(lastprocessed, 4, sheet.getLastRow(), 1).getValues();
    var backgrounds=sheet.getRange(lastprocessed, 4, sheet.getLastRow(), 1).getBackgrounds();
    
    sheet.getRange(lastprocessed-1, 36).setValue("Y");
    
    for (var i=0; i<values.length; i++)// start from row 5
    {
      
      if(backgrounds[i][0]=="#ead1dc")
      {
        sheet.getRange(i+lastprocessed, 36).setValue("M");
      }
    } 
    
    //SORT MULTIS, SORT ORDER ID, SORT HIGH PRICE
    range.sort([ {column: sortFifth, ascending: sortFifthAsc}, {column: sortThird, ascending: sortThirdAsc} , {column: sortEight, ascending: sortEightAsc}  ]);
    
    var lastm = last_row(sheet, 36);
    var range2 = sheet.getRange(lastm, 1, sheetlast-lastm, sheet.getLastColumn());
    
    //SORT LOW STOCK
    range2.sort([ {column: sortFour, ascending: sortFourAsc} ]);
    
    
    
    var lastm = last_row(sheet, 36);
    var values=sheet.getRange(lastm, 39, sheet.getLastRow(), 1).getValues();
    var backgrounds=sheet.getRange(lastm, 39, sheet.getLastRow(), 1).getBackgrounds();
    
    
    for (var i=0; i<values.length; i++)
    {
      
      if(backgrounds[i][0]=="#f4c7c3")
      {
        sheet.getRange(i+lastm, 36).setValue("S");
      }
      
      
    } 
    
    
    var lastaz = last_row(sheet, 36);
    
    
    var diff = lastm-lastaz
    
    if(diff == 0)
    {
      
      var junk = "";
      
    }
    else if(diff >0 )
    {
      
      var range8 = sheet.getRange(lastm, 1, lastaz-lastm, sheet.getLastColumn());
      
      
      range8.sort([ {column: sortSeven, ascending: sortSevenAsc} ]);
      
    }
    
    
    var lasts = last_row(sheet, 36);
    var range3 = sheet.getRange(lasts, 1, sheetlast-lasts, sheet.getLastColumn());
    
    // SORT HIGH QTY ORDERS
    range3.sort([ {column: sortSix, ascending: sortSixAsc} ]);
    
    
    
    var lasts = last_row(sheet, 36);
    var values=sheet.getRange(lasts, 3, sheet.getLastRow(), 1).getValues();
    
    for (var i=0; i<values.length; i++)
    {
      
      if(values[i][0] > 1)
      {
        sheet.getRange(i+lasts, 36).setValue("Q");
      }
    } 
    
    
    var lastqty = last_row(sheet, 36);
    
    var diff = lastqty-lastaz
    
    if(diff == 0)
    {
      
      var junk = "";
      
    }
    else if(diff >0 )
    {
      var range9 = sheet.getRange(lastaz, 1, lastqty-lastaz, sheet.getLastColumn());
      
      
      range9.sort([ {column: sortSeven, ascending: sortSevenAsc} ]);
    }
    
    
    
    
    
    
    var lastq = last_row(sheet, 36);
    var range4 = sheet.getRange(lastq, 1, sheetlast-lastq, sheet.getLastColumn());
    
    
    // SORT SUPPLIERS / COUPON,SALE
    range4.sort([ {column: sortFirst, ascending: sortFirstAsc}, {column: sortSecond, ascending: sortSecondAsc}, {column: sortSeven, ascending: sortSevenAsc}  ]);
    
  }


  if(popup == "no"){
    return 0;
  }
  
  
  Browser.msgBox("Complete! Wait 10 More Seconds to Finish Loading - Double Check for Errors"); 
  
}





