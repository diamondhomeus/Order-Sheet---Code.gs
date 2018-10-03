

function highLightDuplicatesauto() {

    var ss=SpreadsheetApp.getActiveSpreadsheet();
  
          var currentDate=new Date();
        var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
        
        
    var sheet=ss.getSheetByName(sheetName);
    
    
    var lr = last_row(sheet, 2);


if(lr < 300)
{
 var start = 5;

}

else{


var start = lr-300;

}

Logger.log(start);
    
    var orderIds=sheet.getRange(start , 4, 300  , 1).getValues();
   

    var orderIds1D=orderIds.join("|").split("|")
    var currentBackgrounds=sheet.getRange(start , 2, 300  , 1).getBackgrounds();
    
    var orderBackgrounds=sheet.getRange(start , 4, 300  , 1).getBackgrounds();
    //var sold=sheet.getRange("D5:D").getValues();
    
    for ( var i=8; i<orderIds1D.length; i++)
    {
          var orderIdtemp=orderIds[i][0];
          if(orderIdtemp=="")
            {continue};
            var a=orderIds1D.lastIndexOf(orderIdtemp);
          if(orderIds1D.lastIndexOf(orderIdtemp)>i )
          {
     
 
              if( orderBackgrounds[i][0] == "#ead1dc" )
{
              currentBackgrounds[i][0]="White"
}
else
{
currentBackgrounds[i][0]="Red"

}
    
  
              
          }
    
    }
    sheet.getRange(start , 2, 300  , 1).setBackgrounds(currentBackgrounds);



  
}



function highLightDuplicatestitleauto() {

    var ss=SpreadsheetApp.getActiveSpreadsheet();
          var currentDate=new Date();
        var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
        
        
    var sheet=ss.getSheetByName(sheetName);
    
     var lr = last_row(sheet, 2);

if(lr < 300)
{
 var start = 5;

}

else{


var start = lr-300;

}

Logger.log(start);

    var orderIds=sheet.getRange(start , 25, 300  , 1).getValues();
    var orderIds1D=orderIds.join("|").split("|")
    var currentBackgrounds=sheet.getRange(start , 2, 300  , 1).getBackgrounds();
     var orderBackgrounds=sheet.getRange(start , 4, 300  , 1).getBackgrounds();
    
    for ( var i=8; i<orderIds1D.length; i++)
    {
          var orderIdtemp=orderIds[i][0];
          if(orderIdtemp=="")
            {continue};
            var a=orderIds1D.lastIndexOf(orderIdtemp);
          if(orderIds1D.lastIndexOf(orderIdtemp)>i)
          {
    
          
          
           if( orderBackgrounds[i][0] == "#ead1dc" )
{
              currentBackgrounds[i][0]="White"
}
else
{
currentBackgrounds[i][0]="Orange"

}
}
    
    
    }
   sheet.getRange(start , 2, 300  , 1).setBackgrounds(currentBackgrounds);
  
}







function highLightDuplicates() {

    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    
    

   

    var orderIds=sheet.getRange("D5:D").getValues();
    var orderIds1D=orderIds.join("|").split("|")
    var currentBackgrounds=sheet.getRange("B5:B").getBackgrounds();
    //var sold=sheet.getRange("D5:D").getValues();
    
    for ( var i=8; i<orderIds1D.length; i++)
    {
          var orderIdtemp=orderIds[i][0];
          if(orderIdtemp=="")
            {continue};
            var a=orderIds1D.lastIndexOf(orderIdtemp);
          if(orderIds1D.lastIndexOf(orderIdtemp)>i)
          {
                currentBackgrounds[i][0]="Red"
        
          }
    
    }
    sheet.getRange("B5:B").setBackgrounds(currentBackgrounds);
  
highLightDuplicatestitle();
}



function resetDuplicates()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    sheet.getRange("B5:B").setBackground("#ffffff");

}


function highLightDuplicatestitle() {

    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    var orderIds=sheet.getRange("B5:B").getValues();
    var orderIds1D=orderIds.join("|").split("|")
    var currentBackgrounds=sheet.getRange("B5:B").getBackgrounds();
    //var sold=sheet.getRange("D5:D").getValues();
    
    for ( var i=8; i<orderIds1D.length; i++)
    {
          var orderIdtemp=orderIds[i][0];
          if(orderIdtemp=="")
            {continue};
            var a=orderIds1D.lastIndexOf(orderIdtemp);
          if(orderIds1D.lastIndexOf(orderIdtemp)>i)
          {
                currentBackgrounds[i][0]="Orange"
        
          }
    
    }
    sheet.getRange("B5:B").setBackgrounds(currentBackgrounds);
  
}


