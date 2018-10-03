var couponSsId="1iV_wVlpuwDxhZBFqHv5ldlMHdo4r99Q05W-oPYKyAkA";


function staticstock2() 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("InStock");  


// sheet.getRange("A5:Z").copyTo(sheet.getRange("A5"), {contentsOnly:true});

sheet.getRange("A5:Z").copyTo(sheet.getRange("A5"));

}




function staticstock() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

 var sheet = ss.getSheetByName('InStock'); //replace with source Sheet tab name

 var range = sheet.getRange('A5:Z'); //assign the range you want to copy

 var data = range.getValues();

 sheet.getRange('A5:Z').setValues(data); //you will need to define the size of the copied data see getRange()

}




function formulastock() 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();

var currentDate=new Date();
  var current = Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")



  var sheet = ss.getSheetByName("InStock");  

var formula ="=filter('"+current+"'!A:Z,ISBLANK('"+current+"'!M:M),ISNUMBER('"+current+"'!L:L),('"+current+"'!I:I>0))";
 
 
 var range=sheet.getRange("A5");



sheet.getRange("A5:AC").clearContent();

sheet.getRange("AI5:AI").clearContent();

  range.setFormula(formula);

}



function refreshOnSale()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheetIS=ss.getSheetByName('InStock');
    var currentDate= new Date();
    var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy");
    var sheet=ss.getSheetByName(sheetName);
    var lr=sheet.getLastRow();


    var values=sheet.getRange(5, 1, lr-5+1, 19).getValues();
    var formulas=sheet.getRange(5, 1, lr-5+1, 19).getFormulas();
    
    var destValues=[];
    var destValues2=[]
    for (var i=0; i<values.length; i++)
    {
        if(values[i][13-1]=="" && !(isNaN(values[i][12-1]))  && values[i][9-1]>0)
        {
            destValues2.push(values[i]);
            for (var j=0; j<values[i].length; j++)
            {
                if(formulas[i][j]!="")
                {
                    values[i][j]=formulas[i][j];
                }
            }
            
            destValues.push(values[i]);
            
        
        }
    
    }
    
    sheetIS.getRange(5, 1, sheetIS.getLastRow()-5+1, 19).clearContent();
    sheetIS.getRange(5, 1, destValues.length, 19).setValues(destValues)
    
    
    for (var i=0; i<destValues.length; i++)
    {
        var url=destValues[i][19-1];
        var row=i+5;
        var orderId1=destValues2[i][4-1]
               var btn=url;  //if none found then just put url
               var bValue="";
               var bUrl="";
               
               
               var cbValue="";
               var cbUrl="";
               
               sheetIS.getRange(row, 7).setFontColor("#1155cc");//set font color blue for all order
               
               var couponButton="";
               
               if(url.indexOf("walmart")>=0)
               {
                     bValue="WM";
                     bUrl="https://www.topcashback.com/earncashback.aspx?mpurl=walmart&moid=50524";
                     /*
                     try
                     {
                         var stockLevel=walmarQty(url, currentVariation)
                         }
                         
                     catch(err){var stockLevel=4}    
                     sheet.getRange(row, 39).setValue(stockLevel);
                     */
               }
               
                              
               else if(url.indexOf("overstock")>=0)
               {
                 bValue="OS";
                 bUrl="https://www.topcashback.com/EarnCashback.aspx?mpurl=overstock&mpID=1004891";
                //   GmailApp.sendEmail("sakib118.biz@gmail.com", "Order Id", orderId);
                
                var headers = {                           
                                'ostkid': 'OSTK-VIP_18-A77359' //OSTK-VIP_18-A77359                          
                              };                           
                            var option = {                            
                                    "headers": headers,
                                    'muteHttpExceptions' : true                           
                              }; 
                   var html=UrlFetchApp.fetch(url,option).getContentText();
                 
                   var isSale=isOnSale(url, html);
                   
                     if(isSale)
                     {
                         sheetIS.getRange(row, 7).setFontColor("Red")
                       
                       
                       //change the formula of column N
                         sheetIS.getRange(row, 14).setValue("");
                         var frmO='=IF(M'+row+'="","",sum(K'+row+'*0.1188)+(L'+row+'))'
                         sheetIS.getRange(row, 15).setValue(frmO);
  
                     
                     }
                     
                     else
                     {
                         sheetIS.getRange(row, 7).setFontColor("#1155cc");
                         var mode="live"
                         if(mode=="test"){couponButton=retrieveCouponUrl(orderId1);}
                         else {couponButton=retrieveCouponUrl(orderId1);} //add coupon URl sale is not on
                         bUrl=url;  //new update 24 Mar 2017, if not on sale then button url is OS url
    
                     }
                     
                   // var stockLevel= getOSQty(html, currentVariation);
                   // sheet.getRange(row, 39).setValue(stockLevel);
                     
                 
               }
               
               else if(url.indexOf("atgstores")>=0)
               {
                 bValue="ATG";
                 bUrl="https://www.themine.com/";
               }
               else if(url.indexOf("themine")>=0)
               {
                 bValue="ATG";
                 bUrl="https://www.themine.com/";
               }
               else if(url.indexOf("wayfair")>=0)
               {
                 bValue="WF";
                 bUrl="http://www.wayfair.com/";
               }
               
               else if(url.indexOf("northerntool")>=0)
               {
                 bValue="NT";
                 bUrl="https://www.topcashback.com/EarnCashback.aspx?mpurl=northern-tool&mpID=1009377";
               }

               
               else if(url.indexOf("amazon")>=0)
               {
                 bValue="AMZ";
                 bUrl="http://www.amazon.com/";
               }
               
                else if(url.indexOf("sportsmansguide")>=0)
               {
                 bValue="SG";
                 bUrl="https://www.topcashback.com/EarnCashback.aspx?mpurl=the-sportsmans-guide&mpID=1006027";
               }
               
               else if(url.indexOf("hayneedle")>=0)
               {
                 bValue="HN";
                 bUrl="http://www.hayneedle.com";
               }
               
               else if(url.indexOf("kmart.com")>=0)
               {
                 bValue="KM";
                 bUrl="https://www.topcashback.com/earncashback.aspx?mpurl=kmart&moid=26887";
               }
               
                else if(url.indexOf("sears.com")>=0)
               {
                 bValue="SE";
                 bUrl="https://www.topcashback.com/earncashback.aspx?mpurl=sears&moid=14188";
               }
               else if(url.indexOf("aliexpress")>=0)
               {
                 bValue="AE";
                 bUrl="https://www.topcashback.com/EarnCashback.aspx?mpurl=aliexpress-by-alibaba&mpID=1005872";
               }
               
               
               
               
               
               
               
               
               if(bValue!="")
               {
                  btn='=HYPERLINK("'+bUrl+'","'+bValue+'")';
                             
               }
               
               //this will be used to combine prices of all items in a same order
               //combinedPrice.push([bValue, itemPrice, itemAfees] );
              /*
                if(i>0)   //when more than 1 item in order
               {
                   if(combinedPrice[i-1][0]!=combinedPrice[i][0]) //when supplier is not same
                   {
                       flagSameSup=1;
                   }
               
               }
               */
               
               
               sheetIS.getRange(row, 7).setValue(btn);
               sheetIS.getRange(row, 8).setValue(couponButton);
        
        
               
    
    }//end of for
    
    
    
    
    
    
    
    
    


}







function retrieveCouponUrl(orderId1)
{
    var ss=SpreadsheetApp.openById(couponSsId);
    var sheet=ss.getSheetByName("TEMPLATE");
    var lr=last_row(sheet,1);
    
    if(lr>3)
    {    
          var cUrl=sheet.getRange(lr,1).getValue();
          var cFrm='=HYPERLINK("'+cUrl+'","12%")';
    
          sheet.getRange(lr,2).setValue(sheet.getRange(lr,1).getValue());//move the url to column B
          sheet.getRange(lr,1).clearContent();
          sheet.getRange(lr,10).setValue(orderId1);
          return cFrm;
    }
    
    
    else if(lr<=3)  //take in 10%
    {
    
          var lr=last_row(sheet,3);
          
          if(lr<=3){return "N/A"};
          var cUrl=sheet.getRange(lr,3).getValue();
          var cFrm='=HYPERLINK("'+cUrl+'","10%")';
    
          sheet.getRange(lr,4).setValue(sheet.getRange(lr,3).getValue());//move the url to column D
          sheet.getRange(lr,3).clearContent();
          sheet.getRange(lr,11).setValue(orderId1);

    
          return cFrm;
    
    
    
    
    
    }
    
    
}














