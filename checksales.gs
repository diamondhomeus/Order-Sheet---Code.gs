
function checkSales() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetInv = ss.getSheetByName("InStock");



sheetInv.getRange("AI5:AI").clearContent();



 var range = sheetInv.getRange('A5:Z'); //assign the range you want to copy

 var data = range.getValues();

 sheetInv.getRange('A5:Z').setValues(data); //you will need to define the size of the copied data see getRange()


  var numRows=100;
  var lrK=last_row(sheetInv, 35);
  
  
  
  var sr=lrK+1;
  var lr=sr+numRows;
  var lrAll=last_row(sheetInv, 19);//sheetInv.getLastRow();
  
  if(lrAll-lrK<=numRows)
  {
      lr=lrAll;
  
  }
  
  
  if(sr>=lrAll){ return 0; }

  
  
  var rng = sheetInv.getRange(sr, 19, lr-sr+1).getValues();
  var initials = sheetInv.getRange(sr, 26, lr-sr+1).getValues();
  var vals = [];
  var myDate=Utilities.formatDate(new Date(), "CST", "MMM-dd-yyyy");
  var refreshDate=[];
  
  var prev_sale = false;
  
  var startTime=(new Date()).getTime();
  var threshold=4.9*60*1000;
  
  for (var i=0; i<rng.length; i++)
  {
    var source = rng[i].toString();
    
    refreshDate.push([myDate]);

    
    if (source == "" || source.indexOf("overstock")<0) { vals.push(["N/A"]); continue; }
    
    
    var index=rng.indexOf(source)
    
    
    
    
    if (i > 0) //second iteration
    {
            var prevSource = rng[i-1].toString();
            
            if (prevSource == source) {
              var isSale = prev_sale;
            } 
            
            else {               //when source 1 is different, check if -SA items
                      if(initials[i].toString().toLowerCase().indexOf('-sa')>0)
                      { 
                          var isSale=true; 
                          
                          }
                      else
                      {
                            var isSale = isOnSale(source);
                            prev_sale = isSale;
                      }      
            }
    } 
    
    else if (i == 0) //first iteration
    {
      
                          if(initials[i].toString().toLowerCase().indexOf('-sa')>0)
                                          { 
                                              var isSale=true; 
                                              }
                                              
                          else{                    
                                  var isSale = isOnSale(source);
                                  }
                          prev_sale = isSale;
    }
    
    if (isSale == true) {
      vals.push(["Yes"]);
    } else {
      vals.push(["No"]);
    }
    
    if(i%100==0)  //check time for each 100 rows
    {
            var endTime=(new Date()).getTime();
            Logger.log((endTime-startTime)+"                  "+ i)
            
            
            if(endTime-startTime>threshold)
            {
               break;
            
            }

    
    
    }
    
    
    if (i == numRows) {break;}  // Comment for whole loop
  }
  
  sheetInv.getRange(sr, 35, vals.length).setValues(vals);

  //Browser.msgBox("Check for ON SALE for inventory list complete!")
}





function isOnSale(url)
{
  //var url="http://www.overstock.com/Home-Garden/HomePop-Large-Teal-Blue-Decorative-Storage-Ottoman/10293207/product.html";
  try
  {
      var headers = {                           
        'ostkid': 'OSTK-VIP_18-A77359' //OSTK-VIP_18-A77359                          
      };                           
    var option = {                            
      "headers": headers,
      'muteHttpExceptions' : true                           
    }; 
      
      
      var html=UrlFetchApp.fetch(url,option).getContentText();
  }
  
  catch(e)
  {
      var r=true 
  }
  
  if(html==undefined){return true}
  
  var n1=html.indexOf("price-title");
  var n2=html.indexOf(">",n1);
  var n3=html.indexOf("<",n2);
  var priceTitle=html.slice(n2+1,n3);
  // GmailApp.sendEmail("sakib118.biz@gmail.com", "Sale", priceTitle)
  if(priceTitle.indexOf("Sale")>=0)
  {var r= true;}
  else
  {r= false;}
  
  if(html.indexOf('DoorBustersIcon')>0)
  {
    r=true;
  }
  
  if(html.indexOf('DoorbusterIcon')>0)  //for weekly flash deals
  {
    r=true;
  }
  
  
  
  
  
  return r;
}
