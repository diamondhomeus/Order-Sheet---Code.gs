

function importVariation() 
{
  var chkStock_count_rows = 150;  //number of rows that it can go max
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSTOCK = ss.getSheetByName("InStock");  //changed by Brodie


 var range = sheetSTOCK.getRange('A5:Z'); //assign the range you want to copy

 var data = range.getValues();

 sheetSTOCK.getRange('A5:Z').setValues(data); //you will need to define the size of the copied data see getRange()



  
  var sortFirst = 7; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
  var sortFirstAsc = false; //Set to false to sort descending

  var sortSecond = 8; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
  var sortSecondAsc = false; //Set to false to sort descending
  
  
  
  //Number of header rows
  
  var headerRows = 4; 

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getActiveSheet();


  var range = sheet.getRange(headerRows+1, 1, sheet.getMaxRows()-headerRows, sheet.getLastColumn());

  
  range.sort([{column: sortFirst, ascending: sortFirstAsc},{column: sortSecond, ascending: sortSecondAsc}]);


  //var f=sheetSTOCK.getRange("T6").getFormulaR1C1()
  var start = last_row(sheetSTOCK, 27)+1;
  var lr=sheetSTOCK.getLastRow();
  
  if(start==lr){
        //sheetSTOCK.getRange("AJ5:AL").clearContent();
        //start=5;
    return 0;
    }
var numRows=lr-start+1;  
  if(numRows<chkStock_count_rows){chkStock_count_rows=numRows};
  
  var gridRng = sheetSTOCK.getRange(start, 1, chkStock_count_rows, 29).getValues();
  var vals = [];
  var isStock = "";
  var list_vals = [];
  var formulas = [];
  var prev_stock = "";
  var partialCount = 0;
  //var getFormula=sheetSTOCK.getRange("F4").getFormulaR1C1();
  var getFormula='=IFERROR(IF(R[0]C[-1]="Yes", IF(SEARCH(">"&R[0]C[-22]&"<",R[0]C[1])>0,"Yes","No"),""),"No")';
  
  var startTime=(new Date()).getTime();
  var threshold=4*60*1000; // in miliseconds
  
  
  
  
  
  
  for (var i=0; i<gridRng.length; i++)  //for each values
  {
        var tmpArr = [];
      
        var url = gridRng[i][19-1];   //source 1   //changed by Brodie
        Logger.log(url);
        var stock = gridRng[i][2-1]; //skugrid stock
        var asin = gridRng[i][3-1];  //asin 
        var old_variation = gridRng[i][6-1];         //changed by Brodie
        
            
        if (url == "") {
                txtVars="";
                vals.push([txtVars]);
                txtVars = "";
                formulas.push([""]);
                list_vals.push([""]);
                continue;
        }
        
        
        
        
        
        
        
        
        
        
     
     
     
     
       
    
        if(i>=1)  //if same url do not import
        {        
                 var prevUrl=gridRng[i-1][1-1];
                 if(prevUrl==url)
                 {
                                  vals.push([txtVars]);
                                  list_vals.push([isStock]);
                                  formulas.push([getFormula]);
                                  continue;
                 
                 }
        
        
        }
    
               
               
               txtVars = "";
               isStock = "";
    
    
    
    

       
       try
       {
             if(url.indexOf('overstock')>=0)
                
                {
                            var headers = {                           
                                'ostkid': 'OSTK-VIP_18-A77359' //OSTK-VIP_18-A77359                          
                              };                           
                            var option = {                            
                                    "headers": headers,
                                    'muteHttpExceptions' : true                           
                              }; 
                }                
                else
                {
                          var option = {
                                'muteHttpExceptions' : true 
                          };
                }
                var html = UrlFetchApp.fetch(url, option).getContentText();               
                
       }
       
       
       catch(err)
       {
                vals.push([txtVars]);
                var txtVars = "";
                isStock = "Bad Url";
                list_vals.push([isStock]);
                formulas.push([""]);
                continue;                                 
   
             
       
       }
          if (html.toLowerCase().indexOf("Out of stock") >= 0 || html==undefined) //out of stock on OS
          {
                
                var txtVars = "";
                vals.push([txtVars]);
                isStock = "No";
                list_vals.push([isStock]);
                formulas.push([""]);
                continue;
          }
          
          if(url.indexOf('overstock')>=0)
          {
                           var n11=html.indexOf("facetGroupId");
                           var n12=html.indexOf("options-dropdown");
                           if(n11>=-1 && n12>=-1){var n1=n11+n12+1}
                           else {var n1=-1}
    
                          if (n1 >= 0) //there is a variation
                          {
                                  var n2=html.indexOf("</select>",n1);
                                  var prodOptions=html.slice(n1,n2);
                            
                                   var q1=html.lastIndexOf('dropDownOptions:');
                                   var q2=html.indexOf('fromOptionBasedRequest',q1);
                                   var qtyHtml=html.slice(q1,q2);
                                   var qtyArr=qtyHtml.split('description');
                                           
                            
                                  var arrOptions=prodOptions.split("value");
                                  var txtVars="";
                                  isStock="No"  
                                    for (var k=1; k<qtyArr.length ; k++) 
                                    {
                                            var value=qtyArr[k];
                                            
                                            
                                            var m1=value.indexOf(":")+2;
                                            var m2=value.indexOf("containsProduct",m1);
                                            var variation=value.slice(m1,m2-3);
                                            variation = replaceAll(replaceAll(replaceAll(variation,"&quot;",'"'),"&amp;","&"),'\u0027',"'");
                                            variation=replaceAll(variation,'\\u0027',"'");
                                            variation=variation.replace(/\r?\n|\r/g,"")
                                            
                                            
                                                                                        
                                            
                                            // var longQty=qtyArr[k-1];
                                            var q1=value.indexOf("maxQuantity");
                                            var q2=value.indexOf("status",q1);
                                            var qty=value.slice(q1+13,q2-2);
                                            
                                            var qty_int = parseInt(qty);
                                            if(qty_int>0)
                                            {
                                                isStock="Yes"; 
                                            }
                                            if(txtVars==""){txtVars=">"};
                                            txtVars += variation+"<"+qty+">";
                                            
                                                                                       
                                      }  //end of k loop
                                  
                                          var targetRow = start+1+i;
                                          
                                          list_vals.push([isStock])
                                          formulas.push([getFormula]);
                                          vals.push([txtVars]);
                            
                            } //end of if there is variation
                          
                          else if(n1==-1) 
                          {
                                  txtVars = "";
                                  vals.push([txtVars]);
                                  
                                            var q1=html.indexOf('dropDownOptions');
                                            var q2=html.indexOf('fromOptionBasedRequest',q1);
                                            var qtyHtml=html.slice(q1,q2);
                                            var qtyArr=qtyHtml.split('maxQuantity');
                                            var longQty=qtyArr[1];
                                            if(longQty!=undefined) //whole listing is not OOS
                                            {
                                              var q1=longQty.indexOf(":");
                                              var q2=longQty.indexOf(",",q1);
                                              var qty=longQty.slice(q1+1,q2);
                                              var qty=Number(qty);
                                             } 
                                             
                                             else
                                             {
                                                 qty=0;
                                             }
                                  
                                  qty=qty==0?"No":"Yes"

                                  isStock = qty;  //there is no variation but stock is availalble
                                  list_vals.push([isStock]);
                                  formulas.push([""]);

                          }//if no varitaion
              }// end of overstock
              
              else if(url.indexOf('walmart')>=0)
              {

                      try                    //when oos in WM
                       {
                           var jsonData=getMyJson(html);
                       }
                       catch(err)
                       { 
                           var txtVars = "";
                            vals.push([txtVars]);
                            isStock = "No";
                            list_vals.push([isStock]);
                            formulas.push([""]);
                            continue;
                       }
                       
                       var prodId=jsonData.productId;
                       var prodName="";
        
        
        
                       var productBasicInfo=jsonData.productBasicInfo;
                       var selectedProduct=productBasicInfo.selectedProductId;
                       var selectedProdDetails=productBasicInfo[selectedProduct];
                       var wmTitle=selectedProdDetails.title;
                       var itemNo=selectedProdDetails.usItemId;
                       
                       var product=jsonData.product;
                       var products=product.products;
                       
                       
                       var allOffers=jsonData.product.offers;
        
        //push all images to an array using 'variantCategoriesMap'
       
                       var prodsForImages=[];
                       var allImageIds=[];
        

                        var primaryProduct=product.primaryProduct; //varaition map starts with base product
                        var varMap=product.variantCategoriesMap[primaryProduct]; // first property is the primay product
            
                        if(varMap==undefined)  //when there is no variation
                        {        Logger.log(url);  
                                 var inStock=importWmNoVariation_(jsonData);
                                 if(inStock==null)
                                 {
                                              txtVars = "";
                                              vals.push([txtVars]);
                                              
                                              isStock = "No"; 
                                              list_vals.push([isStock]);
                                              formulas.push([""]);
                                              continue;
                                  }
                                  
                                 else if(inStock>0)
                                 {
                                              txtVars = "";
                                              vals.push([txtVars]);
                                              
                                              isStock = "Yes";  //there is no variation but stock is availalble
                                              list_vals.push([isStock]);
                                              formulas.push([""]);
                                              continue;
                                  }
                                  
                                  
                         }//no variation ends 
                         
                     var txtVars="";    //to be pasted in the skugrid app
                     isStock="No"
                     for(var p in products)
                     {
                                          //--check if overwrriting the exististing data
                                         
                                        var rowArr=[];
                                        var dProd=products[p]; ///daughter product i.e. this product
                                        //var itemNo=dProd.usItemId; // commented because using selected item id
                                        
                                        var variantsProp=dProd.variants; // all variants of this product
                                       
                                         var count=0;
                                         var variation="";
                                         
                                         var flag1=0;
                                         var flag2=0;
                             try{             // these two arrays will all variation information
                                         var cv=varMap.actual_color;
                                         if(cv!=undefined)
                                          {var colorVars=cv.variants;}
                                          else
                                          {flag1=1;}
                                          
                                          
                                         var sv= varMap.size;
                                         if(sv!=undefined) 
                                          {var sizeVars=sv.variants;}
                                         else
                                         {flag2=1;}
                                        
                                      
                                         
                                         
                                         
                                         
                                         
                                         
                                         //get the variant details
                                         if(flag1==0 && flag2==0)
                                         {
                                                 var sizeProp=variantsProp.size;
                                                 var sizeName=sizeVars[sizeProp].name;
                                                 
                                                 var colorProp=variantsProp.actual_color;
                                                 var colorName=colorVars[colorProp].name;
                                                 
                                                 var skugridVar=sizeName+'|'+colorName;
                                          }
                                          
                                          else if  (flag1==0)  //only color vari
                                          {
                                                 var colorProp=variantsProp.actual_color;
                                                 var colorName=colorVars[colorProp].name;
                                                 var skugridVar=colorName;
                                          }
                                          
                                          
                                          else if  (flag2==0)  //only color vari
                                          {
                                                 var sizeProp=variantsProp.size;
                                                 var sizeName=sizeVars[sizeProp].name;
                                                 var skugridVar=sizeName;
                    
                                          }
                                          
                                          //end of generating variations
                                          
                                          
                                          var thisItem=dProd.usItemId;
                                          var sellerUrl= 'https://www.walmart.com/product/'+thisItem+'/sellers'; // imports the list of sellers
                                          
                                            var option = {
                                                  'muteHttpExceptions' : true
                                            };
                    
                                         var htmlSeller = UrlFetchApp.fetch(sellerUrl, option).getContentText();
                                         var n1=htmlSeller.indexOf('window.__WML_REDUX_INITIAL_STATE__ = ')+('window.__WML_REDUX_INITIAL_STATE__ = ').length;
                                         var n2=htmlSeller.indexOf('</script>',n1)-3;
                                         var htmlSeller=htmlSeller.slice(n1,n2);
                                         var jsonDataSeller=JSON.parse(htmlSeller);
                                         var selectedProd=jsonDataSeller.product.selected.product;
                                        
                                         var detailsOfSelectedProd=jsonDataSeller.product.products[selectedProd];
                                         var flagWm=0;
                                         if(detailsOfSelectedProd!=undefined)
                                          {
                                            var myOffers=jsonDataSeller.product.products[selectedProd].offers
                                            var allOffers=jsonDataSeller.product.offers;
                                          }
                                          
                                         else
                                          {
                                              var flagWm=1;
                                          }
                                          
                                          
                    
                                          
                                          
                                         
                                          //var k=-1  //stop loop for texting
                                          for(var k=0; k<myOffers.length && flagWm==0; k++)
                                          {
                                                   var tempOfferId=myOffers[k];
                                                   var tempOffer=allOffers[tempOfferId];
                                                   var isStock2=tempOffer.productAvailability.availabilityStatus;
                                                   
                                            }       
                                           
                                          
                                            if(isStock2=='IN_STOCK')
                                            {
                                                  var isWMStock=5;
                                                  isStock="Yes"
                                            
                                            }
                      
                                            else
                                            {
                                                  var isWMStock=0;
                                                  
                                            }
                               }//end of try      
                               
                               
                            catch(err)
                            {
                                          variation="Invalid";
                                          isWMStock=0;
                                          qty=0;
                                          if(txtVars==""){txtVars=">"};
                                          txtVars += variation+"<"+qty+">";    
                            
                            
                            
                            }
                            
                                          variation=skugridVar;
                                          qty=isWMStock;
                                          if(isWMStock>0)
                                          {
                                            if(txtVars==""){txtVars=">"};
                                            txtVars += variation+"<"+qty+">";
                                          }
                                         
                      
                     }
                     //end of for product for
                     
                var targetRow = start+1+i;
                list_vals.push([isStock])
                formulas.push([getFormula]);
                vals.push([txtVars]);
                         
                         
                         
             
              
              }  //end of walmart
              
              else
              {
                      var targetRow = start+1+i;
                      list_vals.push([""])
                      formulas.push([""]);
                      vals.push([""]);
              
              
              }
              
              
              
              
              
                                
              if(i%10==0)   //check time elapsed in each fifty iterations 
              {
                          var currentTime=(new Date()).getTime();
                          
                          if(currentTime-startTime>threshold)
                          {
                                break;
                          
                          }
                    
              
              }
                                
                                
                                
    
  }  //end of i loop
  
  sheetSTOCK.getRange(start, 27, list_vals.length).setValues(list_vals);   //changed by Brodie
  sheetSTOCK.getRange(start, 28, formulas.length).setValues(formulas);      //changed by Brodie
  sheetSTOCK.getRange(start, 29, vals.length).setValues(vals);             //changed by Brodie
  
  ss.toast("Listing Complete!");
}


























//imports stock from single variation walmart
function importWmNoVariation_(jsonData)
{

       
        
        if(jsonData==undefined){
           
              return null;
        
        }
        
        
        var prodId=jsonData.productId;
        var prodName="";
        
        
        
        var productBasicInfo=jsonData.productBasicInfo;
        var selectedProduct=productBasicInfo.selectedProductId;
        var selectedProdDetails=productBasicInfo[selectedProduct];
        var wmTitle=selectedProdDetails.title;
        
   
        var product=jsonData.product;
        var products=product.products;

        
              
        
        var allOffers=jsonData.product.offers;
        
        //push all images to an array using 'variantCategoriesMap'
       
        var prodsForImages=[];
        var allImageIds=[];
        

        var primaryProduct=product.primaryProduct; //varaition map starts with base product
            
                  
        
        
        
        
        for(var i in products)
        {
                      
                     //---------------------
                   
                    var rowArr=[];
                    var dProd=products[i]; ///daughter product i.e. this product
                    var itemNo=dProd.usItemId;
                    
                    var variantsProp=dProd.variants; //variants of this product
                   
                    var count=0;
                    var variation="";
 
                      var thisItem=dProd.usItemId;
                      var sellerUrl= 'https://www.walmart.com/product/'+thisItem+'/sellers'; // imports the list of sellers
                      
                        var option = {
                              'muteHttpExceptions' : true
                        };

                     var htmlSeller = UrlFetchApp.fetch(sellerUrl, option).getContentText();
                     var n1=htmlSeller.indexOf('window.__WML_REDUX_INITIAL_STATE__ = ')+('window.__WML_REDUX_INITIAL_STATE__ = ').length;
                     var n2=htmlSeller.indexOf('</script>',n1)-3;
                     var htmlSeller=htmlSeller.slice(n1,n2);
                     var jsonDataSeller=JSON.parse(htmlSeller);
                     var selectedProd=jsonDataSeller.product.selected.product;
                    
                     var detailsOfSelectedProd=jsonDataSeller.product.products[selectedProd];
                     var flagWm=0;
                     if(detailsOfSelectedProd!=undefined)
                      {
                        var myOffers=jsonDataSeller.product.products[selectedProd].offers
                        var allOffers=jsonDataSeller.product.offers;
                      }
                      
                     else
                      {
                          var flagWm=1;
                      }
                      
                      

                      if(myOffers==undefined || myOffers.length==0){return null};
                      
                     
                      //var k=-1  //stop loop for texting
                      for(var k=0; k<myOffers.length && flagWm==0; k++)
                      {
                               var tempOfferId=myOffers[k];
                               var tempOffer=allOffers[tempOfferId];
                               var isStock=tempOffer.productAvailability.availabilityStatus;                                 
                        
                      }  // emd of offers for


                      if(isStock=='IN_STOCK')
                      {
                            isStock=1;
                      
                      }

                      else
                      {
                          isStock=0;
                      }
    
              
              
   
        
        }//end of product for

       return isStock;
  









}





function getMyJson(html)
{
            
        var n1=html.indexOf('</div><script id="atf-content" type="application/json">')+('</div><script id="atf-content" type="application/json"> ').length;
        var n2=html.indexOf('</script>',n1);

        var html2=html.slice(n1,n2);
      try
      {
        var jsonData=JSON.parse(html2)['atf-content'];
      }
      catch(err)
      {
                  var n1=html.indexOf('WML_REDUX_INITIAL_STATE')+13+15;
                  var n2=html.indexOf('</script>',n1)-3;
                  var html2=html.slice(n1,n2);
                  jsonData=JSON.parse(html2)
        
      
      }
      
      return jsonData;

}

