

function transferToCommission()
{
      
      var mode="tt"; //put "m" for manual and "tt" for time triggered, when manual it will transfer from active sheet, or transfer from current monthly sheet
      
      var ss=SpreadsheetApp.getActiveSpreadsheet();
     if(mode=="tt"){
        var currentDate=new Date();
        var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
        var sheet=ss.getSheetByName(sheetName);
        var month=currentDate.getMonth();
        
        var values=sheet.getDataRange().getValues();
        var valuesF=sheet.getDataRange().getFormulasR1C1();
        
        var ss2="1xyEObomuO8FM5YvP4BQE0PzQYvFo4mKU6KJIMfcy-cQ"
        var values2=SpreadsheetApp.openById(ss2).getSheetByName(sheetName).getDataRange().getValues();
        var valuesF2=SpreadsheetApp.openById(ss2).getSheetByName(sheetName).getDataRange().getFormulasR1C1();
        
        var ss3="1qteBkdSKNJ8RQ3QBbev0UYx4Zc-8mL12CLkMXVpJrAc";
        var values3=SpreadsheetApp.openById(ss3).getSheetByName(sheetName).getDataRange().getValues();
        var valuesF3=SpreadsheetApp.openById(ss3).getSheetByName(sheetName).getDataRange().getFormulasR1C1();
        
        values=values.concat(values2).concat(values3);
        valuesF=valuesF.concat(valuesF3).concat(valuesF3);
        
     } else {
        var sheet=SpreadsheetApp.getActiveSpreadsheet();
        var values=sheet.getDataRange().getValues();
        var valuesF=sheet.getDataRange().getFormulasR1C1();
        
     }
      
     
     
  
  var shamsArr=[];  //<------------------create a blank array for each of the listers
  var saraArr=[];
  var bradArr=[];
  var reillyArr=[];
  var rohitArr=[];
  var zainabArr=[];
  var jeremyArr=[];
  var gageArr=[];
  var steveArr=[];
  var matthewArr=[];
  var trevorArr=[];
  var marilynArr=[];
     
     for (var i=4; i<values.length; i++){
       
         var initial=values[i][25];
         var supId=values[i][13-1];
         if(supId==""){continue}
         var date=values[i][1-1];
         var orderId=valuesF[i][4-1];
         var asin=valuesF[i][24-1];
         var markUp=values[i][12-1];
         var title=values[i][2-1];

                      
                     
         var initials=initial;
         var salePrice=values[i][9-1];
         var sku=valuesF[i][25-1];
         var finalProfit=values[i][14-1]+values[0][15-1];
         var supId=values[i][13-1];             
         var tempArr=[date, orderId, asin, markUp, "", 0.0, title, initial, salePrice, sku, finalProfit, "", "", supId] 
         
         if(initial.indexOf('RHCU')>=0 && orderId.indexOf('amazon.com')>=0){
             var perc=0.40
             tempArr[4]='='+perc+'*R[0]C[-1]';
             rohitArr.push(tempArr);
         }
         
  
       if(initial.toUpperCase().indexOf('JECU')>=0 && orderId.indexOf('amazon.com')>=0){     
         var perc=0.35             
         tempArr[4]='='+perc+'*R[0]C[-1]';
         jeremyArr.push(tempArr); 
       }
       if(initial.indexOf('Shams')>=0 && orderId.indexOf('amazon.com')>=0){
         var perc=0.40
         tempArr[4]='='+perc+'*R[0]C[-1]';
         shamsArr.push(tempArr);
       }
       if(initial.indexOf('SGCU')>=0 && orderId.indexOf('amazon.com')>=0){
         var perc=0.40
         tempArr[4]='='+perc+'*R[0]C[-1]';
         saraArr.push(tempArr);
       }
       if(initial.indexOf('GLCU')>=0 && orderId.indexOf('amazon.com')>=0){
         var perc=0.40
         tempArr[4]='='+perc+'*R[0]C[-1]';
         gageArr.push(tempArr);
       }
       if(initial.indexOf('BPCU')>=0 && orderId.indexOf('amazon.com')>=0){
         var perc=0.40
         tempArr[4]='='+perc+'*R[0]C[-1]';
         bradArr.push(tempArr);
       }
       if(initial.indexOf('SHCU')>=0 && orderId.indexOf('amazon.com')>=0){
         var perc=0.40
         tempArr[4]='='+perc+'*R[0]C[-1]';
         steveArr.push(tempArr);
       }
       if(initial.indexOf('MGCU')>=0 && orderId.indexOf('amazon.com')>=0){
         var perc=0.40
         tempArr[4]='='+perc+'*R[0]C[-1]';
         matthewArr.push(tempArr);
       }
       if(initial.indexOf('TPCU')>=0 && orderId.indexOf('amazon.com')>=0){
         var perc=0.40
         tempArr[4]='='+perc+'*R[0]C[-1]';
         trevorArr.push(tempArr);
       }
       if(initial.indexOf('MMCU')>=0 && orderId.indexOf('amazon.com')>=0){
         var perc=0.50
         tempArr[4]='='+perc+'*R[0]C[-1]';
         marilynArr.push(tempArr);
       }
       if(initial.indexOf('RNCU')>=0 && orderId.indexOf('amazon.com')>=0){
         var perc=0.40
         tempArr[4]='='+perc+'*R[0]C[-1]';
         reillyArr.push(tempArr);
       }
       if(initial.indexOf('ZACU')>=0 && orderId.indexOf('amazon.com')>=0){
         var perc=0.40
         tempArr[4]='='+perc+'*R[0]C[-1]';
         zainabArr.push(tempArr);
       }

       
       
       
       
     }
     
     
     
     
  
  //---------------add this if block for each lister-----------------//
  var arr=rohitArr;  //<--------------------change array name
  var ss2Id=rohitSheetId; //<-------------change sheet id
  if(arr.length>0){ //
    var ss2=SpreadsheetApp.openById(ss2Id);
    var sheet2=ss2.getSheetByName(sheet.getName());
    if(sheet2==null){
      var templateSheet=ss2.getSheetByName("TEMPLATE");
      sheet2=ss2.insertSheet(sheet.getName(), 2, {template: templateSheet});
      if(sheet2.isSheetHidden()){sheet2.showSheet()}
    }
    
    sheet2.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }//------------end of block for each lister 
  
  
  
  
  //---------------add this if block for each lister-----------------//
  var arr=jeremyArr;  //<--------------------change array name
  var ss2Id=jeremySheetId; //<-------------change sheet id
  if(arr.length>0){ //
    var ss2=SpreadsheetApp.openById(ss2Id);
    var sheet2=ss2.getSheetByName(sheet.getName());
    if(sheet2==null){
      var templateSheet=ss2.getSheetByName("TEMPLATE");
      sheet2=ss2.insertSheet(sheet.getName(), 2,{template: templateSheet});
      if(sheet2.isSheetHidden()){sheet2.showSheet()}
    }
    
    sheet2.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }//------------end of block for each lister 
  
  
  
  var arr=shamsArr; 
  var ss2Id=shamsSheetId; 
  if(arr.length>0){ //
    var ss2=SpreadsheetApp.openById(ss2Id);
    var sheet2=ss2.getSheetByName(sheet.getName());
    if(sheet2==null){
      var templateSheet=ss2.getSheetByName("TEMPLATE");
      sheet2=ss2.insertSheet(sheet.getName(), 2, {template: templateSheet});
      if(sheet2.isSheetHidden()){sheet2.showSheet()}
    }
    
    sheet2.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }
  
  
  var arr=bradArr; 
  var ss2Id=bradSheetId; 
  if(arr.length>0){ //
    var ss2=SpreadsheetApp.openById(ss2Id);
    var sheet2=ss2.getSheetByName(sheet.getName());
    if(sheet2==null){
      var templateSheet=ss2.getSheetByName("TEMPLATE");
      sheet2=ss2.insertSheet(sheet.getName(), 3, {template: templateSheet});
      if(sheet2.isSheetHidden()){sheet2.showSheet()}
    }
    
    sheet2.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }
  
  
  var arr=saraArr; 
  var ss2Id=saraSheetId; 
  if(arr.length>0){ //
    var ss2=SpreadsheetApp.openById(ss2Id);
    var sheet2=ss2.getSheetByName(sheet.getName());
    if(sheet2==null){
      var templateSheet=ss2.getSheetByName("TEMPLATE");
      sheet2=ss2.insertSheet(sheet.getName(), 2,{template: templateSheet});
      if(sheet2.isSheetHidden()){sheet2.showSheet()}
    }
    
    sheet2.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }
  
  
  
  var arr=matthewArr; 
  var ss2Id=matthewSheetId; 
  if(arr.length>0){ //
    var ss2=SpreadsheetApp.openById(ss2Id);
    var sheet2=ss2.getSheetByName(sheet.getName());
    if(sheet2==null){
      var templateSheet=ss2.getSheetByName("TEMPLATE");
      sheet2=ss2.insertSheet(sheet.getName(), 2, {template: templateSheet});
      if(sheet2.isSheetHidden()){sheet2.showSheet()}
    }
    
    sheet2.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }
  
  
  
  var arr=gageArr; 
  var ss2Id=gageSheetId; 
  if(arr.length>0){ //
    var ss2=SpreadsheetApp.openById(ss2Id);
    var sheet2=ss2.getSheetByName(sheet.getName());
    if(sheet2==null){
      var templateSheet=ss2.getSheetByName("TEMPLATE");
      sheet2=ss2.insertSheet(sheet.getName(), 2, {template: templateSheet});
      if(sheet2.isSheetHidden()){sheet2.showSheet()}
    }
    
    sheet2.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }
  
  
  
  var arr=trevorArr; 
  var ss2Id=trevSheetId; 
  if(arr.length>0){ //
    var ss2=SpreadsheetApp.openById(ss2Id);
    var sheet2=ss2.getSheetByName(sheet.getName());
    if(sheet2==null){
      var templateSheet=ss2.getSheetByName("TEMPLATE");
      sheet2=ss2.insertSheet(sheet.getName(), 2, {template: templateSheet});
      if(sheet2.isSheetHidden()){sheet2.showSheet()}
    }
    
    sheet2.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }
  
  
  
  var arr=marilynArr; 
  var ss2Id=mmSheetId; 
  if(arr.length>0){ //
    var ss2=SpreadsheetApp.openById(ss2Id);
    var sheet2=ss2.getSheetByName(sheet.getName());
    if(sheet2==null){
      var templateSheet=ss2.getSheetByName("TEMPLATE");
      sheet2=ss2.insertSheet(sheet.getName(), 2, {template: templateSheet});
      if(sheet2.isSheetHidden()){sheet2.showSheet()}
    }
    
    sheet2.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }
  
  
  
  var arr=steveArr; 
  var ss2Id=steveSheetId; 
  if(arr.length>0){ //
    var ss2=SpreadsheetApp.openById(ss2Id);
    var sheet2=ss2.getSheetByName(sheet.getName());
    if(sheet2==null){
      var templateSheet=ss2.getSheetByName("TEMPLATE");
      sheet2=ss2.insertSheet(sheet.getName(), 3, {template: templateSheet});
      if(sheet2.isSheetHidden()){sheet2.showSheet()}
    }
    
    sheet2.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }
  
  
  
  var arr=reillyArr; 
  var ss2Id=reillySheetId; 
  if(arr.length>0){ //
    var ss2=SpreadsheetApp.openById(ss2Id);
    var sheet2=ss2.getSheetByName(sheet.getName());
    if(sheet2==null){
      var templateSheet=ss2.getSheetByName("TEMPLATE");
      sheet2=ss2.insertSheet(sheet.getName(), 3, {template: templateSheet});
      if(sheet2.isSheetHidden()){sheet2.showSheet()}
    }
    
    sheet2.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }
  
  
  
  var arr=zainabArr; 
  var ss2Id=zainabSheetId; 
  if(arr.length>0){ //
    var ss2=SpreadsheetApp.openById(ss2Id);
    var sheet2=ss2.getSheetByName(sheet.getName());
    if(sheet2==null){
      var templateSheet=ss2.getSheetByName("TEMPLATE");
      sheet2=ss2.insertSheet(sheet.getName(), 2, {template: templateSheet});
      if(sheet2.isSheetHidden()){sheet2.showSheet()}
    }
    
    sheet2.getRange(2, 1, arr.length, arr[0].length).setValues(arr)
  }
  
  
  
      
      
      


}

