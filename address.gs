


function clearpofr(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  
  var currentDate=new Date();
  
  var current = Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  
  var sheet = ss.getSheetByName(current);
  
  
  
  sheet.getRange("AN100:AN").clearContent();
  
}
/*

function updateOrderList(){
  var api_url = getApiURL("");
  if(api_url !=""){
    
    try{
      var responseXML=UrlFetchApp.fetch(api_url).getContentText();
      
      var responseXMLObj = XML_to_JSON(responseXML);
      var orders = responseXMLObj.ListOrdersResponse.ListOrdersResult.Orders.Order;
      if(orders.length > 0){
        var orderIDData = [];
        var buyerEmailData = [];
        var buyerNameData = [];
        var phoneData = [];
        
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheetInv = ss.getSheetByName("Data");
        var values = sheetInv.getDataRange().getValues();
        
        for(var c=0; c < orders.length; c++){
          var order = orders[c];
          var orderID = order.AmazonOrderId.Text;
          var orderExists = false;
          for(var i=2; i<values.length; i++){
            var order_id=values[i][0];
            if(order_id == orderID){
              orderExists = true;
            }
          }
          if(orderExists) continue;
          
          var buyerEmail = "";
          var buyerName = "";
          var buyerPhone = "";
          if(order.ShippingAddress && order.ShippingAddress.AddressLine1){
            buyerEmail = order.ShippingAddress.AddressLine1.Text;
          }
          if(order.ShippingAddress && order.ShippingAddress.Name){
            buyerName = order.ShippingAddress.Name.Text;
          }
          if(order.ShippingAddress && order.ShippingAddress.AddressLine2){
            buyerPhone = order.ShippingAddress.AddressLine2.Text;
          }
          
          orderIDData.push([orderID]);
          buyerEmailData.push([buyerEmail]);
          buyerNameData.push([buyerName]);
          phoneData.push([buyerPhone]);
        }
       
        var rownum=last_row(sheetInv,6);
        var pobox='=if(R[0]C[-4]>0,if(OR(NOT(ISERR(SEARCH("PO ",R[0]C[-3]))),NOT(ISERR(SEARCH("*BOX*",R[0]C[-3]))),NOT(ISERR(SEARCH("*BOX*",R[0]C[-1]))),NOT(ISERR(SEARCH("PO ",R[0]C[-1]))),NOT(ISERR(SEARCH("*BOX*",R[0]C[-2]))),NOT(ISERR(SEARCH("PO ",R[0]C[-2])))),"PO",""),"")';
        var freight='=if(R[0]C[-5]>0,if(isnumber(SPLIT(LOWER(R[0]C[-3]) ,"abcdefghijklmnopqrstuvwxyz.-/\() ")),"FR",if(R[0]C[-5]>0,iferror(if(len(SPLIT(LOWER(R[0]C[-2]) ,"abcdefghijklmnopqrstuvwxyz.-/\() "))>4,"FR","")))))';

       var freight2='=if(R[0]C[-6]>0,if(OR(NOT(ISERR(SEARCH("*Ship*",R[0]C[-5]))),NOT(ISERR(SEARCH("Logistic*",R[0]C[-5]))),NOT(ISERR(SEARCH("*Freight*",R[0]C[-5]))),NOT(ISERR(SEARCH("FR",R[0]C[-1]))),NOT(ISERR(SEARCH("Freight",R[0]C[-3]))),NOT(ISERR(SEARCH("*Logistics*",R[0]C[-3]))),NOT(ISERR(SEARCH("*Ship*",R[0]C[-3]))),NOT(ISERR(SEARCH("Freight",R[0]C[-4]))),NOT(ISERR(SEARCH("*Logistics*",R[0]C[-4]))),NOT(ISERR(SEARCH("*Ship*",R[0]C[-4])))),"FR",""),"")';
        
        
        if(orderIDData.length > 0){
          sheetInv.getRange(rownum,6,orderIDData.length).setValues(orderIDData);
          sheetInv.getRange(rownum,7,buyerEmailData.length).setValues(buyerEmailData);
          sheetInv.getRange(rownum,8,buyerNameData.length).setValues(buyerNameData);
          sheetInv.getRange(rownum,9,phoneData.length).setValues(phoneData);
          sheetInv.getRange(rownum,10,phoneData.length).setValue(pobox);
          sheetInv.getRange(rownum,11,phoneData.length).setValue(freight);
          sheetInv.getRange(rownum,12,phoneData.length).setValue(freight2);
          //Logger.log(rownum);
        }
      }
      
    }
    catch(e){
      Logger.log(e);
      return false;
    }
  }




    // deletes it's own trigger after finishing
                  var allTriggers = ScriptApp.getProjectTriggers();
                  for (var i = 0; i < allTriggers.length; i++) 
                  {
                                var trigger=allTriggers[i];
                                var name=trigger.getHandlerFunction();
                                if(name=='updateOrderList')
                                {
                                    ScriptApp.deleteTrigger(allTriggers[i])
                                }
                  }
  
  
  
               var triggerDay = new Date();
               triggerDay.setHours(0)
               triggerDay.setMinutes(0)
               triggerDay.setSeconds(0); //mid nigh time
               
               triggerDay.setTime(triggerDay.getTime()+24*60*60*1000 +  8*60*60*1000 + 30*60*1000); //7AM next day
               ScriptApp.newTrigger("updateOrderList")
               .timeBased()
               .at(triggerDay)
               .create();



  var ss5=SpreadsheetApp.getActiveSpreadsheet();
  var statussheet=ss5.getSheetByName("InStock");
  var newdate = new Date();

  statussheet.getRange(2, 34).setValue(newdate);
  
}

function updateBlankOrderData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInv = ss.getSheetByName("Data");
  var values = sheetInv.getDataRange().getValues();
  var orderIDs = {};
  var boc=1;
  for(var i=2; i<values.length; i++){
    var order_id=values[i][0];
    if(order_id){
      var email=values[i][1];
      var name=values[i][2];
      var phone=values[i][3];
      if(!email || !name || !phone){
        orderIDs["AmazonOrderId.Id."+boc]=order_id;
        if(boc>=1){
          var api_url = getApiURL(orderIDs);
          updateBlankOrderRows(api_url);
          orderIDs = {};
          boc=0;
          return;
        }
        boc++;
      }
    }
  }
  
  if(Object.getOwnPropertyNames(orderIDs).length > 0){
    var api_url = getApiURL(orderIDs);
    updateBlankOrderRows(api_url);
  }
}
function updateBlankOrderRows(api_url){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInv = ss.getSheetByName("Data");
  var values = sheetInv.getDataRange().getValues();
  try{
    var responseXML=UrlFetchApp.fetch(api_url).getContentText();
    var responseXMLObj = XML_to_JSON(responseXML);
    var orders = responseXMLObj.GetOrderResponse.GetOrderResult.Orders.Order;
    if(orders){
      
      var buyerEmailData = [];
      var buyerNameData = [];
      var phoneData = [];
      for(var i=2; i<values.length; i++){
        var order_id=values[i][0];
        var email=values[i][1];
        var name=values[i][2];
        var phone=values[i][3];
        var orderFound = false;
        
        var amzOrderID = orders.AmazonOrderId;
        
        if(amzOrderID !== "" && amzOrderID !== null && amzOrderID !== undefined){
          var orderID = amzOrderID.Text;
          if(order_id == orderID){
            
            var order = orders;
          if(order.ShippingAddress && order.ShippingAddress.AddressLine1){
            buyerEmail = order.ShippingAddress.AddressLine1.Text;
          }
          if(order.BuyerEmail){
            buyerName = order.BuyerName.Text;
          }
          if(order.ShippingAddress && order.ShippingAddress.AddressLine2){
            buyerPhone = order.ShippingAddress.AddressLine2.Text;
            }
              
            buyerEmailData.push([email]);
            buyerNameData.push([name]);
            phoneData.push([phone]);
            orderFound = true;
          }
        }
        else{
          for(var c=0; c < orders.length; c++){
            var order = orders[c];
            var orderID = order.AmazonOrderId.Text;
            if(order_id == orderID){
              if(order.BuyerEmail){
                email = order.BuyerEmail.Text;
              }
              if(order.BuyerEmail){
                name = order.BuyerName.Text;
              }
              if(order.ShippingAddress && order.ShippingAddress.Phone){
                phone = order.ShippingAddress.Phone.Text;
              }
              
              buyerEmailData.push([email]);
              buyerNameData.push([name]);
              phoneData.push([phone]);
              orderFound = true;
            }
          }
        }
        if(orderFound == false){
          buyerEmailData.push([email]);
          buyerNameData.push([name]);
          phoneData.push([phone]);
        }
      }
      sheetInv.getRange(3,7,buyerEmailData.length).setValues(buyerEmailData);
      sheetInv.getRange(3,8,buyerNameData.length).setValues(buyerNameData);
      sheetInv.getRange(3,9,phoneData.length).setValues(phoneData);
    }
  }
  catch(e){
    Logger.log(e);
    return false;
  }  
}


function updateLastCheckTime(checktime){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInv = ss.getSheetByName("Data");
  var checktimedata=[];
  checktimedata.push([checktime]);
  sheetInv.getRange(1,1,checktimedata.length).setValues(checktimedata);
}
function getApiURL(orderIDs){
  var now = new Date();
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  var dateObj = new Date(now.getTime() - 1 * MILLIS_PER_DAY);
  var lastCheckDate = Utilities.formatDate(dateObj,"GMT","MMMM d, yyyy HH:mm:ss");
  var createdAfter = Utilities.formatDate(dateObj,"GMT","yyyy-MM-dd")+"T"+Utilities.formatDate(dateObj,"GMT","HH:mm:ss")+".000Z";
  var timeStamp = Utilities.formatDate(now,"GMT","yyyy-MM-dd")+"T"+Utilities.formatDate(now,"GMT","HH:mm:ss")+".000Z";
  
  var awsAccessKeyId = "AKIAIUI2RB3WAHEHMVUQ";
  var sellerId = "A2KTDJV6EUITJE";
  var marketplaceId = "ATVPDKIKX0DER";
  var secretAccessKey = "y0mQiaFzq9o2WxFSWB2pajfljdizB028/7ns1MFx";
  var action = "ListOrders";
  var payload = {};
  if(orderIDs!="" && Object.getOwnPropertyNames(orderIDs).length > 0){
    action = "GetOrder";
    payload["AWSAccessKeyId"]=awsAccessKeyId;
    payload["Action"]=action;
    for(var key in orderIDs){
      payload[key]=orderIDs[key];
    }
    payload["MarketplaceId.Id.1"]=marketplaceId;
    payload["SellerId"]=sellerId;
    payload["SignatureMethod"]="HmacSHA256";
    payload["SignatureVersion"]="2";
    payload["Timestamp"]=timeStamp;
    payload["Version"]="2013-09-01";
    
  }
  else{
    payload ={
      "AWSAccessKeyId":awsAccessKeyId,
      "Action":action,
      "CreatedAfter":createdAfter,
      "MarketplaceId.Id.1":marketplaceId,
      "SellerId": sellerId,
      "SignatureMethod":"HmacSHA256",
      "SignatureVersion":"2",
      "Timestamp":timeStamp,
      "Version":"2013-09-01"
    };
  }
  
  var stringToSign="";
  for(var key in payload){
    if(stringToSign!==""){
      stringToSign=stringToSign+"&";
    }
    stringToSign=stringToSign+key+"="+encodeURIComponent(payload[key]);
  }
  var parms=stringToSign;
  stringToSign = "GET\nmws.amazonservices.com\n/Orders/2013-09-01\n"+stringToSign;
  //var signature = Utilities.computeHmacSha256Signature(stringToSign,"y0mQiaFzq9o2WxFSWB2pajfljdizB028/7ns1MFx");
  var signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256,stringToSign,secretAccessKey);
  var blob = Utilities.newBlob(signature);
  var encoded = Utilities.base64Encode(blob.getBytes());
  var api_url="https://mws.amazonservices.com/Orders/2013-09-01?"+parms+"&Signature="+encodeURIComponent(encoded);
  if(orderIDs =="" || Object.getOwnPropertyNames(orderIDs).length === 0){
    updateLastCheckTime(lastCheckDate);
  }
  //Logger.log(api_url);
  return api_url;
  
}


function XML_to_JSON(xml) { 
  var doc = XmlService.parse(xml);
  var result = {};
  var root = doc.getRootElement();
  result[root.getName()] = elementToJSON(root);
  return result;
}
function elementToJSON(element) {
  var result = {};
  // Attributes.
  element.getAttributes().forEach(function(attribute) {
    result[attribute.getName()] = attribute.getValue();
  });
  // Child elements.
  element.getChildren().forEach(function(child) {
    var key = child.getName();
    var value = elementToJSON(child);
    if (result[key]) {
      if (!(result[key] instanceof Array)) {
        result[key] = [result[key]];
      }
      result[key].push(value);
    } else {
      result[key] = value;
    }
  });
  // Text content.
  if (element.getText()) {
    result['Text'] = element.getText();
  }
  return result;
}
*/
