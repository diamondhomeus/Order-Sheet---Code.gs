


function updateOrderIds(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Compare");
  var values = sheet.getDataRange().getValues();


 sheet.getRange("A5:C").clearContent();

 var currentDate = new Date();
  var currentmonth=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM");
  var pastDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1);
  var pastmonth=Utilities.formatDate(pastDate, ss.getSpreadsheetTimeZone(), "MMMM");

  var indexmatch = "=iferror(index('"+currentmonth+"-2018'!R2C4:C4,match(R[0]C[-2],'"+currentmonth+"-2018'!R2C4:C4,0)),iferror(index('Cancelled Orders'!R2C4:C4,match(R[0]C[-2],'Cancelled Orders'!R2C4:C4,0)),iferror(index('"+pastmonth+"-2018'!R2C4:C4,match(R[0]C[-2],'"+pastmonth+"-2018'!R2C4:C4,0)))))";
 


  Logger.log(values.length);

  
  var daysBack = 300;
  for(var key in values){
    if(values[key][0] == "DAYS"){
      daysBack = parseInt(values[key][1]);
      break;
    }
  }
  
  try{
    var orderIDs = [];
    var sellerIDs = [];
    var api_url = getApiURL("ListOrders","",daysBack);
    var response,responseArr,orders;
    var fetchOrder = true;
    var requestCount = 0;
    while(fetchOrder){
      //if(fetchOrder){
      response = UrlFetchApp.fetch(api_url,{'method' : 'get','muteHttpExceptions': true}).getContentText();
      requestCount++;
      if(response.indexOf("Request is throttled") !== -1){
        Utilities.sleep(120000);
      }
      else{
        fetchOrder = false;
      }
      
      if(response){
        responseArr = XML_to_JSON(response);
        
        if(responseArr.ListOrdersResponse != undefined && responseArr.ListOrdersResponse.ListOrdersResult != undefined){
          responseArr = responseArr.ListOrdersResponse.ListOrdersResult;
          if(responseArr.Orders != undefined && responseArr.Orders.Order != undefined){
            orders = responseArr.Orders.Order;
          }
        }
        else if(responseArr.ListOrdersByNextTokenResponse != undefined && responseArr.ListOrdersByNextTokenResponse.ListOrdersByNextTokenResult != undefined){
          responseArr = responseArr.ListOrdersByNextTokenResponse.ListOrdersByNextTokenResult;
          if(responseArr.Orders != undefined && responseArr.Orders.Order != undefined){
            orders = responseArr.Orders.Order;
          }
        }
        
       
        if(orders){
          //if(responseArr.ListOrdersResult.Orders.Order != undefined){
            //orders = responseArr.ListOrdersResult.Orders.Order;
            for(var key in orders){
              var amazon_order_id = orders[key].AmazonOrderId.Text;
              if(amazon_order_id.indexOf("7") === 0){
                continue;
              }
              var seller_order_id = "";
              if(orders[key].SellerOrderId != undefined){
                seller_order_id = orders[key].SellerOrderId.Text;
              }
              orderIDs.push([amazon_order_id]);
              sellerIDs.push([seller_order_id]);
            }
            if(responseArr.NextToken != undefined){
              var nextToken = responseArr.NextToken.Text;
              api_url = getApiURL("ListOrdersByNextToken",nextToken,daysBack);
              fetchOrder = true;
            }
          //}
          if(requestCount >= 5){
            Utilities.sleep(60000);
            requestCount = 0;
          }
        }
        else{
          
        }
      }
      
      //}
    }
  }
  catch(e){
    Logger.log(e);
  }
  
  if(orderIDs){
    Logger.log("Order id: "+orderIDs.length);
    sheet.getRange(5,1,orderIDs.length).setValues(orderIDs);
    sheet.getRange(5,2,sellerIDs.length).setValues(sellerIDs);
    sheet.getRange(5,3,orderIDs.length).setFormula(indexmatch);
  }

var headerRows = 5;

 var sortFirst = 3; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
    var sortFirstAsc = false; //Set to false to sort descending

 var range = sheet.getRange(headerRows, 1, sheet.getMaxRows()-headerRows, sheet.getLastColumn());

 range.sort([ {column: sortFirst, ascending: sortFirstAsc}]);


}






function updateOrderIdsReturn(){
  var status = getLastReportStatus();
  if(status["op"] == "request_report"){
    var api_url = getApiURLReport("RequestReport","");
    if(api_url !=""){
      try{
        var options = {'method' : 'post','muteHttpExceptions': true};
        var responseXML=UrlFetchApp.fetch(api_url,options).getContentText();
        if(responseXML){
          var responseXMLObj = XML_to_JSON(responseXML);
          
          if(responseXMLObj.RequestReportResponse.RequestReportResult != undefined){
            var reportProcessingStatus = responseXMLObj.RequestReportResponse.RequestReportResult.ReportRequestInfo.ReportProcessingStatus.Text;
            if(reportProcessingStatus == "_SUBMITTED_"){
              var reportRequestId = responseXMLObj.RequestReportResponse.RequestReportResult.ReportRequestInfo.ReportRequestId.Text;
              updateRequestData(reportRequestId,"",reportProcessingStatus);
            }         
          }
        }
      }
      catch(e){
        Logger.log(e);
        return false;
      }
    }
  }
  else if(status["op"] == "get_report_request_list"){
    var request_id = status["request_id"];
    var api_url = getApiURLReport("GetReportRequestList",request_id);
    if(api_url !=""){
      try{
        var options = {'method' : 'post','muteHttpExceptions': true};
        var responseXML=UrlFetchApp.fetch(api_url,options).getContentText();
        if(responseXML){
          var responseXMLObj = XML_to_JSON(responseXML);
          Logger.log(responseXMLObj);
          if(responseXMLObj.GetReportRequestListResponse.GetReportRequestListResult != undefined){
            var reportProcessingStatus = responseXMLObj.GetReportRequestListResponse.GetReportRequestListResult.ReportRequestInfo.ReportProcessingStatus.Text;
            if(reportProcessingStatus == "_DONE_"){
              var reportId = responseXMLObj.GetReportRequestListResponse.GetReportRequestListResult.ReportRequestInfo.GeneratedReportId.Text;
              var reportRequestId = responseXMLObj.GetReportRequestListResponse.GetReportRequestListResult.ReportRequestInfo.ReportRequestId.Text;
              var api_url_report = getApiURLReport("GetReport",reportId);
              if(api_url_report != ""){
                var options = {'method' : 'post','muteHttpExceptions': true};
                var responseData=UrlFetchApp.fetch(api_url_report,options).getContentText();
                if(responseData){
                  var lineArray = responseData.split("\n");
                  if(lineArray){
                    var orderIDData=[];
                    var sellerIDData=[];
                    var i=0;
                    for(var key in lineArray){
                      if(i == 0){
                        i++;
                        continue;
                      }
                      var row = lineArray[key].split("\t");
                      Logger.log(row);
                      return;
                      if(row[0] != undefined && row[1] != undefined ){
                        var sku = row[0].toLowerCase();
                        if(sku.indexOf("dhd") === -1){
                          orderIDData.push([row[0]]);
                          sellerIDData.push([row[1]]);
                        }
                      }
                      //if(i>50){
                        //break;
                      //}
                      i++;
                    }
                    
                    updateRowsData(orderIDData,sellerIDData,reportRequestId,reportId);
                    
                    /*
                    updateRequestData(reportRequestId,reportId,"updating...");
                    var ss = SpreadsheetApp.getActiveSpreadsheet();
                    var sheetInv = ss.getActiveSheet();
                    sheetInv.getRange(3,1,skuData.length).setValues(skuData);
                    sheetInv.getRange(3,2,asinData.length).setValues(asinData);
                    sheetInv.getRange(3,3,stockData.length).setValues(stockData);
                    updateRequestData(reportRequestId,reportId,reportProcessingStatus);*/
                  }
                  
                  
                }
                
              }
              
              //updateRequestData(reportRequestId,reportId,reportProcessingStatus);
              
            }
            else if(reportProcessingStatus == "_DONE_NO_DATA_"){
              var reportId = responseXMLObj.GetReportRequestListResponse.GetReportRequestListResult.ReportRequestInfo.GeneratedReportId.Text;
              var reportRequestId = responseXMLObj.GetReportRequestListResponse.GetReportRequestListResult.ReportRequestInfo.ReportRequestId.Text;
              updateRequestData(reportRequestId,reportId,reportProcessingStatus);
            }
          }
        }
        
      }
      catch(e){
        Logger.log(e);
        return false;
      }
    }
  } 
}



function updateRowsData(orderIDArr,sellerIDArr,reportRequestId,reportId){
  updateRequestData(reportRequestId,reportId,"updating...");
     
  updateRequestData(reportRequestId,reportId,"DONE");
}

function getLastReportStatus(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reports");
  var values = sheet.getDataRange().getValues();
  
  var requestId=values[0][0];
  var reportId=values[0][1];
  var reportStatus=values[0][2];
  var status = {};
  if((reportStatus == "" && requestId == "" && reportId == "") || (reportStatus=="_DONE_" || reportStatus=="_DONE_NO_DATA_")){
    status["op"] = "request_report";
  }
  else if((requestId != "") && (reportStatus=="_IN_PROGRESS_" || reportStatus == "_SUBMITTED_")){
    status["op"] = "get_report_request_list";
    status["request_id"] = requestId;
  }
  else{
    status["op"] = "request_report";
  }
  return status;
}

function updateRequestData(request_id,report_id,status){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reports");
  var data1=[],data2=[],data3=[];
  data1.push([request_id]);
  data2.push([report_id]);
  data3.push([status]);
  sheet.getRange(1,1,data1.length).setValues(data1);
  sheet.getRange(1,2,data2.length).setValues(data2);
  sheet.getRange(1,3,data3.length).setValues(data3);
}



function getApiURL(action,data,daysBack){
  
  if(!daysBack){
    daysBack = 300;
  }
  
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  
  var now = new Date();
  var timeStamp = Utilities.formatDate(now,"GMT","yyyy-MM-dd")+"T"+Utilities.formatDate(now,"GMT","HH:mm:ss")+".000Z";
  var awsAccessKeyId = "AKIAIUI2RB3WAHEHMVUQ";
  var sellerId = "A2KTDJV6EUITJE";
  var marketplaceId = "ATVPDKIKX0DER";
  var secretAccessKey = "y0mQiaFzq9o2WxFSWB2pajfljdizB028/7ns1MFx";
  
  var dateObj = new Date(now.getTime() - daysBack * MILLIS_PER_DAY);
  var createdAfter = Utilities.formatDate(dateObj,"GMT","yyyy-MM-dd")+"T"+Utilities.formatDate(dateObj,"GMT","HH:mm:ss")+".000Z";
  
  var payload = {};
  payload["AWSAccessKeyId"]=awsAccessKeyId;
  payload["Action"] = action;  
  if(action == "ListOrdersByNextToken"){
    payload["NextToken"] = data;
  }
  else{
    payload["CreatedAfter"] = createdAfter;
    payload["MarketplaceId.Id.1"] = marketplaceId;
    payload["OrderStatus.Status.1"] = "Unshipped";
    payload["OrderStatus.Status.2"] = "PartiallyShipped";
  }
  payload["SellerId"]=sellerId;
  payload["SignatureMethod"]="HmacSHA256";
  payload["SignatureVersion"]="2";
  payload["Timestamp"]=timeStamp;
  payload["Version"]="2013-09-01";
  
  var stringToSign="";
  for(var key in payload){
    if(stringToSign!==""){
      stringToSign=stringToSign+"&";
    }
    stringToSign=stringToSign+key+"="+encodeURIComponent(payload[key]);
  }
  var parms=stringToSign;
  stringToSign = "GET\nmws.amazonservices.com\n/Orders/2013-09-01\n"+stringToSign;
  var signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256,stringToSign,secretAccessKey);
  var blob = Utilities.newBlob(signature);
  var encoded = Utilities.base64Encode(blob.getBytes());
  var api_url="https://mws.amazonservices.com/Orders/2013-09-01?"+parms+"&Signature="+encodeURIComponent(encoded);
  return api_url;
}


function getApiURLReport(action,data){
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
  var payload = {};
  payload["AWSAccessKeyId"]=awsAccessKeyId;
  payload["Action"]=action;
  payload["Marketplace"]=marketplaceId;
  if(action == "RequestReport"){
    payload["ReportType"]="_GET_FLAT_FILE_RETURNS_DATA_BY_RETURN_DATE_";
  }
  else if(action == "GetReportRequestList"){
    payload["ReportRequestIdList.Id.1"]=data;
  }
  else if(action == "GetReport"){
    payload["ReportId"]=data;
  }
  payload["SellerId"]=sellerId;
  payload["SignatureMethod"]="HmacSHA256";
  payload["SignatureVersion"]="2";
  payload["Timestamp"]=timeStamp;
  payload["Version"]="2009-01-01";
  
  var stringToSign="";
  //var requestParms={};
  for(var key in payload){
    if(stringToSign!==""){
      stringToSign=stringToSign+"&";
    }
    stringToSign=stringToSign+key+"="+encodeURIComponent(payload[key]);
    //requestParms[key]=encodeURIComponent(payload[key]);
  }
  var parms=stringToSign;
  stringToSign = "POST\nmws.amazonservices.com\n/Reports/2009-01-01\n"+stringToSign;
  var signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256,stringToSign,secretAccessKey);
  var blob = Utilities.newBlob(signature);
  var encoded = Utilities.base64Encode(blob.getBytes());
  var api_url="https://mws.amazonservices.com/Reports/2009-01-01?"+parms+"&Signature="+encodeURIComponent(encoded);
  return api_url;
}


/* functions for xml to json  */
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
/* functions for xml to json  */


 