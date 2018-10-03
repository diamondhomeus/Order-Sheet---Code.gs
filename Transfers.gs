
function transfertoLS ()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  var currentDate = new Date();
  
  var Ded = SpreadsheetApp.openById(resSsId);
  var sheetD=Ded.getSheetByName("LS");
  
  var lr= last_row(sheetD, 4);
  
  var orderIDv = sheet.getRange(row, 4).getFormula();
  
  
  var orderId=sheet.getRange(row,4).getValue();
  orderId=orderId.toString();
  
  var MSG='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+orderId+'","MSG")';
  
  
  
  var trackingInfo=trackingCalc(supId);
  var trackingNo=trackingInfo[0];
 
  
  
  
  var profit=sheet.getRange(row, 12).getValue();
  var loss=0-profit;
  
  var supId=sheet.getRange(row, 13).getValue();
  var prodTitle=sheet.getRange(row, 2).getValue();
  
  var values1=[loss,supId,prodTitle]
  var values2=sheet.getRange(row, 1, 1,27).getValues();
  
  var frm=sheet.getRange(row, 7).getFormula();
  values2[0][7-1]=frm;
  
  var rtrns=values1.concat(values2[0]);

var cog=sheet.getRange(row, 11).getValue();
  
   
  
  var ASINhyp='=HYPERLINK("https://www.amazon.com/gp/product/'+rtrns[24-1+3]+'","'+rtrns[24-1+3]+'")';
  var SKUhyp='=HYPERLINK("https://catalog.amazon.com/abis/edit/RelistProduct.amzn?marketplaceID=ATVPDKIKX0DER&ref=xx_myirelis_cont_myifba&sku='+rtrns[25-1+3]+'&asin='+rtrns[24-1+3]+'&productType=HOME","'+rtrns[25-1+3]+'")';
  
  

  
  
  sheetD.getRange(lr+1,1).setValue(currentDate);
  sheetD.getRange(lr+1,2).setValue("");
  sheetD.getRange(lr+1,3).setValue("PENDING");
  sheetD.getRange(lr+1,4).setFormula(orderIDv);
  
  sheetD.getRange(lr+1,5).setValue(addHypToSupId(supId)) 
  
  sheetD.getRange(lr+1,7).setFormula(MSG);
  sheetD.getRange(lr+1, 20).setFormula(trackingNo)
  sheetD.getRange(lr+1, 21).setValue(cog);
  sheetD.getRange(lr+1, 22).setValue(loss);
  
  Browser.msgBox("Transferred to LS!");    
}  




function transfertoADV ()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  var currentDate = new Date();
  
  var Ded = SpreadsheetApp.openById(resSsId);
  var sheetD=Ded.getSheetByName("ADV");
  
  var lr= last_row(sheetD, 4);
  
  var orderIDv = sheet.getRange(row, 4).getFormula();
  var supId=sheet.getRange(row,13).getValue();
  
  
  
  var orderId=sheet.getRange(row,4).getValue();
  orderId=orderId.toString();
  
  var MSG='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+orderId+'","MSG")';
  
  
  
  var trackingInfo=trackingCalc(supId);
  var trackingNo=trackingInfo[0];
  
  
  
  sheetD.getRange(lr+1,1).setValue(currentDate);
  sheetD.getRange(lr+1,2).setValue("");
  sheetD.getRange(lr+1,3).setValue("PENDING");
  sheetD.getRange(lr+1,4).setFormula(orderIDv); 
  sheetD.getRange(lr+1,5).setValue(addHypToSupId(supId))
  
  sheetD.getRange(lr+1,7).setFormula(MSG);
  sheetD.getRange(lr+1, 13).setFormula(trackingNo)
  
  
  Browser.msgBox("Transferred to ADV!");    
}  






function transfertoReturns ()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  var currentDate = new Date();
  
  var ColAH ='=iferror(SUBSTITUTE("Hello "&iferror(index(Names!R2C2:C2,match(R[0]C[-32],Names!R2C1:C1,0)),"XXCUSTOMERXX")&","&index(AutoDrafts!R1C3:C3,match(if(ISERR(SEARCH("/L",R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])-2))&if(OR(NOT(ISERR(SEARCH("Replacement",R[0]C[-30]))),NOT(ISERR(SEARCH("RP",R[0]C[-30])))),"RP","")&if(OR(NOT(ISERR(SEARCH("Exchange",R[0]C[-30]))),NOT(ISERR(SEARCH("EX",R[0]C[-30])))),"EX",""),AutoDrafts!R1C1:C1,0))&R[0]C[-31]&index(AutoDrafts!R1C4:C4,match(if(ISERR(SEARCH("/L",R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])-2))&if(OR(NOT(ISERR(SEARCH("Replacement",R[0]C[-30]))),NOT(ISERR(SEARCH("RP",R[0]C[-30])))),"RP","")&if(OR(NOT(ISERR(SEARCH("Exchange",R[0]C[-30]))),NOT(ISERR(SEARCH("EX",R[0]C[-30])))),"EX",""),AutoDrafts!R1C1:C1,0))&if(isnumber(R[0]C[-23]),text(R[0]C[-23],"#.##"),"")&index(AutoDrafts!R1C5:C5,match(if(ISERR(SEARCH("/L",R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])-2))&if(OR(NOT(ISERR(SEARCH("Replacement",R[0]C[-30]))),NOT(ISERR(SEARCH("RP",R[0]C[-30])))),"RP","")&if(OR(NOT(ISERR(SEARCH("Exchange",R[0]C[-30]))),NOT(ISERR(SEARCH("EX",R[0]C[-30])))),"EX",""),AutoDrafts!R1C1:C1,0)),CHAR(10),CHAR(13)))';
  var ColAJ ='=if(exact(R[0]C[-33],R[0]C[1]),"","N/A")';
  var ColAK ='=iferror(index(SHIPBY!R1C2:C2,match(R[0]C[-35],SHIPBY!R1C1:C1,0)),iferror(index(SHIPBY!R1C5:C5,match(R[0]C[-35],SHIPBY!R1C4:C4,0)),iferror(index(SHIPBY!R1C8:C8,match(R[0]C[-35],SHIPBY!R1C7:C7,0)),iferror(index(SHIPBY!R1C11:C11,match(R[0]C[-35],SHIPBY!R1C10:C10,0))))))';
  var Authdata='=R[0]C[-32]&" "&IF(len(R[0]C[-26])=12,"FedEx",IF(len(R[0]C[-26])=15,"FedEx",IF(len(R[0]C[-26])=18,"UPS",IF(len(R[0]C[-26])=22,"FedEx SmartPost",IF(len(R[0]C[-26])=20,"FedEx","")))))&" "&R[0]C[-26]&" "&TEXT(R[0]C[-24],"#.##")&if(right(R[0]C[-24],2)="00","00","")';
  
  
  
  var arr=[];
  
  arr[1-1]=new Date();
  var resVals=sheet.getRange(row, 1,1,sheet.getLastColumn()).getValues()[0];
  var orderId=resVals[4-1];
  
  arr[2-1]=sheet.getRange(row, 4).getFormula();
  
   var loss=sheet.getRange(row,12).getValue();
  var supId=sheet.getRange(row,13).getValue();
  var title=sheet.getRange(row,2).getValue();
  var orderDate=sheet.getRange(row,1).getValue();
  
  arr[3-1]=addHypToSupId(supId);
  
  
  arr[4-1]=""; // column I of res
  arr[5-1]="PENDING"; // issue pending
  arr[6-1]=""; //reason empty
  arr[7-1]=""; //description empty
  arr[8-1]=sheet.getRange(row,2).getValue(); //product name
  arr[9-1]="";   //trackingNo;
  arr[10-1]=0-sheet.getRange(row,12).getValue();;
  arr[11-1]="";//label empty
  arr[12-1]=""; //restock
  arr[13-1]='=HYPERLINK("https://sellercentral.amazon.com/gp/returns","AUTH")';
  arr[14-1]='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+orderId+'","MSG")'
  arr[15-1]="";// column O expiry empty
  arr[16-1]=""; //column F of res
  arr[17-1]='=HYPERLINK("https://sellercentral.amazon.com/gp/returns","CL")';
 arr[32-1]="";
  arr[33-1]='=HYPERLINK("https://sellercentral.amazon.ca/gp/orders-v2/refund?orderID='+orderId+'",R[0]C[-2])';
  
  
  var trackingInfo=trackingCalc(supId);
  var trackingNo=trackingInfo[0];  
  
  arr[18-1]=trackingNo;
  
  
  var ssRet=SpreadsheetApp.openById(returnSheetId);
  var sheetR=ssRet.getSheetByName("Pending");
  var lrR=sheetR.getLastRow()+1;
  
  
  sheetR.getRange(lrR, 1,1,arr.length).setValues([arr]);


  var profit=sheet.getRange(row, 12).getValue();
  var loss=0-profit;
  
  var supId=sheet.getRange(row, 13).getValue();
  var prodTitle=sheet.getRange(row, 2).getValue();
  
  var values1=[loss,supId,prodTitle]
  var values2=sheet.getRange(row, 1, 1,27).getValues();
  
  var frm=sheet.getRange(row, 7).getFormula();
  values2[0][7-1]=frm;
  
  var rtrns=values1.concat(values2[0]);
  
  
  var vals=[rtrns[2-1+3], rtrns[1-1+3],   rtrns[14-1+3]+rtrns[15-1+3],  rtrns[5-1+3],  rtrns[6-1+3],   rtrns[7-1+3], rtrns[19-1+3],  rtrns[25-1+3],  rtrns[24-1+3], rtrns[16-1+3],  rtrns[11-1+3],  rtrns[26-1+3], rtrns[9-1+3]] //last te, is sale price and to be set at column AE  
  sheetR.getRange(lrR, 19,1,vals.length).setValues([vals]);
  
  var authdata = [ColAH, Authdata, ColAJ, ColAK];
  sheetR.getRange(lrR, 34,1,authdata.length).setValues([authdata]);
  
  
  sheetR.getRange(lrR,14).setFontLine(['line-through']);
  
  
  sheetR.getRange(lrR, 4).setValue("MSG");
  
  if(supId.toString().length >= 14){

    Browser.msgBox("Wrong Returns Tab - Check if AE");
return 0; 
  }
  

  
  var ms = new Date(orderDate).getTime() + (40*86400000);
  var twDay = Utilities.formatDate(new Date(ms),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),"MM/dd/yyyy");
  
  var lastHoliday=new Date("1/1/2018");
  var todaysdate=new Date();
  
  var originalOrderDate=new Date(orderDate);
  if(originalOrderDate.getTime()<lastHoliday.getTime())// if the item is ordered before holiday date then last date is 31st Jan
  {
    twDay="31/01/2018";
  } 
  
  
  
  
  var order20=new Date().getTime() + (20*86400000);
  
  
  if(ms<order20)// if the item is ordered before holiday date then last date is 31st Jan
  {
    
    twDay = Utilities.formatDate(new Date(order20),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),"MM/dd/yyyy");
    sheetR.getRange(lrR, 15).setFontColor("Orange");
  } 
  
  
  
  
  
  sheetR.getRange(lrR, 15).setValue(twDay);
  
  sheetR.getRange(lrR, 2).setFontColor("#1155cc");
  sheetR.getRange(lrR, 3).setFontColor("#1155cc");
  sheetR.getRange(lrR, 13).setFontColor("#1155cc");
  sheetR.getRange(lrR, 14).setFontColor("#1155cc");
  sheetR.getRange(lrR, 17).setFontColor("#1155cc");
  sheetR.getRange(lrR, 18).setFontColor("#1155cc");
  sheetR.getRange(lrR, 25).setFontColor("#1155cc");
  sheetR.getRange(lrR, 33).setFontColor("#1155cc");
  
  
  Browser.msgBox("Added to Returns!");  
  
}



function transfertoReturnsPartial ()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  var currentDate = new Date();
  
  var ColAH ='=iferror(SUBSTITUTE("Hello "&iferror(index(Names!R2C2:C2,match(R[0]C[-32],Names!R2C1:C1,0)),"XXCUSTOMERXX")&","&index(AutoDrafts!R1C3:C3,match("PR",AutoDrafts!R1C1:C1,0))&text(R[0]C[-3]*0.5,"0.00")-0.01&index(AutoDrafts!R1C4:C4,match("PR",AutoDrafts!R1C1:C1,0))&if(isnumber(R[0]C[-23]),text(R[0]C[-23],"#.##"),"")&index(AutoDrafts!R1C5:C5,match("PR",AutoDrafts!R1C1:C1,0)),CHAR(10),CHAR(13)))';
  
  // var ColAH ='=iferror(SUBSTITUTE("Hello "&iferror(index(Names!R2C2:C2,match(R[0]C[-32],Names!R2C1:C1,0)),"XXCUSTOMERXX")&","&index(AutoDrafts!R1C3:C3,match(if(ISERR(SEARCH("/L",R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])-2))&if(OR(NOT(ISERR(SEARCH("Replacement",R[0]C[-30]))),NOT(ISERR(SEARCH("RP",R[0]C[-30])))),"RP","")&if(OR(NOT(ISERR(SEARCH("Exchange",R[0]C[-30]))),NOT(ISERR(SEARCH("EX",R[0]C[-30])))),"EX",""),AutoDrafts!R1C1:C1,0))&R[0]C[-31]&index(AutoDrafts!R1C4:C4,match(if(ISERR(SEARCH("/L",R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])-2))&if(OR(NOT(ISERR(SEARCH("Replacement",R[0]C[-30]))),NOT(ISERR(SEARCH("RP",R[0]C[-30])))),"RP","")&if(OR(NOT(ISERR(SEARCH("Exchange",R[0]C[-30]))),NOT(ISERR(SEARCH("EX",R[0]C[-30])))),"EX",""),AutoDrafts!R1C1:C1,0))&if(isnumber(R[0]C[-23]),text(R[0]C[-23],"#.##"),"")&index(AutoDrafts!R1C5:C5,match(if(ISERR(SEARCH("/L",R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])-2))&if(OR(NOT(ISERR(SEARCH("Replacement",R[0]C[-30]))),NOT(ISERR(SEARCH("RP",R[0]C[-30])))),"RP","")&if(OR(NOT(ISERR(SEARCH("Exchange",R[0]C[-30]))),NOT(ISERR(SEARCH("EX",R[0]C[-30])))),"EX",""),AutoDrafts!R1C1:C1,0)),CHAR(10),CHAR(13)))';
  var ColAJ ='=if(exact(R[0]C[-33],R[0]C[1]),"","N/A")';
  var ColAK ='=iferror(index(SHIPBY!R1C2:C2,match(R[0]C[-35],SHIPBY!R1C1:C1,0)),iferror(index(SHIPBY!R1C5:C5,match(R[0]C[-35],SHIPBY!R1C4:C4,0)),iferror(index(SHIPBY!R1C8:C8,match(R[0]C[-35],SHIPBY!R1C7:C7,0)),iferror(index(SHIPBY!R1C11:C11,match(R[0]C[-35],SHIPBY!R1C10:C10,0))))))';
  var Authdata='=R[0]C[-32]&" "&IF(len(R[0]C[-26])=12,"FedEx",IF(len(R[0]C[-26])=15,"FedEx",IF(len(R[0]C[-26])=18,"UPS",IF(len(R[0]C[-26])=22,"FedEx SmartPost",IF(len(R[0]C[-26])=20,"FedEx","")))))&" "&R[0]C[-26]&" "&TEXT(R[0]C[-24],"#.##")&if(right(R[0]C[-24],2)="00","00","")';
  
  
  
  var arr=[];
  
  arr[1-1]=new Date();
  var resVals=sheet.getRange(row, 1,1,sheet.getLastColumn()).getValues()[0];
  var orderId=resVals[4-1];
  
  arr[2-1]=sheet.getRange(row, 4).getFormula();
  
  
  var loss=sheet.getRange(row,12).getValue();
  var supId=sheet.getRange(row,13).getValue();
  var title=sheet.getRange(row,2).getValue();
  var orderDate=sheet.getRange(row,1).getValue();
  
  arr[3-1]=addHypToSupId(supId);
  
  
  arr[4-1]=""; // column I of res
  arr[5-1]="PENDING"; // issue pending
  arr[6-1]=""; //reason empty
  arr[7-1]=""; //description empty
  arr[8-1]=sheet.getRange(row, 2).getValue(); //product name
  arr[9-1]="";   //trackingNo;
  arr[10-1]=0-sheet.getRange(row,12).getValue();
  arr[11-1]="";//label empty
  arr[12-1]=""; //restock
  arr[13-1]='=HYPERLINK("https://sellercentral.amazon.com/gp/returns","AUTH")';
  arr[14-1]='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+orderId+'","MSG")'
  arr[15-1]="";// column O expiry empty
  arr[16-1]=""; //column F of res
  arr[17-1]='=HYPERLINK("https://sellercentral.amazon.com/gp/returns","CL")';
   arr[32-1]="";
  arr[33-1]='=HYPERLINK("https://sellercentral.amazon.ca/gp/orders-v2/refund?orderID='+orderId+'",R[0]C[-2])';
  
  /*
  var trackingInfo=trackingCalc(supId);
  var trackingNo=trackingInfo[0];  
  

*/
  arr[18-1]="";

  
  
  var ssRet=SpreadsheetApp.openById(returnSheetId);
  var sheetR=ssRet.getSheetByName("Partial");
  var lrR=sheetR.getLastRow()+1;
  
  
  var profit=sheet.getRange(row, 12).getValue();
  var loss=0-profit;
  
  var supId=sheet.getRange(row, 13).getValue();
  var prodTitle=sheet.getRange(row, 2).getValue();
  
  var values1=[loss,supId,prodTitle]
  var values2=sheet.getRange(row, 1, 1,27).getValues();
  
  var frm=sheet.getRange(row, 7).getFormula();
  values2[0][7-1]=frm;
  
  var rtrns=values1.concat(values2[0]);
  
  
  
  sheetR.getRange(lrR, 1,1,arr.length).setValues([arr]);
  
  var vals=[rtrns[2-1+3], rtrns[1-1+3],   rtrns[14-1+3]+rtrns[15-1+3],  rtrns[5-1+3],  rtrns[6-1+3],   rtrns[7-1+3], rtrns[19-1+3],  rtrns[25-1+3],  rtrns[24-1+3], rtrns[16-1+3],  rtrns[11-1+3],  rtrns[26-1+3], rtrns[9-1+3]] //last te, is sale price and to be set at column AE  
  sheetR.getRange(lrR, 19,1,vals.length).setValues([vals]);
  
  var authdata = [ColAH, Authdata, ColAJ, ColAK];
  sheetR.getRange(lrR, 34,1,authdata.length).setValues([authdata]);
  
  
  sheetR.getRange(lrR,14).setFontLine(['line-through']);
  
  
  sheetR.getRange(lrR, 4).setValue("MSG");
  
  if(supId.toString().length == 14){
    sheetR.getRange(lrR, 4).setValue("MSG");
  }
  
  if(title.indexOf("Duvet") >= 0 && supId.toString().length == 14  ){
    sheetR.getRange(lrR, 4).setValue("MSG");
  }
  
  
  var ms = new Date(orderDate).getTime() + (40*86400000);
  var twDay = Utilities.formatDate(new Date(ms),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),"MM/dd/yyyy");
  
  var lastHoliday=new Date("1/1/2018");
  var todaysdate=new Date();
  
  var originalOrderDate=new Date(orderDate);
  if(originalOrderDate.getTime()<lastHoliday.getTime())// if the item is ordered before holiday date then last date is 31st Jan
  {
    twDay="31/01/2018";
  } 
  
  
  
  
  var order20=new Date().getTime() + (20*86400000);
  
  
  if(ms<order20)// if the item is ordered before holiday date then last date is 31st Jan
  {
    
    twDay = Utilities.formatDate(new Date(order20),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),"MM/dd/yyyy");
    sheetR.getRange(lrR, 15).setFontColor("Orange");
  } 
  
  
  
  
  
  sheetR.getRange(lrR, 15).setValue(twDay);
  
  sheetR.getRange(lrR, 2).setFontColor("#1155cc");
  sheetR.getRange(lrR, 3).setFontColor("#1155cc");
  sheetR.getRange(lrR, 13).setFontColor("#1155cc");
  sheetR.getRange(lrR, 14).setFontColor("#1155cc");
  sheetR.getRange(lrR, 17).setFontColor("#1155cc");
  sheetR.getRange(lrR, 18).setFontColor("#1155cc");
  sheetR.getRange(lrR, 25).setFontColor("#1155cc");
  sheetR.getRange(lrR, 33).setFontColor("#1155cc");
  
  
  Browser.msgBox("Added to Returns - Partial Tab!"); 
  
}



function transfertoRF ()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  var currentDate = new Date();
  
var ColAH = "";
 // var ColAH ='=iferror(SUBSTITUTE("Hello "&iferror(index(Names!R2C2:C2,match(R[0]C[-32],Names!R2C1:C1,0)),"XXCUSTOMERXX")&","&index(AutoDrafts!R1C3:C3,match(if(ISERR(SEARCH("/L",R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])-2))&if(OR(NOT(ISERR(SEARCH("Replacement",R[0]C[-30]))),NOT(ISERR(SEARCH("RP",R[0]C[-30])))),"RP","")&if(OR(NOT(ISERR(SEARCH("Exchange",R[0]C[-30]))),NOT(ISERR(SEARCH("EX",R[0]C[-30])))),"EX",""),AutoDrafts!R1C1:C1,0))&R[0]C[-31]&index(AutoDrafts!R1C4:C4,match(if(ISERR(SEARCH("/L",R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])-2))&if(OR(NOT(ISERR(SEARCH("Replacement",R[0]C[-30]))),NOT(ISERR(SEARCH("RP",R[0]C[-30])))),"RP","")&if(OR(NOT(ISERR(SEARCH("Exchange",R[0]C[-30]))),NOT(ISERR(SEARCH("EX",R[0]C[-30])))),"EX",""),AutoDrafts!R1C1:C1,0))&if(isnumber(R[0]C[-23]),text(R[0]C[-23],"#.##"),"")&index(AutoDrafts!R1C5:C5,match(if(ISERR(SEARCH("/L",R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])),left(R[0]C[-28],len(R[0]C[-28])-2))&if(OR(NOT(ISERR(SEARCH("Replacement",R[0]C[-30]))),NOT(ISERR(SEARCH("RP",R[0]C[-30])))),"RP","")&if(OR(NOT(ISERR(SEARCH("Exchange",R[0]C[-30]))),NOT(ISERR(SEARCH("EX",R[0]C[-30])))),"EX",""),AutoDrafts!R1C1:C1,0)),CHAR(10),CHAR(13)))';
  var ColAJ ='=if(exact(R[0]C[-33],R[0]C[1]),"","N/A")';
  var ColAK ='=iferror(index(SHIPBY!R1C2:C2,match(R[0]C[-35],SHIPBY!R1C1:C1,0)),iferror(index(SHIPBY!R1C5:C5,match(R[0]C[-35],SHIPBY!R1C4:C4,0)),iferror(index(SHIPBY!R1C8:C8,match(R[0]C[-35],SHIPBY!R1C7:C7,0)),iferror(index(SHIPBY!R1C11:C11,match(R[0]C[-35],SHIPBY!R1C10:C10,0))))))';
  var Authdata='=R[0]C[-32]&" "&IF(len(R[0]C[-26])=12,"FedEx",IF(len(R[0]C[-26])=15,"FedEx",IF(len(R[0]C[-26])=18,"UPS",IF(len(R[0]C[-26])=22,"FedEx SmartPost",IF(len(R[0]C[-26])=20,"FedEx","")))))&" "&R[0]C[-26]&" "&TEXT(R[0]C[-24],"#.##")&if(right(R[0]C[-24],2)="00","00","")';
  
  
  
  var arr=[];
  
  arr[1-1]=new Date();
  var resVals=sheet.getRange(row, 1,1,sheet.getLastColumn()).getValues()[0];
  var orderId=resVals[4-1];
  
  arr[2-1]=sheet.getRange(row, 4).getFormula();
  
  
  var loss=sheet.getRange(row,12).getValue();
  var supId=sheet.getRange(row,13).getValue();
  var title=sheet.getRange(row,2).getValue();
  var orderDate=sheet.getRange(row,1).getValue();
  
  arr[3-1]=addHypToSupId(supId);
  
  arr[4-1]=""; // column I of res
  arr[5-1]="RF"; // issue pending
  arr[6-1]="NL"; //reason empty
  arr[7-1]="No longer needed"; //description empty
  arr[8-1]=sheet.getRange(row, 2).getValue(); //product name
  arr[9-1]="Pending Tracking";   //trackingNo;
  arr[10-1]=0-sheet.getRange(row,12).getValue();
  arr[11-1]="N/A";//label empty
  arr[12-1]="0.00"; //restock
  arr[13-1]='=HYPERLINK("https://sellercentral.amazon.com/gp/returns","AUTH")';
  arr[14-1]='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+orderId+'","MSG")'
  arr[15-1]="";// column O expiry empty
  arr[16-1]=""; //column F of res
  arr[17-1]='=HYPERLINK("https://sellercentral.amazon.com/gp/returns","CL")';
  arr[32-1]="";
  arr[33-1]='=HYPERLINK("https://sellercentral.amazon.ca/gp/orders-v2/refund?orderID='+orderId+'",R[0]C[-2])';
  
  
  var trackingInfo=trackingCalc(supId);
  var trackingNo=trackingInfo[0];  
  
  arr[18-1]=trackingNo;
  
  
  var ssRet=SpreadsheetApp.openById(returnSheetId);
  var sheetR=ssRet.getSheetByName("RF");
  var lrR=sheetR.getLastRow()+1;
  
  
  sheetR.getRange(lrR, 1,1,arr.length).setValues([arr]);
  
  
  var profit=sheet.getRange(row, 12).getValue();
  var loss=0-profit;
  
  var supId=sheet.getRange(row, 13).getValue();
  var prodTitle=sheet.getRange(row, 2).getValue();
  
  var values1=[loss,supId,prodTitle]
  var values2=sheet.getRange(row, 1, 1,27).getValues();
  
  var frm=sheet.getRange(row, 7).getFormula();
  values2[0][7-1]=frm;
  
  var rtrns=values1.concat(values2[0]);  
  
  
  var vals=[rtrns[2-1+3], rtrns[1-1+3],   rtrns[14-1+3]+rtrns[15-1+3],  rtrns[5-1+3],  rtrns[6-1+3],   rtrns[7-1+3], rtrns[19-1+3],  rtrns[25-1+3],  rtrns[24-1+3], rtrns[16-1+3],  rtrns[11-1+3],  rtrns[26-1+3], rtrns[9-1+3]] //last te, is sale price and to be set at column AE  
  sheetR.getRange(lrR, 19,1,vals.length).setValues([vals]);
  
  
  
  
  var authdata = [ColAH, Authdata, ColAJ, ColAK];
  sheetR.getRange(lrR, 34,1,authdata.length).setValues([authdata]);
  
  
  
  var ms = new Date(orderDate).getTime() + (40*86400000);
  var twDay = Utilities.formatDate(new Date(ms),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),"MM/dd/yyyy");
  
  var lastHoliday=new Date("1/1/2018");
  var todaysdate=new Date();
  
  var originalOrderDate=new Date(orderDate);
  if(originalOrderDate.getTime()<lastHoliday.getTime())// if the item is ordered before holiday date then last date is 31st Jan
  {
    twDay="31/01/2018";
  } 
  
  
  
  
  var order20=new Date().getTime() + (20*86400000);
  
  
  if(ms<order20)// if the item is ordered before holiday date then last date is 31st Jan
  {
    
    twDay = Utilities.formatDate(new Date(order20),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),"MM/dd/yyyy");
    sheetR.getRange(lrR, 15).setFontColor("Orange");
  } 
  
  
  
  
  
  sheetR.getRange(lrR, 15).setValue(twDay);
  
  sheetR.getRange(lrR, 2).setFontColor("#1155cc");
  sheetR.getRange(lrR, 3).setFontColor("#1155cc");
  sheetR.getRange(lrR, 13).setFontColor("#1155cc");
  sheetR.getRange(lrR, 14).setFontColor("#1155cc");
  sheetR.getRange(lrR, 17).setFontColor("#1155cc");
  sheetR.getRange(lrR, 18).setFontColor("#1155cc");
  sheetR.getRange(lrR, 25).setFontColor("#1155cc");
  sheetR.getRange(lrR, 33).setFontColor("#1155cc");
  
  Browser.msgBox("Added to RF!"); 
  
}


function transfertoRP ()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  var currentDate = new Date();
  
  var ssRet=SpreadsheetApp.openById(returnSheetId);
  var sheetR=ssRet.getSheetByName("RP");
  var lrR=sheetR.getLastRow()+1;
  var first = ssRet.getSheetByName("RP");
  first.setTabColor("#ee799f");
  
  var date = new Date();
  
  var orderIdv = sheet.getRange(row, 4).getValue();
  var orderId = sheet.getRange(row, 4).getFormula();
  var status = "PENDING";
  
  
  
  
  var loss=sheet.getRange(row, 12).getValue();
  var supId=sheet.getRange(row, 13).getValue();
  var MSG = '=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+orderIdv+'","MSG")';
  var Draft='=iferror(SUBSTITUTE("Hello "&iferror(index(Names!R1C2:C2,match(R[0]C[-15],Names!R1C1:C1,0)),"XXCUSTOMERXX")&","&iferror(index(AutoDrafts!R1C3:C3,match(R[0]C[-18],AutoDrafts!R1C1:C1,0))&R[0]C[-3]&if(R[0]C[-18]="COMPLETED"," - ","")&IF(len(R[0]C[-3])=12,"FedEx",IF(len(R[0]C[-3])=15,"FedEx",IF(len(R[0]C[-3])=18,"UPS",IF(len(R[0]C[-3])=22,"FedEx SmartPost",IF(len(R[0]C[-3])=20,"FedEx","")))))&iferror(index(AutoDrafts!R1C4:C4,match(R[0]C[-18],AutoDrafts!R1C1:C1,0)))," "),CHAR(10),CHAR(13))," ")'; 
  
  var Profit ="=M"+lrR+"-N"+lrR+"-O"+lrR;
  var Diff = "=P"+lrR+"-"+"R"+lrR;
  
  var qty = sheet.getRange(row, 3).getValue();
  var itemn = sheet.getRange(row, 5).getValue();
  var variation = sheet.getRange(row, 6).getValue();
  var soldprice = sheet.getRange(row, 9).getValue();
  var amazonfees = sheet.getRange(row, 10).getValue();
  var orgprofit = sheet.getRange(row, 12).getValue();
  
  var supplier = sheet.getRange(row, 19).getValue();
  
  if(supplier.indexOf("overstock")>=0)
  {
    
    var OScoupon = retrieveCouponUrl(orderIdv);
    var cashbackbutton ='=hyperlink("'+supplier+'","OS")';
  }
  
  else{
    
    var OScoupon = "";
    var cashbackbutton =sheet.getRange(row, 7).getFormula();
  }
  
  
  sheetR.getRange(lrR, 1).setValue(date);
  
  sheetR.getRange(lrR, 5).setValue(status);
  sheetR.getRange(lrR, 7).setValue(qty);
  sheetR.getRange(lrR, 8).setValue(orderId);
  sheetR.getRange(lrR, 9).setValue(itemn);
  sheetR.getRange(lrR, 10).setValue(variation);
  sheetR.getRange(lrR, 11).setValue(cashbackbutton);
  sheetR.getRange(lrR, 12).setValue(OScoupon);
  sheetR.getRange(lrR, 13).setValue(soldprice);
  sheetR.getRange(lrR, 14).setValue(amazonfees);
  sheetR.getRange(lrR, 16).setValue(Profit);
  sheetR.getRange(lrR, 18).setValue(orgprofit);
  sheetR.getRange(lrR, 19).setValue(Diff);
  sheetR.getRange(lrR, 21).setValue(MSG);
  sheetR.getRange(lrR, 23).setValue(Draft);
  
  
  Browser.msgBox("Added to RP!");  
  
}



function transfertoDEDcog ()
{
  var ss =  SpreadsheetApp.getActiveSpreadsheet();
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  var currentDate = new Date();
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  
  var Ded = SpreadsheetApp.openById(dedId)
  var sheetD=Ded.getSheetByName(sheetName);
  
  var lr= last_row(sheetD, 1);
  
  var orderIDv = sheet.getRange(row, 4).getFormula();
  var profitloss=sheet.getRange(row,12).getValue();
  var cogloss=sheet.getRange(row,11).getValue();
  
  
  sheetD.getRange(lr+1,1).setValue(currentDate);
  sheetD.getRange(lr+1,2).setValue(orderIDv);
  sheetD.getRange(lr+1,3).setValue("Order Issue");
  sheetD.getRange(lr+1,4).setValue("-"+profitloss);
  sheetD.getRange(lr+1,5).setValue("-"+cogloss);
  sheetD.getRange(lr+1,7).setValue("FR/COG Included");
  
  
  
  
  Browser.msgBox("Transferred to ADV!");    
}  


function transfertoDEDloss ()
{
  var ss =  SpreadsheetApp.getActiveSpreadsheet();
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  var currentDate = new Date();
  var sheetName=Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "MMMM-yyyy")
  
  var Ded = SpreadsheetApp.openById(dedId)
  var sheetD=Ded.getSheetByName(sheetName);
  
  var lr= last_row(sheetD, 1);
  
  var orderIDv = sheet.getRange(row, 4).getFormula();
  var profitloss=sheet.getRange(row,12).getValue();
  var cogloss=sheet.getRange(row,11).getValue();
  
  
  sheetD.getRange(lr+1,1).setValue(currentDate);
  sheetD.getRange(lr+1,2).setValue(orderIDv);
  sheetD.getRange(lr+1,3).setValue("Order Issue");
  sheetD.getRange(lr+1,4).setValue("-"+profitloss);
  sheetD.getRange(lr+1,5).setValue("0.00");
  
  
  
  Browser.msgBox("Transferred to ADV!");    
}  




function zerooutalladv() {
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheetName=sheet.getName();
  
  sheet.getRange(row, 3).setValue("0").setFontColor("orange");
  sheet.getRange(row, 9).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 10).setValue("0.00").setFontColor("orange");
  
  sheet.getRange(row, 11).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 12).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 26).setValue("ISSUE").setFontColor("orange");
  sheet.getRange(row, 27).clearContent();
  
  var currentDate = new Date();
  var Ded = SpreadsheetApp.openById(resSsId);
  var sheetD=Ded.getSheetByName("ADV");
  
  var lr= last_row(sheetD, 4);
  
  var orderIDv = sheet.getRange(row, 4).getFormula();
  var supId=sheet.getRange(row,13).getValue();
  
  var orderId = sheet.getRange(row, 4).getValue();
  
  
  var MSG='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+orderId+'","MSG")';
  
  if(supId.length < 14)
  {
    var trackingInfo=trackingCalc(supId);
    var trackingNo=trackingInfo[0];
  }
  else
  {
    var trackingNo = "";
  }
  
  sheetD.getRange(lr+1,1).setValue(currentDate);
  sheetD.getRange(lr+1,2).setValue("");
  sheetD.getRange(lr+1,3).setValue("REFUND");
  sheetD.getRange(lr+1,4).setFormula(orderIDv); 
  sheetD.getRange(lr+1,5).setValue(addHypToSupId(supId))
  
  sheetD.getRange(lr+1,7).setFormula(MSG);
  sheetD.getRange(lr+1,8).setValue("LS");
  sheetD.getRange(lr+1,12).setValue("O");
  sheetD.getRange(lr+1,13).setValue("Yes");
  sheetD.getRange(lr+1, 14).setFormula(trackingNo)
  
   


  Browser.msgBox("Transferred to ADV!");

  var h1 = '<html>';
  



  var amzrefundpage = "https://sellercentral.amazon.com/gp/orders-v2/refund?orderID="+orderId;
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+amzrefundpage+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');

  var disputeae = "https://trade.aliexpress.com/issue/fastissue/createIssueStep1.htm?orderId="+supId;
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+disputeae+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );





   
}


function zerooutalladvcog() {
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheetName=sheet.getName();
  
  sheet.getRange(row, 3).setValue("0").setFontColor("orange");
  sheet.getRange(row, 9).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 10).setValue("0.00").setFontColor("orange");
  
  
  sheet.getRange(row, 12).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 26).setValue("ISSUE").setFontColor("orange");
  sheet.getRange(row, 27).clearContent();
  
  var currentDate = new Date();
  var Ded = SpreadsheetApp.openById(resSsId);
  var sheetD=Ded.getSheetByName("ADV");
  
  var lr= last_row(sheetD, 4);
  
  var orderIDv = sheet.getRange(row, 4).getFormula();
  var supId=sheet.getRange(row,13).getValue();
  
  var orderId = sheet.getRange(row, 4).getValue();
  
  
  var MSG='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+orderId+'","MSG")';
  
  if(supId.length < 14)
  {
    var trackingInfo=trackingCalc(supId);
    var trackingNo=trackingInfo[0];
  }
  else
  {
    var trackingNo = "";
  }
  
  sheetD.getRange(lr+1,1).setValue(currentDate);
  sheetD.getRange(lr+1,2).setValue("");
  sheetD.getRange(lr+1,3).setValue("REFUND");
  sheetD.getRange(lr+1,4).setFormula(orderIDv); 
  sheetD.getRange(lr+1,5).setValue(addHypToSupId(supId))
  
  sheetD.getRange(lr+1,7).setFormula(MSG);
  sheetD.getRange(lr+1,8).setValue("LS");
  sheetD.getRange(lr+1,12).setValue("O");
  sheetD.getRange(lr+1,13).setValue("No");
  sheetD.getRange(lr+1, 14).setFormula(trackingNo)
  
  
  Browser.msgBox("Transferred to ADV!");

  var h1 = '<html>';
  



  var amzrefundpage = "https://sellercentral.amazon.com/gp/orders-v2/refund?orderID="+orderId;
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+amzrefundpage+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');

  var disputeae = "https://trade.aliexpress.com/issue/fastissue/createIssueStep1.htm?orderId="+supId;
  
  var h1 = h1.concat('<script>'
                     +'function sleep(miliseconds) {var currentTime = new Date().getTime();while (currentTime + miliseconds >= new Date().getTime()) {}}sleep(0000);var a = document.createElement("a"); a.href="'+disputeae+'"; a.target="_blank";'
                     +'if(document.createEvent){var event=document.createEvent("MouseEvents");event.initEvent("click",true,true);a.dispatchEvent(event)}else{a.click()}'
                     +'</script>');
  
  
  
  
  
  var h1 = h1.concat('<script>google.script.host.close();</script><body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="#" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.</body>'
                     +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script></html>');
  //Browser.msgBox(h1); 
  var html1 = HtmlService.createHtmlOutput(h1)
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html1, "Opening ..." );
}




function transfertoPENDING ()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  var currentDate = new Date();
  
  var Ded = SpreadsheetApp.openById(resSsId);
  var sheetD=Ded.getSheetByName("PENDING");
  
  var lr= last_row(sheetD, 4);
  
  var orderIDv = sheet.getRange(row, 4).getFormula();
  var supId=sheet.getRange(row,13).getValue();
  
  
  
  var orderId=sheet.getRange(row,4).getValue();
  orderId=orderId.toString();
  
  var MSG='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+orderId+'","MSG")';
  
  
  
  
  
  sheetD.getRange(lr+1,1).setValue(currentDate);
  sheetD.getRange(lr+1,2).setValue("");
  sheetD.getRange(lr+1,3).setValue("PENDING");
  sheetD.getRange(lr+1,4).setFormula(orderIDv); 
  sheetD.getRange(lr+1,5).setValue(addHypToSupId(supId))
  sheetD.getRange(lr+1,7).setFormula(MSG);
  
  
  
  var profit=sheet.getRange(row, 12).getValue();
  var loss=0-profit;
  
  var supId=sheet.getRange(row, 13).getValue();
  var prodTitle=sheet.getRange(row, 2).getValue();
  
  var values1=[loss,supId,prodTitle]
  var values2=sheet.getRange(row, 1, 1,27).getValues();
  
  var frm=sheet.getRange(row, 7).getFormula();
  values2[0][7-1]=frm;
  
  var rtrns=values1.concat(values2[0]);
  
  
  
  var ASINhyp='=HYPERLINK("https://www.amazon.com/gp/product/'+rtrns[24-1+3]+'","'+rtrns[24-1+3]+'")';
  var SKUhyp='=HYPERLINK("https://catalog.amazon.com/abis/edit/RelistProduct.amzn?marketplaceID=ATVPDKIKX0DER&ref=xx_myirelis_cont_myifba&sku='+rtrns[25-1+3]+'&asin='+rtrns[24-1+3]+'&productType=HOME","'+rtrns[25-1+3]+'")';
  
  
  var vals=[rtrns[2-1+3], rtrns[1-1+3],   rtrns[12-1+3],  rtrns[5-1+3],  rtrns[6-1+3],   rtrns[7-1+3], rtrns[19-1+3],  SKUhyp,  ASINhyp, rtrns[16-1+3],  rtrns[11-1+3],  rtrns[26-1+3]]  
  sheetD.getRange(lr+1, 13,1,vals.length).setValues([vals]);
  
  
  
  Browser.msgBox("Transferred to PENDING!");    
} 







function transfertoOI ()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  var currentDate = new Date();
  
  var Ded = SpreadsheetApp.openById(resSsId);
  var sheetD=Ded.getSheetByName("OI");
  
  var lr= last_row(sheetD, 4);
  
  var orderIDv = sheet.getRange(row, 4).getFormula();
  var supId=sheet.getRange(row,13).getValue();
  
  
  
  var orderId=sheet.getRange(row,4).getValue();
  orderId=orderId.toString();
  
  var MSG='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+orderId+'","MSG")';
  
  
  
  
  
  sheetD.getRange(lr+1,1).setValue(currentDate);
  sheetD.getRange(lr+1,2).setValue("");
  sheetD.getRange(lr+1,3).setValue("PENDING");
  sheetD.getRange(lr+1,4).setFormula(orderIDv); 
  sheetD.getRange(lr+1,5).setValue(addHypToSupId(supId))
  sheetD.getRange(lr+1,7).setFormula(MSG);
  
  var profit=sheet.getRange(row, 12).getValue();
  var loss=0-profit;
  
  var supId=sheet.getRange(row, 13).getValue();
  var prodTitle=sheet.getRange(row, 2).getValue();
  
  var values1=[loss,supId,prodTitle]
  var values2=sheet.getRange(row, 1, 1,27).getValues();
  
  var frm=sheet.getRange(row, 7).getFormula();
  values2[0][7-1]=frm;
  
  var rtrns=values1.concat(values2[0]);
  
  
  
  var ASINhyp='=HYPERLINK("https://www.amazon.com/gp/product/'+rtrns[24-1+3]+'","'+rtrns[24-1+3]+'")';
  var SKUhyp='=HYPERLINK("https://catalog.amazon.com/abis/edit/RelistProduct.amzn?marketplaceID=ATVPDKIKX0DER&ref=xx_myirelis_cont_myifba&sku='+rtrns[25-1+3]+'&asin='+rtrns[24-1+3]+'&productType=HOME","'+rtrns[25-1+3]+'")';
  
  
  var vals=[rtrns[2-1+3], rtrns[1-1+3],   rtrns[12-1+3],  rtrns[5-1+3],  rtrns[6-1+3],   rtrns[7-1+3], rtrns[19-1+3],  SKUhyp,  ASINhyp, rtrns[16-1+3],  rtrns[11-1+3],  rtrns[26-1+3]]  
  sheetD.getRange(lr+1, 13,1,vals.length).setValues([vals]);
  
  
  
  Browser.msgBox("Transferred to PENDING!");    
}  


function transfertoESC ()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  var currentDate = new Date();
  
  var Ded = SpreadsheetApp.openById(resSsId);
  var sheetD=Ded.getSheetByName("ESC");
  
  var lr= last_row(sheetD, 4);
  
  var orderIDv = sheet.getRange(row, 4).getFormula();
  var supId=sheet.getRange(row,13).getValue();
  
  
  
  var orderId=sheet.getRange(row,4).getValue();
  orderId=orderId.toString();
  
  var MSG='=HYPERLINK("https://sellercentral.amazon.com/gp/orders-v2/contact?orderID='+orderId+'","MSG")';
  
  
  
  
  
  sheetD.getRange(lr+1,1).setValue(currentDate);
  sheetD.getRange(lr+1,2).setValue("");
  sheetD.getRange(lr+1,3).setValue("PENDING");
  sheetD.getRange(lr+1,4).setFormula(orderIDv); 
  sheetD.getRange(lr+1,5).setValue(addHypToSupId(supId))
  sheetD.getRange(lr+1,7).setFormula(MSG);
  
  
  var profit=sheet.getRange(row, 12).getValue();
  var loss=0-profit;
  
  var supId=sheet.getRange(row, 13).getValue();
  var prodTitle=sheet.getRange(row, 2).getValue();
  
  var values1=[loss,supId,prodTitle]
  var values2=sheet.getRange(row, 1, 1,27).getValues();
  
  var frm=sheet.getRange(row, 7).getFormula();
  values2[0][7-1]=frm;
  
  var rtrns=values1.concat(values2[0]);
  
  
  
  var ASINhyp='=HYPERLINK("https://www.amazon.com/gp/product/'+rtrns[24-1+3]+'","'+rtrns[24-1+3]+'")';
  var SKUhyp='=HYPERLINK("https://catalog.amazon.com/abis/edit/RelistProduct.amzn?marketplaceID=ATVPDKIKX0DER&ref=xx_myirelis_cont_myifba&sku='+rtrns[25-1+3]+'&asin='+rtrns[24-1+3]+'&productType=HOME","'+rtrns[25-1+3]+'")';
  
  
  var vals=[rtrns[2-1+3], rtrns[1-1+3],   rtrns[12-1+3],  rtrns[5-1+3],  rtrns[6-1+3],   rtrns[7-1+3], rtrns[19-1+3],  SKUhyp,  ASINhyp, rtrns[16-1+3],  rtrns[11-1+3],  rtrns[26-1+3]]  
  sheetD.getRange(lr+1, 13,1,vals.length).setValues([vals]);
  
  
  
  Browser.msgBox("Transferred to ESC!");    
}  


function oshippedorder ()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  
  var costs = sheet.getRange(row,11).getValue();
  
  if( costs == "")
  {
    
    var markupform ='=R[0]C[-3]-R[0]C[-2]-R[0]C[-1]';
    
    
    sheet.getRange(row, 12).setValue(markupform);
    sheet.getRange(row, 12).setBackground("#E9E9E9");
    sheet.getRange(row, 12).setFontColor("#197319");
    sheet.getRange(row, 16).clearContent;
    
  }
  
  var orderId=sheet.getRange(row, 4).getValue();
  
  
  var sheet6=SpreadsheetApp.openById(resSsId);
  var sheetOOS=sheet6.getSheetByName("OOS");
  
  var rowO=lookup(orderId,sheetOOS,4, 6,"row");
  
  
  
  sheetOOS.getRange(rowO, 3).setValue("O-SHIPPED");
  sheetOOS.getRange(rowO, 11).setValue("Shipped Order");
  
  
  Browser.msgBox("Marked Shipped - OOS!");
  
  
  
}



function resetorder ()
{
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  
  
  
  var markupform ='=R[0]C[-3]-R[0]C[-2]-R[0]C[-1]';
  
  sheet.getRange(row, 11).clearContent();
  sheet.getRange(row, 12).setValue(markupform);
  sheet.getRange(row, 12).setBackground("#E9E9E9");
  sheet.getRange(row, 12).setFontColor("#197319");
  sheet.getRange(row, 13).clearContent();
  sheet.getRange(row, 16).clearContent();
  sheet.getRange(row, 17).clearContent();
  
  
  
  
  
}






function zerooutallRF ()
{
  
  
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  
  var rng=ss.getActiveSheet().getActiveRange();
  var sheet=rng.getSheet();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheetName=sheet.getName();
  
  sheet.getRange(row, 3).setValue("0").setFontColor("orange");
  sheet.getRange(row, 9).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 10).setValue("0.00").setFontColor("orange");
  
  sheet.getRange(row, 11).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 12).setValue("0.00").setFontColor("orange");
  sheet.getRange(row, 26).setValue("ISSUE").setFontColor("orange");
  sheet.getRange(row, 27).clearContent();
  
  
  
  
  var rng=SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var row=rng.getRow();
  var col=rng.getColumn();
  var sheet=rng.getSheet();
  
  
  var orderId=sheet.getRange(row, 4).getValue();
  
  
  var sheet6=SpreadsheetApp.openById(returnSheetId);
  var sheetRF=sheet6.getSheetByName("RF");
  
  var rowO=lookup(orderId,sheetRF,2, 6,"row");
  
  
  
  
  if(rowO == null)
  {
    
    Browser.msgBox("Note: Not Found on RF Tab - Order Cancelled Successfully");
    
    return 0;
    
  }
  else if(orderId == sheetRF.getRange(rowO, 2).getValue())
  {
    var orderIDRF = sheetRF.getRange(rowO, 2).getValue();
    sheetRF.deleteRow(rowO);
    
    
    
    Browser.msgBox("Order ID: "+orderIDRF+" Deleted from RF Tab! Order Cancelled Successfully");
  }
  
  
  
}
