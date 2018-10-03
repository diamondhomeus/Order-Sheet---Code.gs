

function fncOpenMyDialog() {
  //Open a dialog
  var htmlDlg = HtmlService.createHtmlOutputFromFile('colorname.html')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(300)
      .setHeight(250);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlDlg, 'Order Issue Menu:');
};



