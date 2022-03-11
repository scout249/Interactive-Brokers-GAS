var token = 'YOUR TOKEN';
var query = 'YOUR QUERY';

// custom menu function
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Interactive Broker')
      .addItem('Trade Activity','getFlexData')
      .addToUi();
}

function initFlexRequest() {
  var url = 'https://gdcdyn.interactivebrokers.com/Universal/servlet/FlexStatementService.SendRequest?t=' + token + '&q=' + query + '&v=3';
  var xml = UrlFetchApp.fetch(url).getContentText();
  var document = XmlService.parse(xml);
  var Status = document.getRootElement().getChild('Status').getValue();
  var ReferenceCode = document.getRootElement().getChild('ReferenceCode').getValue();
  Logger.log(Status);
  Logger.log(ReferenceCode);
  Utilities.sleep(2000);
  return [Status,ReferenceCode];
}

function getFlexData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();  
  Logger.log(initFlexRequest()[1]);
  var csvurl = 'https://gdcdyn.interactivebrokers.com/Universal/servlet/FlexStatementService.GetStatement?q=' + initFlexRequest()[1] + '&t=' + token + '&v=3';
  var csvContent = UrlFetchApp.fetch(csvurl).getContentText();
  var csvData = Utilities.parseCsv(csvContent);
  sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
}

