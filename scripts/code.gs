function doPost(e)
{
  var date = new Date();
  var seetName = Utilities.formatDate(date,Session.getScriptTimeZone(),"yyyyMMdd");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(seetName);
  if(!sheet){
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    sheet.setName(seetName);
  }
  var params = e.postData.getDataAsString().split(" ");
  sheet.insertRows(1,1);
  sheet.getRange(1, 1).setNumberFormat("yyyy/MM/dd H:mm:ss");
  sheet.getRange(1, 1).setValue(new Date());
  for(i=0;i<params.length;i++){
    sheet.getRange(1, 2+i).setValue(Number(params[i]));
  }
}
function doGet(e) {
  var html = HtmlService.createTemplateFromFile('index');
  return html.evaluate();
}

function getSheetUrl(dateString) {
  //Date
  var date = new Date();
  date = new Date(date.getFullYear(), date.getMonth(), 0, 0, 0, 0);
  var q =  dateString.replace(/[^0-9]/g, "");
  //Range
  var ssid = SpreadsheetApp.getActiveSpreadsheet().getId();
  var gid = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(q).getSheetId();
  return "https://docs.google.com/spreadsheets/d/"+ssid+"/gviz/tq?gid="+gid+"#range=A:E";
}