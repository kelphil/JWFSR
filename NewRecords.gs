function getNewRecords()
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet_form = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var sheet_control = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Controls');
  var sheet_rawdata = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FS Raw Data');
  
  var range_control = sheet_control.getRange("A1:A13");
  var row_control = range_control.getValues();
  var lastRecord_form = row_control[10][0];
  var lastRecord_rawdata = row_control[12][0];
  
  var range_form = sheet_form.getRange("A1:R5000");
  var row = range_form.getValues();
  
  var newRecord = false;
  var lastRecord = lastRecord_form;
  var count = 1;
  
  for (var i=lastRecord_form;i<row.length;i++)
  {
    for (var j=0;j<18;j++)
    {
      var val = row[i][j];
      
      if (val != "")
      {
        sheet_rawdata.getRange(lastRecord_rawdata+count,j+1).setValue(val);
        newRecord = true;
        lastRecord = i+1;
      }
    }
    count++;
  }
  
  sheet_control.getRange(11,1).setValue(lastRecord);
  
//  findLastRecord();
  
  return newRecord;
}