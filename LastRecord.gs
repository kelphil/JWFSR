function findLastRecord()
{ 
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet_form = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var sheet_control = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Controls');
  var sheet_rawdata = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FS Raw Data');

  var range_control = sheet_control.getRange("A1:A13");
  var row_control = range_control.getValues();
  var lastRecord_form = row_control[10][0];
  var lastRecord_rawdata = row_control[12][0];
  
  var range_form = sheet_form.getRange("A2:A5000");
  var row = range_form.getValues();
  var lastRecord = 2;
  
  var newRecord = true;
  
  for (var i=0;i<row.length;i++)
  {
    var tstamp = row[i][0];
    
    if (tstamp != "")
    {
      lastRecord = i + 2;
    }
  }
  
  if (lastRecord == lastRecord_form)
  {
    newRecord = false;
  }
  else
  {
    sheet_control.getRange(11,1).setValue(lastRecord);
    var range_rawdata = sheet_rawdata.getRange("A2:A5000");
    var row = range_rawdata.getValues();
    var lastRecord = 1;
    
    for (var i=0;i<row.length;i++)
    {
      var tstamp = row[i][0];
      
      if (tstamp != "")
      {
        lastRecord = i + 2;
      }
    }
    
    sheet_control.getRange(13,1).setValue(lastRecord);
  }
}