function findMonthBoundaries()
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet_control = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Controls');
  var sheet_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  var range_control = sheet_control.getRange("A1:F10");
  var row_control = range_control.getValues();
  
  var fsr_month = row_control[1][3];
  var fsr_year = row_control[2][3];
  var prev_month = row_control[1][5];
  var prev_year = row_control[2][5];
  var prev_month_offset = 0;
  
  var fsr_month_found = false;
  var prev_month_found = false;
    
  var range_db = sheet_db.getRange("A2:C5000");
  var row_db = range_db.getValues();
  
  var header_offset = 2;
  var fsr_month_begin = header_offset;
  var fsr_month_end = 2;
  var prev_month_begin = sheet_db.getLastRow();
  
  for (var i=0;i<row_db.length;i++)
  {
    var y = row_db[i][0];
    var m = row_db[i][1];
    var e = row_db[i][2];
    
    if((y == fsr_year) && (m == fsr_month))
    {
      if(!fsr_month_found)
      {
        fsr_month_begin += + i;
        fsr_month_found = true;
      }
      if((e == 'TOTAL') || (e == 'AVERAGE'))
      {
        prev_month_offset--;
      }
      fsr_month_end = i + header_offset;
    }
    else if((y == prev_year) && (m == prev_month) && (!prev_month_found))
    {
      
      prev_month_begin = i + header_offset;
      prev_month_found = true;
    }
  }
  
  prev_month_begin += prev_month_offset;
  fsr_month_end += prev_month_offset;
  
  sheet_control.getRange(7,1).setValue(fsr_month_begin);
  sheet_control.getRange(9,1).setValue(fsr_month_end);
}