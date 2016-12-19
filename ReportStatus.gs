function findExistingReports()
{
  findMonthBoundaries();
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_control = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Controls');
  var sheet_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  var range_control = sheet_control.getRange("A1:F10");
  var row_control = range_control.getValues();
  
  var fsr_month = row_control[1][3];
  var fsr_year = row_control[2][3];
  var prev_month = row_control[1][5];
  var prev_year = row_control[2][5];
  var fsr_month_begin = row_control[6][0];
  var prev_month_begin = row_control[8][0];
  
  var existingReport = new existingreport();
    
  var range_db = sheet_db.getRange("A2:C5000");
  var row_db = range_db.getValues();
  
  for (var i=0;i<row_db.length;i++)
  {
    var year = row_db[i][0];
    var mon = row_db[i][1];
    var name = row_db[i][2];
    
    if((year == fsr_year) && (mon == fsr_month))
    {
      existingReport.addEntry(name);
    }
  }
  
  return existingReport;
}