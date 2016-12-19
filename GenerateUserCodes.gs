function generateUserCodes()
{  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('Email List'));
  var sheet = spreadsheet.getActiveSheet();
  
  var range = sheet.getRange("A1:B200");
  var row = range.getValues();
  var user_codes = [];
  
  for (var i=1;i<row.length;i++)
  {
    if (row[i][1] != "")
    {
      var uniquecode = false;
      
      while (!uniquecode)
      {
        var rand_code = getRandomIntInclusive(1000,9999);
        
        if (user_codes.indexOf(rand_code) < 0)
        {
          sheet.getRange(i+1,11).setValue(rand_code)
          user_codes.push(rand_code);
          uniquecode = true;
        }
      }
    }
  }
}