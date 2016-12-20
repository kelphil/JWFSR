function generateUserCodes()
{  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('Email List'));
  var sheet = spreadsheet.getActiveSheet();
  
  var range = sheet.getRange("A1:Z200");
  var row = range.getValues();
  var user_codes = []; 
  var header_row = [];
  
  for (var i=0;i<11;i++)
  {
    var header = row[0][i];
    if(header != undefined)
    {
      header_row.push(header)
    }
  }
  
  var idx_code = header_row.indexOf("USERCODE");
  
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
          sheet.getRange(i+1,idx_code+1).setValue(rand_code)
          user_codes.push(rand_code);
          uniquecode = true;
        }
      }
    }
  }
}
