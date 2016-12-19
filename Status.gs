function updateStatus()
{
  var existingReport = new findExistingReports(); 
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  var sheet_control = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Controls');
  var sheet_status = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Status');
  
  var range_control = sheet_control.getRange("A1:F16");
  var row_control = range_control.getValues();
  
  var fsr_month = row_control[1][3];
  var fsr_year = row_control[2][3];
  var email_address_check_column_header = row_control[15][5];
  
  var mailing_list = getEmailAddresses(email_address_check_column_header);
  
  var publisher_names = [];
  var names_db = [];
  var full_names = [];
  
  for (var i=0;i<mailing_list.getCount();i++)
  {
    if(mailing_list.enabled[i])
    {
      var first_name = mailing_list.first_name[i];
      var last_name = mailing_list.last_name[i];
      var name = last_name + " " + first_name;
      full_names.push(name);
    }
  }
  
  var sindex = sortList(full_names);
  
  for (var j=0;j<sindex.length;j++)
  {
    var name = full_names[sindex[j]];
    var names = name.split(" ");
    name = names[1] + " " + names[0];
    publisher_names.push(name);
  }
  
  var range = sheet_db.getRange("A2:C200");
  var row = range.getValues();
  
  for (var i=0;i<row.length;i++)
  {
    var year = row[i][0];
    var month = row[i][1];
    var name = row[i][2];
    
    if ((year == fsr_year) && (month == fsr_month))
    {
      if (name != "")
      {
        names_db.push(name);
      }
    }
  }
  
  var pcount = 0;
  var rcount = 0;
  
  for (var i=0;i<publisher_names.length;i++)
  {
    var name = publisher_names[i];
    var pname = swapName(name).replace(" ",", ");
    pcount++;
    
    sheet_status.getRange(i+4,1).setValue(pname);
    
    if (names_db.indexOf(name) >= 0)
    {
      sheet_status.getRange(i+4,2).setValue('YES');
      rcount++;
    }
    else
    {
      sheet_status.getRange(i+4,2).setValue('NO');
    }
  }
  
  var total_status_rows =  publisher_names.length + 3;
  var last_row = sheet_status.getLastRow();
  var num_rows_to_be_deleted = last_row - total_status_rows;
   
  if(num_rows_to_be_deleted)
  {
    sheet_status.deleteRows(total_status_rows+1, num_rows_to_be_deleted);
  }
   
  var progress = (Math.floor((rcount/pcount)*100))/100;
  sheet_status.getRange(1,2).setValue(progress);
  sheet_status.getRange(2,2).setValue("\'"+rcount+"/"+pcount);
  
};