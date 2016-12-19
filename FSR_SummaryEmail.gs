function fsReportSummaryEmail(emailSubject,fsr_summary,mailing_list)
{ 
  
  var today = new Date();
  var date = Utilities.formatDate(today, "PST", "MMMM dd, YYYY");
  today = today.getDate();
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_control = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Controls');
  
  var range_control = sheet_control.getRange("A1:F16");
  var row_control = range_control.getValues();
  
  var fsr_month = row_control[1][3];
  var fsr_year = row_control[2][3];
  var group_overseer_name = row_control[15][1];
  var group_overseer_email = row_control[15][3];
    
  if(1)
  { 
    var names = group_overseer_name.split(" ");
    var group_overseer_first_name = names[0];
    var group_overseer_last_name = names[1];
    var phone = mailing_list.getPhone(group_overseer_name);
    
    var emails = group_overseer_email; 
    
    var message =  "Hi " + group_overseer_first_name + ",<br><br>Here is the field service report summary for " + getFullMonthName(fsr_month) + " " + fsr_year + " as of " + date + ".";
    message = message + '<h3>Group Overseer: ' + group_overseer_name + '   ' + phone + '</h3>';
    message = message + fsr_summary;
    
    var emailBody = message + "<br>";
    
    MailApp.sendEmail(
      {
        to: emails,
        subject: emailSubject,
        htmlBody: emailBody
      }
    );
  }
};