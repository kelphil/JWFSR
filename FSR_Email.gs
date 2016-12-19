function sendReportReminderEmail()
{
  var today = new Date();
  today = today.getDate();
  
  var existingReport = new findExistingReports(); 
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_control = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Controls');
  
  var range_control = sheet_control.getRange("A1:F16");
  var row_control = range_control.getValues();
  
  var fsr_month = row_control[1][3];
  var fsr_year = row_control[2][3];
  var first_day = row_control[14][1];
  var last_day = row_control[14][3];
  var additional_notes = row_control[16][1];
  var email_address_check_column_header = row_control[15][5];
  
  var mailing_list = getEmailAddresses(email_address_check_column_header);

  for (var i=0;i<mailing_list.getCount();i++)
  {
    if(mailing_list.enabled[i])
    {
      var first_name = mailing_list.first_name[i];
      var last_name = mailing_list.last_name[i];
      var name = first_name + " " + last_name;
      var email = mailing_list.email[i];  
      var usercode = mailing_list.code[i];  
      var emailSubject, message = "";
      
      if (!existingReport.ifExists(name))
      {
        if(1)
        { 
          if ((today >= first_day) && (today <= last_day))
          {
            // Send out email on alternate days (1,3,5,7 etc)
            if(today == first_day)
            {
              emailSubject = "Field Service Report - " + getFullMonthName(fsr_month) + " " + fsr_year;
              message =  "Hi " + first_name + ",<br><br>It is time to send in your field service report for " + getFullMonthName(fsr_month) + " " + fsr_year + ". Please click the link below to submit your field service report; your security code is <b>" + usercode + "</b>.";
            }
            else if(today == last_day)
            {
              emailSubject = "Final Reminder | Field Service Report - " + getFullMonthName(fsr_month) + " " + fsr_year; 
              message =  "Hi " + first_name + ",<br><br>Today is the last day to send in your field service report for " + getFullMonthName(fsr_month) + " " + fsr_year + ". Your field service report cannot be forwarded to the branch if it's not received by the end of this day. Please click the link below to submit your field service report; your security code is <b>" + usercode + "</b>.";
            }
            else if((today & 1) == 1)
            { 
              emailSubject = "Reminder | Field Service Report - " + getFullMonthName(fsr_month) + " " + fsr_year; 
              message =  "Hi " + first_name + ",<br><br>Here is a gentle reminder to send in your field service report for " + getFullMonthName(fsr_month) + " " + fsr_year + ". Please click the link below to submit your field service report; your security code is <b>" + usercode + "</b>.";
            }
            else
            {
              emailSubject = "";
              message = "";
            }
            
            if (emailSubject != "")
            {
              message = message + '<br><br><a href="http://goo.gl/forms/hzCcd1xzO8">Field Service Report Form</a>';
              
              if (additional_notes != "")
              {
                message = message + '<br><br><font color="red">' + additional_notes + '</font>';
              }
              
              message = message + "<br><br>Thanks for your efforts!";
              
              var emails = email; 
              var emailBody = message + "<br>";
              
              MailApp.sendEmail(
                {
                  to: emails,
                  subject: emailSubject,
                  htmlBody: emailBody
                }
              );
            }
          }
        }
      }
    }
  }
};
