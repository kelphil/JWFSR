function updateDashboard()
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet_control = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Controls');
  var sheet_report = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FS Report');
  var sheet_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  var sheet_months = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Months');
 
  sheet_months.hideSheet();
    
  var header_index = new headerRow;
  var range_report = sheet_report.getRange("A1:Q1");
  var row = range_report.getValues();
  
  for (var i=0;i<row[0].length;i++)
  {
    var header = row[0][i];
    
    if (header != "")
    {      
      header_index.addEntry(header);
    }
  }
 
  if(1)
  {
    findMonthBoundaries();
    var newRecord = getNewRecords();
  
    if (newRecord)
    {
      var range_control = sheet_control.getRange("A1:F16");
      var row_control = range_control.getValues();
      
      var fsr_month = row_control[1][3];
      var fsr_year = row_control[2][3];
      var fsr_month_begin = row_control[6][0];
      var fsr_month_end = row_control[8][0];
      var lastRecord_rawdata = row_control[12][0];
      var email_address_check_column_header = row_control[15][5];
      
      var mailing_list = getEmailAddresses(email_address_check_column_header);
      
      var codes = getUserCodes(mailing_list);
      
      var range = sheet_report.getRange("A1:R5000");
      var row = range.getValues();
      var row_offset = fsr_month_begin;
      var pname,email,correct_name,correct_email,year,month,codematch,usercode,hours,mag,vid,rv,bs,comment = "";
      var updated_body,updated_header,bs_rv_check = "";
      
      var existingReport = new findExistingReports();
      var valid_row = new validRows;
      var lastRecord = lastRecord_rawdata;
      
      for (var i=lastRecord_rawdata;i<row.length;i++)
      { 
        var index = header_index.getIndex('Publisher Name');
        pname = row[i][index];
        
        if (pname != "")
        {
          lastRecord = i + 1;
          
          index = header_index.getIndex('Year');
          year = row[i][index];
          
          index = header_index.getIndex('Month');
          month = row[i][index];
          
          if ((year == fsr_year) && (month == fsr_month))
          {
            index = header_index.getIndex('Security Code');
            var enteredcode = row[i][index];
            
            usercode = codes.getCode(pname);
            
            sheet_report.getRange(lastRecord,8).setValue(usercode);
            
            if (enteredcode == usercode)
            {
              codematch = 1;
            }
            else
            {
              codematch = 0;
            }
        
            sheet_report.getRange(lastRecord,6).setValue(codematch);
            
            email = mailing_list.getEmail(pname);
            sheet_report.getRange(lastRecord,3).setValue(email);
            
            if(codematch == 1)
            {
              valid_row.addEntry(year,month,pname,email,codematch,i);
            }
            else
            { 
              if(codes.isValid(enteredcode))
              {
                correct_name = codes.getName(enteredcode);
                correct_email = codes.getEmail(correct_name);
                
                if (pname != correct_name)
                {
                  if(1)
                  { 
                    var names = correct_name.split(" ");
                    var first_name = names[0];
                    
                    var emails = correct_email; 
                    var emailSubject = "Field Service Report - " + getFullMonthName(fsr_month) + " " + fsr_year + " - Incorrect Name Entered";
                    
                    var formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
                    
                    var message =  "Hi " + first_name + ",<br><br>You selected an incorrect publisher name while submitting the " + getFullMonthName(fsr_month) + " " + fsr_year + " field service report. Please submit the report details again using your security code : <b>" + enteredcode + "</b>.";
                    message = message + '<br><br><a href="' + formUrl + '">Field Service Report Form</a>';
                    message = message + "<br><br>Thanks for your efforts!";
                    
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
              else
              {
                index = header_index.getIndex('Email');
                email = row[i][index];
                
                usercode = codes.getCode(pname);
                
                if(1)
                { 
                  var names = pname.split(" ");
                  var first_name = names[0];
                  
                  var emails = email; 
                  var emailSubject = "Field Service Report - " + getFullMonthName(fsr_month) + " " + fsr_year + " - Wrong Security Code Entered";
                  
                  var formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
                  
                  var message =  "Hi " + first_name + ",<br><br>You entered a wrong security code while submitting the " + getFullMonthName(fsr_month) + " " + fsr_year + " field service report. Please submit the report details again using your correct security code : <b>" + usercode + "</b>.";
                  message = message + '<br><br><a href="' + formUrl + '">Field Service Report Form</a>';
                  message = message + "<br><br>Thanks for your efforts!";
                  message = message + "<br><br><br><font size='1' color='#737373'><b>Note:</b> If you think you recieved this email by mistake, please ignore this message.</font>";
                  
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
      
      sheet_control.getRange(13,1).setValue(lastRecord);     
      
      for (var k=0;k<valid_row.getCount();k++)
      {
        pname = valid_row.name[k];
        email = valid_row.email[k];
        codematch = valid_row.codematch[k];
        
        var i = valid_row.getRow(pname);
        
        if (pname != "")
        { 
          index = header_index.getIndex('Year');
          year = row[i][index];
          
          index = header_index.getIndex('Month');
          month = row[i][index];
                    
          if ((year == fsr_year) && (month == fsr_month))
          { 
            if(codematch == 1)
            {
              index = header_index.getIndex('Placements (Printed and Electronic)');
              mag = row[i][index];
              
              index = header_index.getIndex('Video Showings');
              vid = row[i][index];
              
              index = header_index.getIndex('Hours');
              hours = row[i][index];
              
              index = header_index.getIndex('Return Visits');
              rv = row[i][index];
              
              index = header_index.getIndex('Number of Different Bible Studies Conducted');
              bs = row[i][index];
              
              index = header_index.getIndex('Comments');
              comment = row[i][index];
              
              if (existingReport.ifExists(pname))
              {
                row_offset = existingReport.getRow(pname);
                updated_body = " updated";
                updated_header = " Updated";
              }
              else
              {
                sheet_db.insertRowBefore(row_offset);
                updated_body = "";
                updated_header = "";
              }
              
              var report = new fieldServiceReport;
              
              report.addEntry("Year",year);
              report.addEntry("Month",month);
              report.addEntry("Publisher",pname);
              report.addEntry("Email",email);
              report.addEntry("Placements (Printed and Electronic)",mag);
              report.addEntry("Video Showings",vid);
              report.addEntry("Hours",hours);
              report.addEntry("Return Visits",rv);
              report.addEntry("Number of Different Bible Studies Conducted",bs);
              report.addEntry("Comments",comment);
              
              if (rv < bs)
              {
                bs_rv_check = "The number of return visits is less than the number of different bible studies conducted as per your field service report. Please ";  
                report.setError("Return Visits");
              }
              
              sheet_db.getRange(row_offset,1).setValue(report.getValue('Year'));
              sheet_db.getRange(row_offset,2).setValue(report.getValue('Month'));
              sheet_db.getRange(row_offset,3).setValue(report.getValue('Publisher'));
              sheet_db.getRange(row_offset,4).setValue(report.getValue('Email'));
              sheet_db.getRange(row_offset,5).setValue(report.getValue('Placements (Printed and Electronic)'));
              sheet_db.getRange(row_offset,6).setValue(report.getValue('Video Showings'));
              sheet_db.getRange(row_offset,7).setValue(report.getValue('Hours'));
              sheet_db.getRange(row_offset,8).setValue(report.getValue('Return Visits'));
              sheet_db.getRange(row_offset,9).setValue(report.getValue('Number of Different Bible Studies Conducted'));
              sheet_db.getRange(row_offset,10).setValue(report.getValue('Comments'));
              
              var fs_report = '<body><h2>Your ' + updated_header + ' Field Service Report - ' + getFullMonthName(fsr_month) + " " + fsr_year + '</h2><p>';
              fs_report = fs_report + '<table cellspacing="0" cellpadding="10" border="1">' + '<tr>';
              
              for (var j=4;j<report.getCount();j++)
              {
                var item = report.item[j];
                var value = report.value[j];
                var valid = report.valid[j];
                var error = report.error[j];
                
                if(valid)
                {
                  fs_report = fs_report + '<tr><td bgcolor=\"#404040\" align="right"><font color=\"#ffffff\">' + item + '</font></td>';
                  
                  if (error)
                  {
                    fs_report = fs_report + '<td bgcolor=\"#ff8080\" align="center"><font color=\"#000000\">' + value + '</font></td></tr>';
                  }
                  else
                  {
                    fs_report = fs_report + '<td bgcolor=\"#f2f2f2\" align="center"><font color=\"#000000\">' + value + '</font></td></tr>';
                  }
                }
              }
              fs_report = fs_report + '</tr>' + '</table>';
              fs_report = fs_report + '</p></body>';
              
              if(1)
              { 
                var names = pname.split(" ");
                var first_name = names[0];
                
                usercode = codes.getCode(pname);
                
                var emails = email; 
                var emailSubject = "Your" + updated_header + " Field Service Report - " + getFullMonthName(fsr_month) + " " + fsr_year;
                
                var formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
                
                var message =  "Hi " + first_name + ",<br><br>Thanks for sending your" + updated_body + " field service report.";
                message = message + "<br><br>" + fs_report;
                message = message + "<br>"
                
                if (bs_rv_check != "")
                {
                  message = message + bs_rv_check;
                }
                else
                {
                  message = message + "Please check if the above information is accurate. If not, please ";
                }
                
                message = message + "resubmit the complete report with the correct information using the link below. Your security code is <b>" + usercode + "</b>.";
                message = message + '<br><br><a href="' + formUrl + '">Resubmit Field Service Report</a>';
                message = message + "<br><br>Thanks for your efforts!";
                
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
      updateStatus();
      fsReportSummary(false);
    }
  }
}