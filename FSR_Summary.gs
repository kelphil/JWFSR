function fsReportSummary(summary_email)
{ 
  findMonthBoundaries();
  
  var today = new Date();
  var date = Utilities.formatDate(today, "PST", "MMMM dd, YYYY");
  today = today.getDate();
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet_db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  var sheet_control = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Controls');
  var sheet_status = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Status');
  
  var range_control = sheet_control.getRange("A1:F16");
  var row_control = range_control.getValues();
  
  var fsr_month = row_control[1][3];
  var fsr_year = row_control[2][3];
  var first_day = row_control[14][5];
  var last_day = row_control[14][3] + 1;
  var group_overseer_name = row_control[15][1];
  var group_overseer_email = row_control[15][3];
  var email_address_check_column_header = row_control[15][5];
  var row_offset = row_control[8][0];

  var mailing_list = getEmailAddresses(email_address_check_column_header);
  var pcount = mailing_list.getCount();
  var reportComplete = false;
  
  var header_index = new headerRow;
  var range_report = sheet_db.getRange("A1:J1");
  var row = range_report.getValues();
  
  for (var i=0;i<row[0].length;i++)
  {
    var header = row[0][i];
    
    if (header != "")
    {      
      header_index.addEntry(header);
    }
  }
  
  var range = sheet_db.getRange("A2:J200");
  var row = range.getValues();
  var report_summary = [];
  var pname,email,phone,year,month,hours,mag,vid,rv,bs,comment = "";
  var reportcount = 0;
  var tot_hours = 0;
  var tot_mag = 0;
  var tot_vid = 0;
  var tot_rv = 0;
  var tot_bs = 0;
  var avg_hours = 0;
  var avg_mag = 0;
  var avg_vid = 0;
  var avg_rv = 0;
  var avg_bs = 0;
  var del_offset = 0;
  
  for (var i=0;i<row.length;i++)
  {
    var index = header_index.getIndex('Year');
    year = row[i][index];
    
    index = header_index.getIndex('Month');
    month = row[i][index];
    
    if ((year == fsr_year) && (month == fsr_month))
    {
      index = header_index.getIndex('Publisher Name');
      pname = row[i][index];
      
      index = header_index.getIndex('Email');
      email = row[i][index];
      
      if ((pname == 'TOTAL') || (pname == 'AVERAGE'))
      {
        sheet_db.deleteRow(i+2-del_offset);
        del_offset++;
      }
      else
      { 
        index = header_index.getIndex('Placements (Printed and Electronic)');
        mag = row[i][index];
        tot_mag += mag; 
        
        index = header_index.getIndex('Video Showings');
        vid = row[i][index];
        tot_vid += vid;
        
        index = header_index.getIndex('Hours');
        hours = row[i][index];
        tot_hours += hours;
        
        index = header_index.getIndex('Return Visits');
        rv = row[i][index];
        tot_rv += rv;
        
        index = header_index.getIndex('Number of Different Bible Studies Conducted');
        bs = row[i][index];
        tot_bs += bs;
        
        index = header_index.getIndex('Comments');
        comment = row[i][index];
        
        var report = new fieldServiceReport;
        
        report.addEntry("Publisher",pname);
        report.addEntry("Placements (Printed and Electronic)",mag);
        report.addEntry("Video Showings",vid);
        report.addEntry("Hours",hours);
        report.addEntry("Return Visits",rv);
        report.addEntry("Number of Different Bible Studies Conducted",bs);
        report.addEntry("Comments",comment);
        
        report_summary[pname] = report;
        
        reportcount++;
      }
    }
  }
  
  if (reportcount == pcount)
  {
    reportComplete = true;
  }
  
  // Record Totals for the month
  
  sheet_db.insertRowAfter(row_offset);
  row_offset++;
  
  var cell = sheet_db.getRange(row_offset,1,1,10);
  cell.setBackgroundRGB(255,221,153);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("center");
  cell = sheet_db.getRange(row_offset,3,1,2).merge();
  
  sheet_db.getRange(row_offset,1).setValue(fsr_year);
  sheet_db.getRange(row_offset,2).setValue(fsr_month);
  sheet_db.getRange(row_offset,3).setValue('TOTAL');
  sheet_db.getRange(row_offset,4).setValue('');
  sheet_db.getRange(row_offset,5).setValue(tot_mag);
  sheet_db.getRange(row_offset,6).setValue(tot_vid);
  sheet_db.getRange(row_offset,7).setValue(tot_hours);
  sheet_db.getRange(row_offset,8).setValue(tot_rv);
  sheet_db.getRange(row_offset,9).setValue(tot_bs);
  sheet_db.getRange(row_offset,10).setValue('');
  
  // Record averages for the month
  
  sheet_db.insertRowAfter(row_offset);
  row_offset++;
  
  avg_mag = (tot_mag/reportcount).toFixed(1);
  avg_vid = (tot_vid/reportcount).toFixed(1);
  avg_hours = (tot_hours/reportcount).toFixed(1);
  avg_rv = (tot_rv/reportcount).toFixed(1);
  avg_bs = (tot_bs/reportcount).toFixed(1);
  
  var cell = sheet_db.getRange(row_offset,1,1,10);
  cell.setBackgroundRGB(255,238,204);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("center");
  cell = sheet_db.getRange(row_offset,3,1,2).merge();
  
  sheet_db.getRange(row_offset,1).setValue(fsr_year);
  sheet_db.getRange(row_offset,2).setValue(fsr_month);
  sheet_db.getRange(row_offset,3).setValue('AVERAGE');
  sheet_db.getRange(row_offset,4).setValue('');
  sheet_db.getRange(row_offset,5).setValue(avg_mag);
  sheet_db.getRange(row_offset,6).setValue(avg_vid);
  sheet_db.getRange(row_offset,7).setValue(avg_hours);
  sheet_db.getRange(row_offset,8).setValue(avg_rv);
  sheet_db.getRange(row_offset,9).setValue(avg_bs);
  sheet_db.getRange(row_offset,10).setValue('');
  
  if(summary_email)
  {
    var emailSubject = 'Field Service Report Summary - ' + getFullMonthName(fsr_month) + ' ' + fsr_year;
    
    if (((today >= first_day) && (today <= last_day)) || ((reportComplete == true) && (today <= last_day)))
    {
      if(today == last_day)
      {
        emailSubject = 'Final Field Service Report Summary - ' + getFullMonthName(fsr_month) + ' ' + fsr_year;
      }
    
      // Prepare Content for Email Summary
      var color_header = '#333333';
      var color_data = '#f2f2f2';
      
      var fsr_summary = '<body><h2>Field Service Report Summary - ' + getFullMonthName(fsr_month) + " " + fsr_year + '</h2><p>';
      fsr_summary = fsr_summary + '<table cellspacing="0" cellpadding="4" border="1">' + '<tr>';
      fsr_summary = fsr_summary + '<tr>'
      + '<td bgcolor=\"' + color_header + '\" align="right"><font color=\"#ffffff\">' + "Publisher Name" + '</font></td>'
      + '<td bgcolor=\"' + color_header + '\" align="right"><font color=\"#ffffff\">' + "Phone Number" + '</font></td>'
      + '<td bgcolor=\"' + color_header + '\" align="right"><font color=\"#ffffff\">' + "Placements (Printed and Electronic)" + '</font></td>'
      + '<td bgcolor=\"' + color_header + '\" align="right"><font color=\"#ffffff\">' + "Video Showings" + '</font></td>'
      + '<td bgcolor=\"' + color_header + '\" align="right"><font color=\"#ffffff\">' + "Hours" + '</font></td>'
      + '<td bgcolor=\"' + color_header + '\" align="right"><font color=\"#ffffff\">' + "Return Visits" + '</font></td>'
      + '<td bgcolor=\"' + color_header + '\" align="right"><font color=\"#ffffff\">' + "Number of Different Bible Studies Conducted" + '</font></td>'
      + '<td bgcolor=\"' + color_header + '\" align="right"><font color=\"#ffffff\">' + "Comments" + '</font></td>'
      +'</tr>';
      
      
      var full_name = [];
      
      for (var i=0;i<pcount;i++)
      {
        var first_name = mailing_list.first_name[i];
        var last_name = mailing_list.last_name[i];
        var name = last_name + " " + first_name;
        full_name.push(name);
      }
      
      var sindex = sortList(full_name); 
      
      for (var j=0;j<sindex.length;j++)
      {
        var name = swapName(full_name[sindex[j]]);
        var email = mailing_list.getEmail(name);
        var i = mailing_list.getIndex(name);
        
        if(mailing_list.enabled[i])
        {
          var first_name = mailing_list.first_name[i];
          var last_name = mailing_list.last_name[i];
          name = last_name + ", " + first_name;
          pname = first_name + " " + last_name;
          var phone = mailing_list.phone[i];
          
          if (report_summary[pname] != undefined)
          {
            mag = report_summary[pname].getValue('Placements (Printed and Electronic)');
            vid = report_summary[pname].getValue('Video Showings');
            hours = report_summary[pname].getValue('Hours');
            rv = report_summary[pname].getValue('Return Visits');
            bs = report_summary[pname].getValue('Number of Different Bible Studies Conducted');
            comment = report_summary[pname].getValue('Comments');
            
            fsr_summary = fsr_summary + '<tr>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + name + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + phone + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + mag + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + vid + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + hours + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + rv + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + bs + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + comment + '</font></td>'
            + '</tr>';
          }
          else
          {
            fsr_summary = fsr_summary + '<tr>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + name + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + phone + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + "" + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + "" + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + "" + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + "" + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + "" + '</font></td>'
            + '<td bgcolor=\"' + color_data + '\" align="center"><font color=\"#000000\">' + "" + '</font></td>'
            + '</tr>';
          }
        }
      }
      
      fsr_summary = fsr_summary + '</table>';
      fsr_summary = fsr_summary + '</p></body>';
    
      fsReportSummaryEmail(emailSubject,fsr_summary,mailing_list);
    }
  }
};