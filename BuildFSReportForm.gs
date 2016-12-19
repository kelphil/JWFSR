function createForm()
{ 
  var formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  
  if(formUrl == null)
  {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
    // Create and open a form.
    var newForm = FormApp.create('Field Service Report');

    // Update the form's response destination.
    newForm.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
    
    // Hide form response sheet
    var sheet_form = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  
    sheet_form.hideSheet();
  }
 
  var sheet_control = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Controls');
  
  sheet_control.getRange(11,1).setValue(0);
  sheet_control.getRange(13,1).setValue(0);
  
  buildForm();
}

function buildForm()
{
  var sheet_form = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sheet_form.hideSheet();
  
  var formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  //---------------------------------------------------------------------
  // Get Month Names
  //---------------------------------------------------------------------
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('Months'));
  var sheet = spreadsheet.getActiveSheet();
  
  var range = sheet.getRange("A3:E14");
  var row = range.getValues();
  
  var months = [];
  
  for (var i = 0; i < row.length; i++)
  { 
    var month = row[i][2];
    var delta = row[i][4];
    
    if ((delta < 0) && (delta >= -1))
    {
      if (months.indexOf(month) < 0)
      {
        months.push(month);
      }
    }
  }
  
  
  //---------------------------------------------------------------------
  // Get Publisher Names
  //---------------------------------------------------------------------
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('Controls'));
  var sheet_control = spreadsheet.getActiveSheet();
  
  var range_control = sheet_control.getRange("A1:F16");
  var row_control = range_control.getValues();
  
  var full_names = [];
  var publisher_names = [];
  var email_address_check_column_header = row_control[15][5];
  
  var mailing_list = getEmailAddresses(email_address_check_column_header);
  
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
  
 
  //---------------------------------------------------------------------
  // Get Existing Content
  //---------------------------------------------------------------------
  var blankform = true;
  var items = form.getItems();
  
  if(items.length > 0)
  {
    blankform = false;
  }
  
  if(1)
  {
    
    form.setAcceptingResponses(true);
    
    if(blankform)
    {
      //---------------------------------------------------------------------
      // Form Setup
      //---------------------------------------------------------------------
      
      var group_name = row_control[13][1] + ' ' + row_control[13][3];
      form.setTitle('Field Service Report')
      .setDescription(group_name)
      .setConfirmationMessage('Thanks for your field service report! You should receive an email shortly.')
      .setAllowResponseEdits(false)
      .setShowLinkToRespondAgain(false)
      .setLimitOneResponsePerUser(false)
      .setAcceptingResponses(true)
      .setProgressBar(false);
    }
    else
    {
      
      //---------------------------------------------------------------------
      // Delete Existing Content
      //---------------------------------------------------------------------
      //    form.deleteItem(0);
    }
    
    //---------------------------------------------------------------------
    // Publisher Name - Dropdown Item
    //---------------------------------------------------------------------
    if(blankform)
    {
      var pname = form.addListItem();
      pname.setRequired(true);
      pname.setTitle('Publisher Name');
      pname.setChoiceValues(publisher_names)
    }
    else
    {
      items[0].asListItem().setChoiceValues(publisher_names);
    }
    
    //---------------------------------------------------------------------
    // Month - Dropdown Item
    //---------------------------------------------------------------------
    if(blankform)
    {
      var mon = form.addListItem();
      mon.setRequired(true);
      mon.setTitle('Month');
      mon.setChoiceValues(months);
      
      var scode = form.addTextItem();
      scode.setTitle('Security Code');
      scode.setRequired(true);
      
      var pageBreak = form.addPageBreakItem();
      var section = form.addSectionHeaderItem();
      section.setTitle('Report Details');
    }
    else
    {
      items[1].asListItem().setChoiceValues(months);
    }
    
    //---------------------------------------------------------------------
    // Report Details
    //---------------------------------------------------------------------    
    if(blankform)
    {
      var item = form.addTextItem();
      item.setTitle('Placements (Printed and Electronic)');
      item.createResponse('0');
      
      item = form.addTextItem();
      item.setTitle('Video Showings');
      item.createResponse('0');
      
      item = form.addTextItem();
      item.setTitle('Hours');
      item.setRequired(true);
      
      item = form.addTextItem();
      item.setTitle('Return Visits');
      item.createResponse('0');
      
      item = form.addTextItem();
      item.setTitle('Number of Different Bible Studies Conducted');
      item.createResponse('0');
      
      item = form.addParagraphTextItem();
      item.setTitle('Comments');
      item.createResponse('0');
    }
  }
  
  var sheet_months = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Months');
 
  sheet_months.hideSheet();
  
};