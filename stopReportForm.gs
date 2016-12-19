function stopForm()
{
  
  var formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  //---------------------------------------------------------------------
  // Form Setup
  //---------------------------------------------------------------------
  form.setAcceptingResponses(false)
  .setCustomClosedFormMessage('The field service report form is closed for the month. Please contact your group overseer for more details.');
};