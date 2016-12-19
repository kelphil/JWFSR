function createTriggers()
{
  deleteAllTriggers();
  
  var formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  
  ScriptApp.newTrigger('generateUserCodes')
  .timeBased()
  .atHour(1)
  .onMonthDay(1)
  .create();
  
  ScriptApp.newTrigger('buildForm')
  .timeBased()
  .atHour(2)
  .onMonthDay(1)
  .create();
  
  ScriptApp.newTrigger('updateDashboard')
  .forForm(form)
  .onFormSubmit()
  .create();
  
  ScriptApp.newTrigger('updateDashboard')
  .timeBased()
  .everyHours(1)
  .create();
  
  ScriptApp.newTrigger('sendReportReminderEmail')
  .timeBased()
  .atHour(8)
  .everyDays(1)
  .create();
  
  ScriptApp.newTrigger('sendFSRSummaryEmail')
  .timeBased()
  .atHour(10)
  .everyDays(1)
  .create();
  
}

function updateDB() {
  var form = FormApp.openById('1ownwFPCmW8EM0asgjPQ2ZEkNge4t8akyONVwUMT_-sE');
  
  ScriptApp.newTrigger('updateDashboard')
  .forForm(form)
  .onFormSubmit()
  .create();
}

function deleteAllTriggers()
{
  // Loop over all triggers.
  var allTriggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < allTriggers.length; i++)
  {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}

function deleteTrigger(triggerId)
{
  // Loop over all triggers.
  var allTriggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < allTriggers.length; i++)
  {
    // If the current trigger is the correct one, delete it.
    if (allTriggers[i].getUniqueId() == triggerId)
    {
      ScriptApp.deleteTrigger(allTriggers[i]);
      break;
    }
  }
}