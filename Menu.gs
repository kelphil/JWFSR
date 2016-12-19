function onOpen()
{
  var subMenus = [
    {name:"Send FSR Summary Email", functionName: "sendFSRSummaryEmail"},
    {name:"Update Dashboard", functionName: "updateDashboard"},
    {name:"Create Form", functionName: "createForm"},
    {name:"Build Form", functionName: "buildForm"},
    //{name:"Stop Form", functionName: "stopForm"}
    {name:"Create Triggers", functionName: "createTriggers"}
  ];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("FSR Menu", subMenus);
}