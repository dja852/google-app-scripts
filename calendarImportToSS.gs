function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Calendar Import")
    .addItem("Get Calendar IDs", "getCalendarIds")
    .addItem("Get Calendar Data", "getCalendarEventsById")
    .addToUi();

}

function getCalendarIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssSheetIDs = ss.getSheetByName("Calendar IDs");
  var cals = CalendarApp.getAllCalendars();

  for (var i = 0; i < cals.length; i++) {

    ssSheetIDs.getRange(i + 2, 1).setValue(cals[i].getName());
    ssSheetIDs.getRange(i + 2, 2).setValue(cals[i].getId());

  }
}

function getCalendarEventsById() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  var ssSheetConfig = ss.getSheetByName("Import Config");
  var ssSheetData = ss.getSheetByName("Calendar Data");

  var cal = CalendarApp.getCalendarById(ssSheetConfig.getRange(1, 2).getValue());
  var startTime = ssSheetConfig.getRange(2, 2).getValue();
  var endTime = ssSheetConfig.getRange(3, 2).getValue();

  var calEvents = cal.getEvents(new Date(startTime), new Date(endTime));

  for (var i = 0; i < calEvents.length; i++) {

    ssSheetData.getRange(i + 2, 1).setValue(calEvents[i].getTitle());
    ssSheetData.getRange(i + 2, 2).setValue(calEvents[i].getDescription());
    ssSheetData.getRange(i + 2, 3).setValue(calEvents[i].isAllDayEvent());

  }
  
}
