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

  calList = []

  for (var i in cals) {
    calList.push([
      cals[i].getName(),
      cals[i].getId()
    ])
  }

  ssSheetIDs.getRange(2, 1, calList.length, 2).setValues(calList)

}

function getCalendarEventsById() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  var ssSheetConfig = ss.getSheetByName("Import Config");
  var ssSheetData = ss.getSheetByName("Calendar Data");

  var cal = CalendarApp.getCalendarById(ssSheetConfig.getRange(1, 2).getValue());
  var startTime = ssSheetConfig.getRange(2, 2).getValue();
  var endTime = ssSheetConfig.getRange(3, 2).getValue();

  var calEvents = cal.getEvents(new Date(startTime), new Date(endTime));

  var calData = [];

  for (var i in calEvents) {

    var isAllDayEvent = calEvents[i].isAllDayEvent();
    var guestList = calEvents[i].getGuestList(true);

    calData.push([
      calEvents[i].getTitle(),
      calEvents[i].getDescription(),
      isAllDayEvent,
      isAllDayEvent ? calEvents[i].getAllDayStartDate() : calEvents[i].getStartTime(),
      isAllDayEvent ? calEvents[i].getAllDayEndDate() : calEvents[i].getEndTime(),
      calEvents[i].getLocation(),
      calEvents[i].getCreators(),
      calEvents[i].getDateCreated(),
      getEventGuestList(guestList),
      guestList.length
    ])
  }

  ssSheetData.getRange(2, 1, calData.length, 10).setValues(calData)
  
}

function getEventGuestList(eventGuestList) {
  var guestList = ""

  for (var i = 0; i < eventGuestList.length; i++) {
    guestList == "" ? guestList = eventGuestList[i].getEmail() : guestList += ";" + eventGuestList[i].getEmail();
  }

  return guestList;

}
