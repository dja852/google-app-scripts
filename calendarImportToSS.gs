const SHEET_NAMES = {
  SS_CONFIG: "Config",
  CAL_DATA_TEMPLATE: "_TEMPLATE_Calendar Data",
  CAL_DATA: "Calendar Data_",
  CAL_IDS: "Calendar IDs",
  INT_CONFIG: "IntConfig",
  RUN_LOG: "Run Log",
}

const RUN_MODES = {
  FULL: 1,
  TODAY: 2,
  NEW: 3,
}

const SS_CONFIG = {
  CAL_ID: 1,
  START_DATE: 2,
  END_DATE: 3,
}

const INT_CONFIG = {
  LAST_UPDATED: 1,
}

const SS = SpreadsheetApp.getActiveSpreadsheet();
const UI = SpreadsheetApp.getUi();

function onOpen() {
  UI.createMenu("Calendar Import")
    .addItem("Get Calendar IDs", "getCalendarIds")
    .addSeparator()
    .addItem("Get Calendar Data (Full)", "menuGetCalendarFull")
    .addItem("Get Calendar Data (Today Onwards)", "menuGetCalendarToday")
    .addItem("Get Calendar Data (New/Changed Only)", "menuGetCalendarNew")
    .addSeparator()
    .addItem("Reset Last Updated Date/Time", "resetLastUpdatedDateTime")
    .addToUi();

}

function menuGetCalendarFull() {
  getCalendarEventsById(RUN_MODES.FULL);
}

function menuGetCalendarToday() {
  getCalendarEventsById(RUN_MODES.TODAY);
}

function menuGetCalendarNew() {
  getCalendarEventsById(RUN_MODES.NEW);
}

function getCalendarIds() {
  const ssSheetIDs = SS.getSheetByName(SHEET_NAMES.CAL_IDS);
  const cals = CalendarApp.getAllCalendars();

  let calList = [];

  for (const i in cals) {
    calList.push([
      cals[i].getName(),
      cals[i].getId(),
      cals[i].getTimeZone(),
    ])
  }

  ssSheetIDs.getRange(2, 1, calList.length, 3).setValues(calList);

  ssSheetIDs.activate();

}

function getCalendarEventsById(runMode) {
  const currentDateTime = Date.now();

  const ssSheetConfig = SS.getSheetByName(SHEET_NAMES.SS_CONFIG);
  const ssSheetDataTemplate = SS.getSheetByName(SHEET_NAMES.CAL_DATA_TEMPLATE);
  const ssSheetData = ssSheetDataTemplate.copyTo(SS);

  const cal = CalendarApp.getCalendarById(ssSheetConfig.getRange(SS_CONFIG.CAL_ID, 2).getValue());

  const startTime = runMode == RUN_MODES.TODAY ? currentDateTime : ssSheetConfig.getRange(SS_CONFIG.START_DATE, 2).getValue();
  const endTime = ssSheetConfig.getRange(SS_CONFIG.END_DATE, 2).getValue();

  const calEvents = cal.getEvents(new Date(startTime), new Date(endTime));

  const dataLastUpdated = getLastUpdatedDateTime();
  
  let calData = [];
  let calDataRange = [];

  for (const i in calEvents) {

    const isAllDayEvent = calEvents[i].isAllDayEvent();
    const guestList = calEvents[i].getGuestList(true);

    const eventCreated = calEvents[i].getDateCreated();
    const eventUpdated = calEvents[i].getLastUpdated();

    calDataRange = [
      calEvents[i].getId(),
      calEvents[i].getTitle(),
      calEvents[i].getDescription(),
      isAllDayEvent,
      isAllDayEvent ? calEvents[i].getAllDayStartDate() : calEvents[i].getStartTime(),
      isAllDayEvent ? calEvents[i].getAllDayEndDate() : calEvents[i].getEndTime(),
      calEvents[i].getLocation(),
      calEvents[i].getCreators(),
      eventCreated,
      eventUpdated,
      getEventGuestList(guestList),
      guestList.length,
      eventCreated > dataLastUpdated || eventUpdated > dataLastUpdated ? "X" : "",
      eventCreated > dataLastUpdated ? "NEW" : eventUpdated > dataLastUpdated ? "CHANGED" : "",
    ];

    if (runMode == RUN_MODES.NEW) {
      if (eventCreated > dataLastUpdated || eventUpdated > dataLastUpdated) {
        calData.push(calDataRange);
      }
    } else {
      calData.push(calDataRange);
    }
  }

  try {
    ssSheetData.getRange(2, 1, calData.length, calDataRange.length).setValues(calData);
  } catch(error) {
    UI.alert(error, UI.ButtonSet.OK);
    return;
    
  }

  ssSheetData.setName(SHEET_NAMES.CAL_DATA + currentDateTime);
  ssSheetData.showSheet();
  ssSheetData.activate();
  updateRunLog(currentDateTime, ssSheetData.getSheetId());
  setLastUpdatedDateTime(currentDateTime);
  
}

function updateRunLog(currentDateTime, sheetID) {
  const ssRunLog = SS.getSheetByName(SHEET_NAMES.RUN_LOG);

  ssRunLog.appendRow(
    [
      dateToString(currentDateTime),
      Session.getActiveUser(),
      '=HYPERLINK("#gid=' + sheetID + '","' + (SHEET_NAMES.CAL_DATA + currentDateTime) + '")',
    ]
  );
}

function getEventGuestList(eventGuestList) {
  let guestList = "";

  for (var i in eventGuestList) {
    guestList == "" ? guestList = eventGuestList[i].getEmail() : guestList += ";" + eventGuestList[i].getEmail();
  }

  return guestList;

}

function getLastUpdatedDateTime() {
  const ssSheetIntConfig = SS.getSheetByName(SHEET_NAMES.INT_CONFIG);

  return ssSheetIntConfig.getRange(INT_CONFIG.LAST_UPDATED, 2).getValue() == "" ? 0 : ssSheetIntConfig.getRange(INT_CONFIG.LAST_UPDATED, 2).getValue(); 

}

function setLastUpdatedDateTime(currentDateTime) {
  const ssSheetIntConfig = SS.getSheetByName(SHEET_NAMES.INT_CONFIG);

  ssSheetIntConfig.getRange(INT_CONFIG.LAST_UPDATED, 2).setValue(currentDateTime);

}

function resetLastUpdatedDateTime() {
  const promptResponse = UI.alert("Data last updated: " + dateToString(getLastUpdatedDateTime()) + ". Reset?",UI.ButtonSet.YES_NO);

  if (promptResponse == UI.Button.YES) {
    setLastUpdatedDateTime("");
  }

}

function dateToString(date) {
  try {
    const stringDate = new Date(date).toISOString();
    return stringDate;
  } catch(error) {
    return date;
  }
  
}

function getRelativeDate(daysOffset, hour) {
  const date = new Date();
  date.setDate(date.getDate() + daysOffset);
  date.setHours(hour);
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}
