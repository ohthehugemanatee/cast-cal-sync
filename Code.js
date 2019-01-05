var calendarId = 'vertesi.com_sa3qs1pe5kjjg7k29a4nvi4f6c@group.calendar.google.com';
var allowedSheets = [
  '2019'
];
var sheet = SpreadsheetApp.getActiveSheet();


/**
* Add menu items.
*/
function onOpen() {
  // Add a menu item for syncing EVERYTHING to the calendar.
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom')
  .addItem('Wipe and re-sync calendar', 'deleteAndPushAll')
  .addItem('Re-sync calendar', 'pushAll')
  .addToUi();
}

/**
 * Delete all events on the Calendar and re-sync from the spreadsheet.
 * Trigger: menu item.
 */
function deleteAndPushAll() {
  console.log('Menu link clicked: Delete and push all')
  if (validateAction() == true) {
    deleteAllEvents();
    pushAllEventsToCalendar();
  }
  else {
    console.log('Sheet is not in the allowed list, skipping.')
  }
}

/**
 * Re-push all events from the spreadsheet to the Calendar.
 * Trigger: menu item.
 */
function pushAll() {
  console.log('Menu link clicked: Push all')
  if (validateAction() == true) {
    pushAllEventsToCalendar();
  } else {
    console.log('Sheet is not in the allowed list, skipping.')
  }
}

/**
* On edit, create/update/delete a calendar item.
* Trigger: user edits a row.
*/
function runOnEdit(e) {
  console.log('runOnEdit trigger activated')
  if (validateAction() == false) {
    console.log('Sheet is not in the allowed list, skipping.')
    return;
  }
  var spreadsheet = e.source;
  var sheet = spreadsheet.getActiveSheet();
  var changedRange = e.range;
  var rowNumber = changedRange.getRow()
  // Ignore changes to header row.
  if (rowNumber == 1) {
    return;
  }
  var numRows = changedRange.getNumRows();
  var numCols = sheet.getLastColumn();
  var dataRange = sheet.getRange(rowNumber, 1, numRows, numCols);
  processRange(dataRange);
  console.log('Finished runOnEdit processing')
}

/**
 * Deletes all events on the calendar.
 */
function deleteAllEvents() {
  var cal = CalendarApp.getCalendarById(calendarId);
  var sheet = SpreadsheetApp.getActiveSheet();
  var thisYear = sheet.getName();
  var from = new Date(thisYear + '-01-01T00:00:00');
  var to = new Date(thisYear + '-12-31T23:59:59');
  var events = cal.getEvents(from, to);
  for (var i = 0; i < events.length; i++) {
    events[i].deleteEvent();
  }
  Logger.log("Deleted " + events.length + " events.");
  // Empty out existing calendar Ids.
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var eventIdCol = parseInt(getIndexByName('Calendar ID')) + 1;
  var range = sheet.getRange(2, eventIdCol, lastRow);
  range.clear();
}

/**
* Push all events on the current sheet.
*/
function pushAllEventsToCalendar() {
  console.log('Pushing all events to calendar');
  var dataRange = sheet.getDataRange();
  // Process the range.
  processRange(dataRange);
  console.log('Finished pushing all events');
}

/**
 * Add/Update calendar items from a data range.
 */
function processRange(dataRange) {
  var data = dataRange.getDisplayValues();
  var rowNumber = dataRange.getRow();
  for (i in data) {
    console.log("processing row ", rowNumber)
    var toProcess = data[i];
    // Validate that we have at least date, title requirements, and status values.
    var date = toProcess[getIndexByName('Date')];
    var city = toProcess[getIndexByName('City')];
    var type = toProcess[getIndexByName('Rehearsal or Performance')];
    var status = toProcess[getIndexByName('Status')];
    if (date && date != null && city && city !== '' && type && type != '' && status && status != '') {
      // Create/Update the event.
      var createdEvent = pushSingleEventToCalendar(sheet, toProcess);
      // Add event ID to the spreadsheet.
      postEventPush(createdEvent, rowNumber);
    }
    // Increment rowNumber.
    rowNumber++
  }
}

/**
 * Post event Push cleanup function.
 */
function postEventPush(createdEvent, rowId) {
  if (createdEvent) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var eventIdCol = parseInt(getIndexByName('Calendar ID')) + 1;
    // If the pushed event is cancelled, clear the value from the event Id field.
    if (createdEvent == 'Cancelled') {
      sheet.getRange(rowId, eventIdCol, 1, 1).clear();
    }
    else { // Otherwise, set the value in the event Id field.
      sheet.getRange(rowId, eventIdCol, 1, 1).setValue(createdEvent.getId());
    }
    // Make sure the cell is updated right away in case the script is interrupted
    SpreadsheetApp.flush();
  }
}

/**
* Push a single event to the calendar.
*/
function pushSingleEventToCalendar(sheet, row) {
  // The calendar to be modified.
  var cal = CalendarApp.getCalendarById(calendarId);
  // Get an Event to work with. Either load an existing one, or create a stub.
  var eventId = row[getIndexByName('Calendar ID')];
  try {
    // Try loading an existing event.
    console.log("attempting to load cal ID ", eventId);
    var calEvent = cal.getEventById(eventId);
    console.log(calEvent);
    if (calEvent === null) {
      throw new Error("getEventById returned null.");
    }
    console.log('Updating existing event')
    // If the spreadsheet says it's cancelled, delete the calendar event and return.
    if (row[getIndexByName('Status')] == "Cancelled") {
      console.log('Existing event cancelled')
      return 'Cancelled';
    }
  }
  catch (e) {
    console.log(e);
    // If we couldn't load the event, create a stub.
    console.log('Could not load Calendar event. Creating a new one instead.');
    var calEvent = cal.createAllDayEvent('Stub event', new Date());
  }
  // Create/update values for the event.
  setTitle(calEvent, row);
  setDescription(calEvent, row);
  setDate(calEvent, row);
  setColor(calEvent, row);
  return calEvent;
}
