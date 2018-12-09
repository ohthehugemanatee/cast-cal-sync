var calendarId = 'vertesi.com_sa3qs1pe5kjjg7k29a4nvi4f6c@group.calendar.google.com';
var allowedSheets = [
  '2018',
  '2019'
];

/**
* Add an item to trigger bulk sync to the menu.
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
 */
function deleteAndPushAll() {
  console.log('Menu link clicked: Delete and push all')
  if (validateAction()) {
    deleteAllEvents();
    pushAllEventsToCalendar();
  }
  else {
    console.log('Sheet is not in the allowed list, skipping.')
  }
}

/**
 * Delete all events on the Calendar and re-sync from the spreadsheet.
 */
function pushAll() {
  console.log('Menu link clicked: Push all')
  if (validateAction()) {
    pushAllEventsToCalendar();
  }
  else {
    console.log('Sheet is not in the allowed list, skipping.')
  }
}

/**
* On edit, create/update/delete a calendar item.
*/
function runOnEdit(e) {
  console.log('runOnEdit trigger activated')
  if (validateAction() == false) {
    return;
  }
  else {
    console.log('Sheet is not in the allowed list, skipping.')
    return;
  }
  var spreadsheet = e.source;
  var sheet = spreadsheet.getActiveSheet();
  var changedRange = e.range;
  var rowNumber = changedRange.getRow()
  var numRows = changedRange.getNumRows();
  var numCols = sheet.getLastColumn();
  var dataRange = sheet.getRange(rowNumber, 1, numRows, numCols);
  var data = dataRange.getValues();
  // Ignore changes to header row.
  if (rowNumber == 1) {
    return;
  }
  for (i in data) {
    var toProcess = data[i];
    // Validate that we have at least date, title requirements, and status values.
    var date = toProcess[getIndexByName('Date')];
    var city = toProcess[getIndexByName('City')];
    var type = toProcess[getIndexByName('Rehearsal or Performance')];
    var status = toProcess[getIndexByName('Status')];
    if (date && date != null && city && city !== '' && type && type != '' && status && status != '') {
      var createdEvent = pushSingleEventToCalendar(sheet, toProcess);
      postEventPush(createdEvent, rowNumber);
      rowNumber++
    }
  }
}

/**
 * Deletes all events on the calendar.
 */
function deleteAllEvents() {
  var cal = CalendarApp.getCalendarById(calendarId);
  var from = new Date('01/01/2017');
  var to = new Date('01/01/2050');
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
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  for (i in data) {
    // Skip the header row.
    if (parseInt(i) === 0) {
      break;
    }
    var rowId = 1 + parseInt(i);
    console.log("processing row ", rowId)
    var createdEvent = pushSingleEventToCalendar(sheet, data[i]);
    postEventPush(createdEvent, rowId);
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

  // Create/update values for the event.
  var title = row[getIndexByName('Rehearsal or Performance')] + ": " + row[getIndexByName('City')];
  // Description combines a lot of fields.
  var desc = '';
  var descFields = [
    'Pianist',
    'Soprano 1',
    'Soprano 2',
    'Mezzo / Soprano 3',
    'Tenor',
    'Baritone 2',
    'Bass / Baritone 3',
    'Babysitter',
    'Notes'
  ];
  for (i in descFields) {
    var fieldName = descFields[i];
    var fieldValue = row[getIndexByName(fieldName)];
    desc += fieldName + ": " + fieldValue + "\r\n";
  }
  var date = row[getIndexByName('Date')];
  var loc = row[getIndexByName('City')];
  var status = row[getIndexByName('Status')];

  // Existing events have an ID saved.
  var eventId = row[getIndexByName('Calendar ID')];
  if (eventId) {
    console.log('Id found in sheet. Updating existing event')
    // Load the existing event.
    try {
      var calEvent = cal.getEventById(eventId);
    }
    catch (e) {
      console.log('Could not load Calendar event. Clearing entry from the sheet')
    // If the Id is invalid, return Cancelled to remove the value from the cell.
      return 'Cancelled';
    }
    // If the event is cancelled, delete it from the calendar.
    if (status == "Cancelled") {
      console.log('Existing event cancelled')
      if (calEvent) {
          calEvent.deleteEvent();
      }
      return 'Cancelled';
    }
    else {
      calEvent.setTitle(title);
      calEvent.setDescription(desc);
      calEvent.setAllDayDate(date);
      calEvent.setLocation(loc);
      if (row[getIndexByName('Status')] != "Confirmed") {
        calEvent.setColor("8");
      }
      else {
        calEvent.setColor("0");
      }

    }
  }
  else { // New event.
    console.log('Creating new event')
    if (status != "Cancelled") {
      var options = {
        location:loc,
        description:desc,
      };
      if (typeof date === 'string' || date instanceof String) {
        return;
      }
      var calEvent = cal.createAllDayEvent(title,
                                               date,
                                               options
                                              );
      if (status != "Confirmed") {
        calEvent.setColor("8");
      }
    }
  }
  return calEvent;
}

/**
 * Get the column index by header row value.
 * NB: this is an array index; it starts from 0!
 */
function getIndexByName(name){
  var headers = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues().shift();
  var colindex = headers.indexOf(name);
  return colindex;
}

/**
 * Check that we should be taking action.
 */
function validateAction() {
  // Returns TRUE if this sheet name is in the allowed sheets list.
  var sheet = SpreadsheetApp.getActiveSheet();
  return allowedSheets.indexOf(sheet.getName()) != -1;
}