/**
* Add an item to trigger bulk sync to the menu.
*/
function onOpen() {
  // Add a menu item for syncing EVERYTHING to the calendar.
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom')
  .addItem('Wipe and re-sync calendar', 'deleteAndPushAll')
  .addToUi();
}

/**
 * Delete all events on the Calendar and re-sync from the spreadsheet.
 */
function deleteAndPushAll() {
  deleteAllEvents();
  pushAllEventsToCalendar();
}

/**
 * Deletes all events on the calendar.
 */
function deleteAllEvents() {
    var cal = CalendarApp.getCalendarById("vertesi.com_sa3qs1pe5kjjg7k29a4nvi4f6c@group.calendar.google.com");
  var from = new Date('01/01/2017');
  var to = new Date('01/01/2050');
  var events = cal.getEvents(from, to);
  for (var i = 0; i < events.length; i++) {
    events[i].deleteEvent();
  }
  Logger.log("Deleted " + events.length + " events.");
}


/**
 * Get the column index by header row value.
 * NB: this is an array index; it starts from 0!
 */
function getIndexByName(name){
  var headers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2018').getDataRange().getValues().shift();
  var colindex = headers.indexOf(name);
  return colindex;
}

/**
* Add an automatic single item sync on edit.
*/
function runOnEdit(e) {
  var spreadsheet = e.source;
  var sheet = spreadsheet.getSheetByName('2018');
  var changedRange = e.range;
  var rowNumber = changedRange.getRow()
  var numRows = changedRange.getNumRows();
  var numCols = sheet.getLastColumn();
  var dataRange = sheet.getRange(rowNumber, 1, numRows, numCols);
  var data = dataRange.getValues();
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
* Push all events on the current sheet.
*/
function pushAllEventsToCalendar() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var dataRange = sheet.getRange("2018!A2:"+lastRow+lastColumn);
  var data = dataRange.getValues();
  for (i in data) {
    var rowId = 2 + parseInt(i);
    console.log("processing row ", rowId)
    var createdEvent = pushSingleEventToCalendar(sheet, data[i]);
    console.log({message: 'Created event', initialData: createdEvent})
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
  var cal = CalendarApp.getCalendarById("vertesi.com_sa3qs1pe5kjjg7k29a4nvi4f6c@group.calendar.google.com");
  
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
    console.log('Updating existing event')
    // Load the existing event.
    try {
      var calEvent = cal.getEventById(eventId);
    }
    catch (e) {
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