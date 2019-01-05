/**
 * @file
 * Functions for setting field values on events.
 */

/**
 * Set the title value.
 *
 * @param {CalendarEvent} event
 * @param {Array} row
 */
function setTitle(event, row) {
  var title = row[getIndexByName('Rehearsal or Performance')] + ": " + row[getIndexByName('City')];
  event.setTitle(title);
}

/**
 * Set the description value.
 *
 * @param {CalendarEvent} event
 * @param {Array} row
 */
function setDescription(event, row) {
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
  event.setDescription(desc);
}

/**
 * Set the date value.
 *
 * @param {CalendarEvent} event
 * @param {Array} row
 */
function setDate(event, row) {
  var date = row[getIndexByName('Date')];
  var from = row[getIndexByName('Time: From')];
  var until = row[getIndexByName('Time: To')];
  // Date could be from/until, or all day.
  // If we have from/until, use those.
  if (from && from != '' && until && until !== '') {
    event.setTime(new Date(date + "T" + from + ":00"), new Date(date + "T" + until + ":00"));
  }
  // Otherwise treat as an all day event.
  else {
    event.setAllDayDate(new Date(date));
  }
}

/**
 * Set the location value.
 *
 * @param {CalendarEvent} event
 * @param {Array} row
 */
function setLocation(event, row) {
  var loc = row[getIndexByName('City')];
  event.setLocation(loc);
}

/**
 * Set the color.
 *
 * @param {CalendarEvent} event
 * @param {Array} row
 */
function setColor(event, row) {
  // Color indicates event status.
  var status = row[getIndexByName('Status')];
  if (status != "Confirmed") {
    event.setColor("8");
  }
  else {
    event.setColor("0");
  }
}