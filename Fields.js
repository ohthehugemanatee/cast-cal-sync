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

/**
 * Set the guest list.
 *
 * @param {CalendarEvent} event
 * @param {Array} row
 */
function setGuests(event, row) {
  // Get the array of performer names / mail mappings.
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var nameMapRange = spreadsheet.getRangeByName('PerformerMails');
  var nameMapArray = nameMapRange.getDisplayValues().filter(function (x) { return x[0] !== '' });
  var nameMap = ObjApp.rangeToObjects(nameMapArray);
  var availableNames = nameMap.map(function(mapping) {
    return mapping.performer
  });
  // Get the intersection of the names for which we have emails, and the names on the row.
  var invitees = ArrayLib.filterByText(row, -1, availableNames);
  // If we don't have emails for any invitees, return empty handed.
  if (invitees === []) {
    console.log('No one to invite');
    return;
  }

  // Loop through the invitees
  for (var iindex in invitees) {
    for(var mindex in nameMap) {
      var guest = nameMap[mindex];
      if (guest.performer == invitees[iindex]) {
        console.log("Adding guest: " + guest.performer);
        event.addGuest(guest.googleEmail);
      }
    }
  }
}
