# cast-cal-sync
Sync script between google sheets and gcal for The Cast

This is probably useful as an  example for others. This Apps Script pushes events from a google sheets document into a specific google calendar.

[The Cast](thecastmusic.com) uses this to track rehearsals and performances in a shared calendar. The spreadsheet can contain lots of other information that isn't necessarily pushed into the calendar, and can be used for other operations too. For example, we base our internal pay calculations on the spreadsheet.

* Offers two menu items:
  * Delete and re-sync: deletes all calendar items, and re-creates based on the current sheet (only)
  * Re-sync: create/update/delete calendar items based on the current sheet.
* Targets the calendar ID named in the `calendarId` variable at top
* Always uses the active sheet as a basis for the menu links
* Updates items as they are entered
* Skips if the current sheet isn't named in the `allowedSheets` variable at top
* Multiple musicians can be entered, separated by a `/` character
* Enter musician names and email addresses in the 'PerformerMails' named range, to have them receive calendar invitations for their events.