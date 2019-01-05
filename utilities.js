/**
 * Get the column index by header row value.
 * NB: this is an array index; it starts from 0!
 * @param {String} name
 */
function getIndexByName(name){
    var headers = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues().shift();
    var colindex = headers.indexOf(name);
    return colindex;
  }

  /**
   * Check that we should be taking action on this sheet.
   */
  function validateAction() {
    // Returns TRUE if this sheet name is in the allowed sheets list.
    var sheet = SpreadsheetApp.getActiveSheet();
    return allowedSheets.indexOf(sheet.getName()) != -1;
  }