function MoveRow() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('26:26').activate();
  spreadsheet.getActiveSheet().moveRows(spreadsheet.getRange('26:26'), 141);
};