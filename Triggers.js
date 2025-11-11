/**
 * @fileoverview All trigger-related functions.
 */

/**
 * Installs the onChange trigger for the spreadsheet.
 */
function installOnChangeTrigger() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onChange')
    .forSpreadsheet(spreadsheet)
    .onChange()
    .create();
}

/**
 * The main onChange trigger handler.
 * @param {object} e The event object.
 */
function onChange(e) {
  // Check if the change was in the MASTERSALARY sheet
  var sheetName = e.source.getActiveSheet().getName();
  if (sheetName === 'MASTERSALARY') {
    // Get the record number from the edited row
    var editedRow = e.source.getActiveRange().getRow();
    var recordNumber = e.source.getActiveSheet().getRange(editedRow, 1).getValue();

    // Sync the loan for the payslip
    syncLoanForPayslip(recordNumber);
  }
}

/**
 * Uninstalls all triggers for the project.
 */
function uninstallTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}
