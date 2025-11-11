/**
 * @fileoverview Main entry points for the web app.
 */

/**
 * Serves the main dashboard HTML page.
 * @return {HtmlOutput} The HTML output for the web app.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard.html')
      .setTitle('HR System Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Includes the content of an HTML file.
 * Used to load different pages into the main dashboard.
 * @param {string} filename The name of the HTML file to include.
 * @return {string} The HTML content of the file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Includes the content of an HTML file from the components folder.
 * Used to load different pages into the main dashboard.
 * @param {string} filename The name of the component file to include.
 * @return {string} The HTML content of the file.
 */
function getComponent(filename) {
  return HtmlService.createHtmlOutputFromFile('components/' + filename).getContent();
}

// --- Employee Management ---
function addEmployee(data) { return Employees.addEmployee(data); }
function updateEmployee(id, data) { return Employees.updateEmployee(id, data); }
function getEmployeeById(id) { return Employees.getEmployeeById(id); }
function listEmployees(filters) { return Employees.listEmployees(filters); }
function terminateEmployee(id, terminationDate) { return Employees.terminateEmployee(id, terminationDate); }

// --- Leave Management ---
function addLeave(data) { return Leave.addLeave(data); }
function listLeave(filters) { return Leave.listLeave(filters); }

// --- Loan Management ---
function addLoanTransaction(data) { return Loans.addLoanTransaction(data); }
function getLoanHistory(employeeId) { return Loans.getLoanHistory(employeeId); }
function getCurrentLoanBalance(employeeId) { return Loans.getCurrentLoanBalance(employeeId); }

// --- Timesheet Management ---
function importTimesheetData(data) { return Timesheets.importTimesheetData(data); }
function approveTimesheet(id) { return Timesheets.approveTimesheet(id); }
function rejectTimesheet(id) { return Timesheets.rejectTimesheet(id); }
function listPendingTimesheets(filters) { return Timesheets.listPendingTimesheets(filters); }

// --- Payroll Management ---
function createPayslip(data) { return Payroll.createPayslip(data); }
function listPayslips(filters) { return Payroll.listPayslips(filters); }
function calculatePayslip(data) { return Payroll.calculatePayslip(data); }
function generatePayslipPDF(recordNumber) { return Payroll.generatePayslipPDF(recordNumber); }

// --- Reporting ---
function generateOutstandingLoansReport(asOfDate) { return Reports.generateOutstandingLoansReport(asOfDate); }
function generateIndividualStatementReport(employeeName, startDate, endDate) { return Reports.generateIndividualStatementReport(employeeName, startDate, endDate); }
function generateWeeklyPayrollSummaryReport(weekEnding) { return Reports.generateWeeklyPayrollSummaryReport(weekEnding); }
