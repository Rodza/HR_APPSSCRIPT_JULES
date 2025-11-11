/**
 * @fileoverview All functions for report generation.
 */

/**
 * Generates a weekly payroll summary report.
 * @param {Date} weekEnding The week ending date.
 * @return {object} Report data or error.
 */
function generateWeeklyReport(weekEnding) {
  try {
    var sheet = getSheet('MASTERSALARY');
    if (!sheet) {
      return { success: false, message: 'Salary sheet not found' };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    if (!headers) {
      return { success: false, message: 'No headers found' };
    }
    
    var weekEndingCol = headers.indexOf('WEEKENDING');
    var employeeCol = headers.indexOf('EMPLOYEE NAME');
    var grossCol = headers.indexOf('GROSSSALARY');
    var netCol = headers.indexOf('NETTSALARY');
    var paidCol = headers.indexOf('PaidToAccount');
    
    var weekData = [];
    var totals = {
      employees: 0,
      grossTotal: 0,
      netTotal: 0,
      paidTotal: 0
    };
    
    for (var i = 0; i < data.length; i++) {
      if (formatDate(data[i][weekEndingCol]) === formatDate(weekEnding)) {
        weekData.push({
          employee: data[i][employeeCol],
          gross: data[i][grossCol],
          net: data[i][netCol],
          paid: data[i][paidCol]
        });
        totals.employees++;
        totals.grossTotal += parseFloat(data[i][grossCol]) || 0;
        totals.netTotal += parseFloat(data[i][netCol]) || 0;
        totals.paidTotal += parseFloat(data[i][paidCol]) || 0;
      }
    }
    
    return {
      success: true,
      weekEnding: weekEnding,
      data: weekData,
      totals: totals
    };
  } catch (e) {
    Logger.log('ERROR in generateWeeklyReport: ' + e.message);
    return { success: false, message: 'Error generating report: ' + e.message };
  }
}

/**
 * Generates a monthly payroll summary report.
 * @param {number} month The month (1-12).
 * @param {number} year The year.
 * @return {object} Report data or error.
 */
function generateMonthlyReport(month, year) {
  try {
    var sheet = getSheet('MASTERSALARY');
    if (!sheet) {
      return { success: false, message: 'Salary sheet not found' };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    if (!headers) {
      return { success: false, message: 'No headers found' };
    }
    
    var weekEndingCol = headers.indexOf('WEEKENDING');
    var employeeCol = headers.indexOf('EMPLOYEE NAME');
    var grossCol = headers.indexOf('GROSSSALARY');
    var netCol = headers.indexOf('NETTSALARY');
    var paidCol = headers.indexOf('PaidToAccount');
    
    var monthData = [];
    var totals = {
      payslips: 0,
      employees: {},
      grossTotal: 0,
      netTotal: 0,
      paidTotal: 0
    };
    
    for (var i = 0; i < data.length; i++) {
      var date = new Date(data[i][weekEndingCol]);
      if (date.getMonth() + 1 === month && date.getFullYear() === year) {
        var employeeName = data[i][employeeCol];
        monthData.push({
          employee: employeeName,
          weekEnding: date,
          gross: data[i][grossCol],
          net: data[i][netCol],
          paid: data[i][paidCol]
        });
        totals.payslips++;
        totals.employees[employeeName] = true;
        totals.grossTotal += parseFloat(data[i][grossCol]) || 0;
        totals.netTotal += parseFloat(data[i][netCol]) || 0;
        totals.paidTotal += parseFloat(data[i][paidCol]) || 0;
      }
    }
    
    totals.uniqueEmployees = Object.keys(totals.employees).length;
    delete totals.employees; // Remove the tracking object
    
    return {
      success: true,
      month: month,
      year: year,
      data: monthData,
      totals: totals
    };
  } catch (e) {
    Logger.log('ERROR in generateMonthlyReport: ' + e.message);
    return { success: false, message: 'Error generating report: ' + e.message };
  }
}

/**
 * Generates an employee history report.
 * @param {string} employeeId The employee ID or name.
 * @return {object} Report data or error.
 */
function generateEmployeeReport(employeeId) {
  try {
    var sheet = getSheet('MASTERSALARY');
    if (!sheet) {
      return { success: false, message: 'Salary sheet not found' };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    if (!headers) {
      return { success: false, message: 'No headers found' };
    }
    
    var employeeCol = headers.indexOf('EMPLOYEE NAME');
    var weekEndingCol = headers.indexOf('WEEKENDING');
    var grossCol = headers.indexOf('GROSSSALARY');
    var netCol = headers.indexOf('NETTSALARY');
    var paidCol = headers.indexOf('PaidToAccount');
    var hoursCol = headers.indexOf('HOURS');
    var overtimeCol = headers.indexOf('OVERTIMEHOURS');
    
    var employeeData = [];
    var totals = {
      payslips: 0,
      totalHours: 0,
      totalOvertimeHours: 0,
      grossTotal: 0,
      netTotal: 0,
      paidTotal: 0
    };
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][employeeCol] === employeeId) {
        employeeData.push({
          weekEnding: data[i][weekEndingCol],
          hours: data[i][hoursCol],
          overtime: data[i][overtimeCol],
          gross: data[i][grossCol],
          net: data[i][netCol],
          paid: data[i][paidCol]
        });
        totals.payslips++;
        totals.totalHours += parseFloat(data[i][hoursCol]) || 0;
        totals.totalOvertimeHours += parseFloat(data[i][overtimeCol]) || 0;
        totals.grossTotal += parseFloat(data[i][grossCol]) || 0;
        totals.netTotal += parseFloat(data[i][netCol]) || 0;
        totals.paidTotal += parseFloat(data[i][paidCol]) || 0;
      }
    }
    
    if (totals.payslips > 0) {
      totals.averageGross = totals.grossTotal / totals.payslips;
      totals.averageNet = totals.netTotal / totals.payslips;
      totals.averagePaid = totals.paidTotal / totals.payslips;
    }
    
    return {
      success: true,
      employeeId: employeeId,
      data: employeeData,
      totals: totals
    };
  } catch (e) {
    Logger.log('ERROR in generateEmployeeReport: ' + e.message);
    return { success: false, message: 'Error generating report: ' + e.message };
  }
}

/**
 * Generates a loan balances report showing all outstanding loans.
 * @return {object} Report data or error.
 */
function generateLoanReport() {
  try {
    var sheet = getSheet('EmployeeLoans');
    if (!sheet) {
      return { success: false, message: 'Loans sheet not found' };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    if (!headers) {
      return { success: false, message: 'No headers found' };
    }
    
    var employeeIdCol = headers.indexOf('Employee ID');
    var balanceAfterCol = headers.indexOf('BalanceAfter');
    
    // Get unique employees and their latest balances
    var employeeBalances = {};
    for (var i = 0; i < data.length; i++) {
      var empId = data[i][employeeIdCol];
      employeeBalances[empId] = parseFloat(data[i][balanceAfterCol]) || 0;
    }
    
    // Build report data
    var loanData = [];
    var totalOutstanding = 0;
    
    for (var empId in employeeBalances) {
      var balance = employeeBalances[empId];
      if (balance > 0) {
        // Get employee name from employee sheet
        var employee = getEmployeeById(empId);
        loanData.push({
          employeeId: empId,
          employeeName: employee ? employee.REFNAME : 'Unknown',
          balance: balance
        });
        totalOutstanding += balance;
      }
    }
    
    // Sort by balance descending
    loanData.sort(function(a, b) {
      return b.balance - a.balance;
    });
    
    return {
      success: true,
      asOfDate: new Date(),
      data: loanData,
      totals: {
        employeesWithLoans: loanData.length,
        totalOutstanding: totalOutstanding
      }
    };
  } catch (e) {
    Logger.log('ERROR in generateLoanReport: ' + e.message);
    return { success: false, message: 'Error generating report: ' + e.message };
  }
}

/**
 * Exports report data to CSV format.
 * @param {object} reportData The report data to export.
 * @return {string} CSV-formatted string.
 */
function exportReportToCSV(reportData) {
  try {
    if (!reportData.data || reportData.data.length === 0) {
      return '';
    }
    
    var headers = Object.keys(reportData.data[0]);
    var csv = headers.join(',') + '\n';
    
    for (var i = 0; i < reportData.data.length; i++) {
      var row = [];
      for (var j = 0; j < headers.length; j++) {
        var value = reportData.data[i][headers[j]];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        row.push(value || '');
      }
      csv += row.join(',') + '\n';
    }
    
    return csv;
  } catch (e) {
    Logger.log('ERROR in exportReportToCSV: ' + e.message);
    return '';
  }
}

/**
 * Saves report data to Google Drive.
 * @param {object} reportData The report data.
 * @param {string} reportName The name for the report file.
 * @return {object} Result with file URL or error.
 */
function saveReportToDrive(reportData, reportName) {
  try {
    var csv = exportReportToCSV(reportData);
    if (!csv) {
      return { success: false, message: 'No data to export' };
    }
    
    var fileName = reportName + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.csv';
    var file = DriveApp.createFile(fileName, csv, MimeType.CSV);
    
    return {
      success: true,
      message: 'Report saved to Google Drive',
      fileUrl: file.getUrl(),
      fileName: fileName
    };
  } catch (e) {
    Logger.log('ERROR in saveReportToDrive: ' + e.message);
    return { success: false, message: 'Error saving report: ' + e.message };
  }
}
