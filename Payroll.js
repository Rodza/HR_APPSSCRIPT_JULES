/**
 * @fileoverview All functions for payroll processing.
 */

// Template ID for payslip PDF generation
var PAYSLIP_TEMPLATE_ID = '1FTgT6tiWFLDVVQEFM_wq6JAt4-ADQr6yJIU70uWzxLI';

/**
 * Creates a new payslip.
 * @param {object} data The payslip data.
 * @return {object} A success or error message.
 */
function createPayslip(data) {
  try {
    var validation = validatePayslip(data);
    if (!validation.isValid) {
      return { success: false, errors: validation.errors };
    }
    
    var calculatedData = calculatePayslip(data);

    var sheet = getSheet('salary');
    if (!sheet) {
      return { success: false, message: 'Salary sheet not found. Please check sheet configuration.' };
    }
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var lastRecordNumber = sheet.getRange(sheet.getLastRow(), 1).getValue() || 7915;

    var newRow = headers.map(function(header) {
        return calculatedData[header] || data[header] || null;
    });
    
    // Manually set fields not in calculatedData
    var recordNumberIndex = headers.indexOf('RECORDNUMBER');
    newRow[recordNumberIndex] = lastRecordNumber + 1;
    newRow[headers.indexOf('TIMESTAMP')] = new Date();
    newRow[headers.indexOf('USER')] = getCurrentUser();


    sheet.appendRow(newRow);
    return { success: true, message: 'Payslip created successfully.', recordNumber: newRow[recordNumberIndex] };
  } catch (e) {
    Logger.log('ERROR in createPayslip: ' + e.message);
    return { success: false, message: 'An error occurred while creating the payslip.' };
  }
}

/**
 * Lists all payslips, with optional filters.
 * @param {object} filters The filters to apply.
 * @return {Array<object>} A list of payslips.
 */
function listPayslips(filters) {
  filters = filters || {};
  var sheet = getSheet('salary');
  if(!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  if(!headers) return [];
  
  var records = data.map(function(row) {
    var record = {};
    headers.forEach(function(header, i) {
      record[header] = row[i];
    });
    return record;
  });

  return records.filter(function(rec) {
      var keep = true;
      if(filters.weekEnding && formatDate(rec.WEEKENDING) !== formatDate(new Date(filters.weekEnding))) {
          keep = false;
      }
      return keep;
  }).sort(function(a,b) {
    return b.RECORDNUMBER - a.RECORDNUMBER;
  });
}


/**
 * Calculates all the fields for a payslip.
 * @param {object} data The payslip data.
 * @return {object} The calculated payslip data.
 */
function calculatePayslip(data) {
    var employee = getEmployeeByName(data['EMPLOYEE NAME']);
    if (!employee) {
      throw new Error('Employee not found: ' + data['EMPLOYEE NAME']);
    }
    
    var hourlyRate = parseFloat(employee['HOURLY RATE']);
    
    var hours = parseFloat(data.HOURS) || 0;
    var minutes = parseFloat(data.MINUTES) || 0;
    var overtimeHours = parseFloat(data.OVERTIMEHOURS) || 0;
    var overtimeMinutes = parseFloat(data.OVERTIMEMINUTES) || 0;
    var leavePay = parseFloat(data['LEAVE PAY']) || 0;
    var bonusPay = parseFloat(data['BONUS PAY']) || 0;
    var otherIncome = parseFloat(data['OTHERINCOME']) || 0;
    var otherDeductions = parseFloat(data['OTHER DEDUCTIONS']) || 0;
    var loanDeduction = parseFloat(data.LoanDeductionThisWeek) || 0;
    var newLoan = parseFloat(data.NewLoanThisWeek) || 0;

    var standardTime = (hours * hourlyRate) + ((hourlyRate / 60) * minutes);
    var overtime = (overtimeHours * hourlyRate * 1.5) + ((hourlyRate / 60) * overtimeMinutes * 1.5);
    var grossSalary = standardTime + overtime + leavePay + bonusPay + otherIncome;
    var uif = (employee['EMPLOYMENT STATUS'] === "Permanent") ? (grossSalary * 0.01) : 0;
    
    var totalDeductions = uif + otherDeductions + loanDeduction;
    var netSalary = grossSalary - totalDeductions;
    var paidToAccount = netSalary - loanDeduction + newLoan;

    var result = {};
    for (var key in data) {
      result[key] = data[key];
    }
    
    result.HOURLYRATE = hourlyRate;
    result.STANDARDTIME = standardTime;
    result.OVERTIME = overtime;
    result.GROSSSALARY = grossSalary;
    result.UIF = uif;
    result.TOTALDEDUCTIONS = totalDeductions;
    result.NETTSALARY = netSalary;
    result.PaidToAccount = paidToAccount;
    
    return result;
}


/**
 * Gets a single payslip by record number.
 * @param {string} recordNumber The unique record number of the payslip.
 * @return {object} The payslip data, or null if not found.
 */
function getPayslip(recordNumber) {
  try {
    var sheet = getSheet('salary');
    if (!sheet) return null;
    
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    if (!headers) return null;
    
    var recordCol = headers.indexOf('RECORDNUMBER');
    if (recordCol === -1) return null;
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][recordCol] == recordNumber) {
        var payslip = {};
        headers.forEach(function(header, idx) {
          payslip[header] = data[i][idx];
        });
        return payslip;
      }
    }
    
    return null;
  } catch (e) {
    Logger.log('ERROR in getPayslip: ' + e.message);
    return null;
  }
}

/**
 * Generates a PDF for a given payslip.
 * @param {string} recordNumber The unique record number of the payslip.
 * @return {string} The URL of the generated PDF.
 */
function generatePayslipPDF(recordNumber) {
    var payslip = getPayslip(recordNumber);
    if(!payslip) return null;
    
    if (!PAYSLIP_TEMPLATE_ID) {
      Logger.log('ERROR: PAYSLIP_TEMPLATE_ID not set. Cannot generate PDF.');
      return null;
    }

    var sheet = getSheet('salary');
    if (!sheet) {
      Logger.log('ERROR: Salary sheet not found');
      return null;
    }

    var template = DriveApp.getFileById(PAYSLIP_TEMPLATE_ID);
    var newDocName = 'Payslip #' + recordNumber + ' - ' + payslip['EMPLOYEE NAME'];
    var newFile = template.makeCopy(newDocName);
    var doc = DocumentApp.openById(newFile.getId());
    var body = doc.getBody();

    // Replace placeholders
    for(var key in payslip) {
        var value = payslip[key];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        body.replaceText('{{' + key + '}}', value || '');
    }
    doc.saveAndClose();

    var pdf = newFile.getAs('application/pdf');
    var pdfFile = DriveApp.createFile(pdf).setName(newDocName + '.pdf');
    newFile.setTrashed(true); // Delete the temporary Google Doc

    // Store link back in the sheet
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    var recordCol = headers.indexOf('RECORDNUMBER');
    var fileLinkCol = headers.indexOf('FILELINK');
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][recordCol] == recordNumber) {
        sheet.getRange(i + 2, fileLinkCol + 1).setValue(pdfFile.getUrl());
        break;
      }
    }

    return pdfFile.getUrl();
}

/**
 * Validates payslip data.
 * @param {object} data The payslip data to validate.
 * @return {object} An object with `isValid` and a list of `errors`.
 */
function validatePayslip(data) {
  var errors = [];
  
  if (!data['EMPLOYEE NAME']) {
    errors.push('Employee Name is required');
  }
  
  if (!data.WEEKENDING) {
    errors.push('Week Ending date is required');
  }
  
  var hours = parseFloat(data.HOURS) || 0;
  var overtimeHours = parseFloat(data.OVERTIMEHOURS) || 0;
  
  if (hours < 0 || overtimeHours < 0) {
    errors.push('Hours cannot be negative');
  }
  
  if (hours + overtimeHours > 168) {
    errors.push('Total hours exceed maximum possible (168 hours/week)');
  }
  
  return { 
    isValid: errors.length === 0, 
    errors: errors 
  };
}
