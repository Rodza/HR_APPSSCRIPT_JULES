/**
 * @fileoverview All functions for timesheet import and approval workflow.
 */

/**
 * Imports timesheet data from CSV (from HTML analyzer output).
 * @param {string} csvText The CSV data as text.
 * @return {object} Result with success status and message.
 */
function importTimesheetCSV(csvText) {
  try {
    var parsed = parseTimesheetCSV(csvText);
    if (!parsed.success) {
      return parsed;
    }
    
    var result = stagePendingTimesheets(parsed.data);
    return result;
  } catch (e) {
    Logger.log('ERROR in importTimesheetCSV: ' + e.message);
    return { success: false, message: 'Error importing timesheet: ' + e.message };
  }
}

/**
 * Parses CSV text into structured timesheet data.
 * @param {string} csvText The CSV data as text.
 * @return {object} Parsed data or error.
 */
function parseTimesheetCSV(csvText) {
  try {
    var lines = csvText.trim().split('\n');
    if (lines.length < 2) {
      return { success: false, message: 'CSV file is empty or invalid' };
    }
    
    var headers = lines[0].split(',').map(function(h) {
      return h.trim();
    });
    var data = [];
    
    for (var i = 1; i < lines.length; i++) {
      var values = lines[i].split(',').map(function(v) {
        return v.trim();
      });
      var record = {};
      headers.forEach(function(header, idx) {
        record[header] = values[idx];
      });
      data.push(record);
    }
    
    return { success: true, data: data };
  } catch (e) {
    Logger.log('ERROR in parseTimesheetCSV: ' + e.message);
    return { success: false, message: 'Error parsing CSV: ' + e.message };
  }
}

/**
 * Stages timesheet data in the PendingTimesheets sheet for approval.
 * @param {Array<object>} data The timesheet records to stage.
 * @return {object} Result with success status and message.
 */
function stagePendingTimesheets(data) {
  try {
    var sheet = getSheet('PendingTimesheets');
    if (!sheet) {
      return { success: false, message: 'PendingTimesheets sheet not found. Please check sheet configuration.' };
    }
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    var addedCount = 0;
    for (var i = 0; i < data.length; i++) {
      var record = data[i];
      var validation = validateTimesheet(record);
      if (!validation.isValid) {
        Logger.log('WARNING: Skipping invalid record - ' + validation.errors.join(', '));
        continue;
      }
      
      var newRow = headers.map(function(header) {
        switch(header) {
          case 'Status':
            return 'Pending';
          case 'ImportDate':
            return new Date();
          case 'RecordID':
            return generateUUID();
          default:
            return record[header] || null;
        }
      });
      
      sheet.appendRow(newRow);
      addedCount++;
    }
    
    return { 
      success: true, 
      message: addedCount + ' timesheet record(s) staged for approval' 
    };
  } catch (e) {
    Logger.log('ERROR in stagePendingTimesheets: ' + e.message);
    return { success: false, message: 'Error staging timesheets: ' + e.message };
  }
}

/**
 * Approves a pending timesheet and creates a payslip.
 * @param {string} recordId The RecordID of the pending timesheet.
 * @return {object} Result with success status and message.
 */
function approveTimesheet(recordId) {
  try {
    var sheet = getSheet('PendingTimesheets');
    if (!sheet) {
      return { success: false, message: 'PendingTimesheets sheet not found' };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    var recordIdCol = headers.indexOf('RecordID');
    var statusCol = headers.indexOf('Status');
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][recordIdCol] === recordId) {
        // Build payslip data from timesheet record
        var payslipData = {};
        headers.forEach(function(header, idx) {
          payslipData[header] = data[i][idx];
        });
        
        // Create the payslip
        var result = createPayslip(payslipData);
        if (!result.success) {
          return result;
        }
        
        // Update status to Approved
        sheet.getRange(i + 2, statusCol + 1).setValue('Approved');
        
        return { 
          success: true, 
          message: 'Timesheet approved and payslip created',
          recordNumber: result.recordNumber 
        };
      }
    }
    
    return { success: false, message: 'Timesheet record not found' };
  } catch (e) {
    Logger.log('ERROR in approveTimesheet: ' + e.message);
    return { success: false, message: 'Error approving timesheet: ' + e.message };
  }
}

/**
 * Rejects a pending timesheet.
 * @param {string} recordId The RecordID of the pending timesheet.
 * @return {object} Result with success status and message.
 */
function rejectTimesheet(recordId) {
  try {
    var sheet = getSheet('PendingTimesheets');
    if (!sheet) {
      return { success: false, message: 'PendingTimesheets sheet not found' };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    var recordIdCol = headers.indexOf('RecordID');
    var statusCol = headers.indexOf('Status');
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][recordIdCol] === recordId) {
        sheet.getRange(i + 2, statusCol + 1).setValue('Rejected');
        return { success: true, message: 'Timesheet rejected' };
      }
    }
    
    return { success: false, message: 'Timesheet record not found' };
  } catch (e) {
    Logger.log('ERROR in rejectTimesheet: ' + e.message);
    return { success: false, message: 'Error rejecting timesheet: ' + e.message };
  }
}

/**
 * Gets approved timesheets for a specific week ending date.
 * @param {Date} weekEnding The week ending date.
 * @return {Array<object>} List of approved timesheet records.
 */
function getApprovedTimesheets(weekEnding) {
  var sheet = getSheet('PendingTimesheets');
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  if (!headers) return [];
  
  var statusCol = headers.indexOf('Status');
  var weekEndingCol = headers.indexOf('WEEKENDING');
  
  var approved = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][statusCol] === 'Approved' && 
        formatDate(data[i][weekEndingCol]) === formatDate(weekEnding)) {
      var record = {};
      headers.forEach(function(header, idx) {
        record[header] = data[i][idx];
      });
      approved.push(record);
    }
  }
  
  return approved;
}

/**
 * Clears all pending timesheets (use with caution).
 * @return {object} Result with success status and message.
 */
function clearPendingTimesheets() {
  try {
    var sheet = getSheet('PendingTimesheets');
    if (!sheet) {
      return { success: false, message: 'PendingTimesheets sheet not found' };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }
    
    return { success: true, message: 'All pending timesheets cleared' };
  } catch (e) {
    Logger.log('ERROR in clearPendingTimesheets: ' + e.message);
    return { success: false, message: 'Error clearing timesheets: ' + e.message };
  }
}

/**
 * Validates timesheet data.
 * @param {object} data The timesheet data to validate.
 * @return {object} An object with `isValid` and a list of `errors`.
 */
function validateTimesheet(data) {
  var errors = [];
  
  if (!data['EMPLOYEE NAME']) {
    errors.push('Employee name is required');
  }
  
  if (!data.WEEKENDING) {
    errors.push('Week ending date is required');
  }
  
  var hours = parseFloat(data.HOURS) || 0;
  var overtimeHours = parseFloat(data.OVERTIMEHOURS) || 0;
  
  if (hours < 0 || overtimeHours < 0) {
    errors.push('Hours cannot be negative');
  }
  
  return { 
    isValid: errors.length === 0, 
    errors: errors 
  };
}
