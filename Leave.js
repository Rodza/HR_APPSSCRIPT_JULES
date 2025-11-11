/**
 * @fileoverview All functions for leave tracking.
 */

/**
 * Adds a new leave record.
 * @param {object} data The leave data from the form.
 * @return {object} A success or error message.
 */
function addLeave(data) {
  try {
    const validation = validateLeave(data);
    if (!validation.isValid) {
      return { success: false, errors: validation.errors };
    }

    const sheet = getSheet('leave');
    if (!sheet) {
      return { success: false, message: 'Leave sheet not found.' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const newRow = headers.map(header => {
      switch(header) {
        case 'TIMESTAMP':
          return new Date();
        case 'USER':
          return getCurrentUser();
        case 'TOTALDAYS.LEAVE':
            const startDate = new Date(data['STARTDATE.LEAVE']);
            const returnDate = new Date(data['RETURNDATE.LEAVE']);
            const diffTime = Math.abs(returnDate - startDate);
            return Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1; 
        default:
          return data[header] || null;
      }
    });

    sheet.appendRow(newRow);
    return { success: true, message: 'Leave record added successfully.' };
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);
    return { success: false, message: 'An error occurred while adding the leave record.' };
  }
}

/**
 * Retrieves the leave history for a specific employee.
 * @param {string} employeeName The name of the employee.
 * @return {Array<object>} A list of leave records.
 */
function getLeaveHistory(employeeName) {
  const sheet = getSheet('leave');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const nameCol = headers.indexOf('EMPLOYEE NAME');
  
  const records = data.map(row => {
    const record = {};
    headers.forEach((header, i) => record[header] = row[i]);
    return record;
  });

  return records.filter(rec => rec['EMPLOYEE NAME'] === employeeName);
}

/**
 * Lists all leave records, with optional filters.
 * @param {object} filters The filters to apply.
 * @return {Array<object>} A list of leave records.
 */
function listLeave(filters = {}) {
  const sheet = getSheet('leave');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  const records = data.map(row => {
    const record = {};
    headers.forEach((header, i) => record[header] = row[i]);
    return record;
  });

  return records.filter(rec => {
      let keep = true;
      // Add filtering logic here if needed
      return keep;
  });
}

/**
 * Validates leave data.
 * @param {object} data The leave data to validate.
 * @return {object} An object with `isValid` and a list of `errors`.
 */
function validateLeave(data) {
  const errors = [];
  if (!data['EMPLOYEE NAME']) errors.push('Employee name is required.');
  if (!data['STARTDATE.LEAVE']) errors.push('Start date is required.');
  if (!data['RETURNDATE.LEAVE']) errors.push('Return date is required.');
  if (!data['REASON']) errors.push('Reason is required.');

  if (data['STARTDATE.LEAVE'] && data['RETURNDATE.LEAVE']) {
      if (new Date(data['RETURNDATE.LEAVE']) < new Date(data['STARTDATE.LEAVE'])) {
          errors.push('Return date cannot be before the start date.');
      }
  }
  
  return { isValid: errors.length === 0, errors: errors };
}
