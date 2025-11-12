/**
 * @fileoverview All functions for employee management.
 */

const EMPLOYEE_SHEET = 'EMPLOYEE DETAILS';

/**
 * Adds a new employee to the EMPLOYEE DETAILS sheet.
 * @param {object} data The employee data from the form.
 * @return {object} A success or error message.
 */
function addEmployee(data) {
  try {
    const validation = validateEmployee(data);
    if (!validation.isValid) {
      return { success: false, errors: validation.errors };
    }

    const sheet = getSheet(EMPLOYEE_SHEET);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const newRow = headers.map(header => {
      switch(header) {
        case 'ID':
          return generateUUID();
        case 'REFNAME':
          return `${data['EMPLOYEE NAME']} ${data['SURNAME']}`;
        case 'USER':
          return getCurrentUser();
        case 'TIMESTAMP':
          return new Date();
        default:
          return data[header] || null;
      }
    });

    sheet.appendRow(newRow);
    return { success: true, message: 'Employee added successfully.' };
  } catch (e) {
    Logger.error(e);
    return { success: false, message: 'An error occurred while adding the employee.' };
  }
}

/**
 * Updates an existing employee's details.
 * @param {string} id The unique ID of the employee to update.
 * @param {object} data The updated employee data.
 * @return {object} A success or error message.
 */
function updateEmployee(id, data) {
  try {
    // Omitting validation for brevity in this example, but it should be here
    const sheet = getSheet(EMPLOYEE_SHEET);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    const idColumnIndex = headers.indexOf('ID');

    for (let i = 1; i < values.length; i++) {
      if (values[i][idColumnIndex] === id) {
        let newRow = headers.map((header, index) => {
          if (header === 'MODIFIED_BY') return getCurrentUser();
          if (header === 'LAST_MODIFIED') return new Date();
          return data.hasOwnProperty(header) ? data[header] : values[i][index];
        });
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([newRow]);
        return { success: true, message: 'Employee updated successfully.' };
      }
    }
    return { success: false, message: 'Employee not found.' };
  } catch (e) {
    Logger.error(e);
    return { success: false, message: 'An error occurred while updating the employee.' };
  }
}

/**
 * Retrieves an employee by their unique ID.
 * @param {string} id The unique ID of the employee.
 * @return {object} The employee data.
 */
function getEmployeeById(id) {
  const sheet = getSheet(EMPLOYEE_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idCol = headers.indexOf('ID');

  for(const row of data) {
    if (row[idCol] === id) {
      const employee = {};
      headers.forEach((header, i) => employee[header] = row[i]);
      return employee;
    }
  }
  return null;
}

/**
 * Retrieves an employee by their name.
 * @param {string} name The name of the employee.
 * @return {object} The employee data.
 */
function getEmployeeByName(name) {
    const sheet = getSheet(EMPLOYEE_SHEET);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const nameCol = headers.indexOf('EMPLOYEE NAME');

    for(const row of data) {
        if (row[nameCol] === name) {
            const employee = {};
            headers.forEach((header, i) => employee[header] = row[i]);
            return employee;
        }
    }
    return null;
}

/**
 * Lists all employees, with optional filters.
 * @param {object} filters The filters to apply (e.g., employer, search term).
 * @return {Array<object>} A list of employees.
 */
function listEmployees(filters = {}) {
  Logger.log('Attempting to list employees from sheet: ' + EMPLOYEE_SHEET);
  const sheet = getSheet(EMPLOYEE_SHEET);
  if (!sheet) {
    Logger.error('Could not find the EMPLOYEE DETAILS sheet. Please check if it exists.');
    return [];
  }
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  const employees = data.map(row => {
    const employee = {};
    headers.forEach((header, i) => employee[header] = row[i]);
    return employee;
  });

  return employees.filter(emp => {
    let keep = true;
    if (filters.employer && emp.EMPLOYER !== filters.employer) {
      keep = false;
    }
    if (filters.searchTerm && !(emp['EMPLOYEE NAME'].toLowerCase().includes(filters.searchTerm.toLowerCase()) || emp['SURNAME'].toLowerCase().includes(filters.searchTerm.toLowerCase()))) {
        keep = false;
    }
    return keep;
  });
}

/**
 * Terminates an employee by setting their termination date.
 * @param {string} id The unique ID of the employee to terminate.
 * @param {string} terminationDate The date of termination.
 * @return {object} A success or error message.
 */
function terminateEmployee(id, terminationDate) {
    return updateEmployee(id, { 'TERMINATION DATE': terminationDate, 'EMPLOYMENT STATUS': 'Terminated' });
}

/**
 * Validates employee data.
 * @param {object} data The employee data to validate.
 * @return {object} An object with `isValid` and a list of `errors`.
 */
function validateEmployee(data) {
  const errors = [];
  const requiredFields = CONFIG.REQUIRED_EMPLOYEE_FIELDS;
  
  requiredFields.forEach(field => {
    if (!data[field]) {
      errors.push(`${field} is required.`);
    }
  });

  if (data['HOURLY RATE'] <= 0) {
    errors.push('Hourly rate must be greater than 0.');
  }
    
  if(!validateSAIdNumber(data['ID NUMBER'])) {
      errors.push('Invalid South African ID number.');
  }
    
  if(!validatePhoneNumber(data['CONTACT NUMBER'])) {
      errors.push('Invalid contact number format.');
  }

  // Check for uniqueness of ID NUMBER and ClockInRef
  const allEmployees = listEmployees();
  if (allEmployees.some(emp => emp['ID NUMBER'] === data['ID NUMBER'])) {
    errors.push('ID Number must be unique.');
  }
  if (allEmployees.some(emp => emp['ClockInRef'] === data['ClockInRef'])) {
    errors.push('ClockInRef must be unique.');
  }

  return { isValid: errors.length === 0, errors: errors };
}
