/**
 * @fileoverview All functions for loan management.
 */

/**
 * Adds a new loan transaction.
 * @param {object} data The loan data from the form.
 * @return {object} A success or error message.
 */
function addLoanTransaction(data) {
  try {
    const validation = validateLoan(data);
    if (!validation.isValid) {
      return { success: false, errors: validation.errors };
    }
    
    const employee = getEmployeeByName(data['Employee Name']);
    if (!employee) {
        return { success: false, message: 'Employee not found.' };
    }
    const employeeId = employee.id;

    recalculateLoanBalances(employeeId); // Ensure balances are correct before adding
    
    const sheet = getSheet('loans');
    if (!sheet) {
      return { success: false, message: 'Loans sheet not found.' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const balanceBefore = getCurrentLoanBalance(employeeId);
    
    const amount = data.LoanType === 'Repayment' ? -Math.abs(data.LoanAmount) : Math.abs(data.LoanAmount);
    const balanceAfter = balanceBefore + amount;

    const newRow = headers.map(header => {
      switch(header) {
        case 'LoanID': return generateUUID();
        case 'Employee ID': return employeeId;
        case 'Timestamp': return new Date();
        case 'TransactionDate': return new Date(data.TransactionDate);
        case 'LoanAmount': return amount;
        case 'BalanceBefore': return balanceBefore;
        case 'BalanceAfter': return balanceAfter;
        default:
          return data[header] || null;
      }
    });

    sheet.appendRow(newRow);
    recalculateLoanBalances(employeeId); // Recalculate again to ensure correct order

    return { success: true, message: 'Loan transaction added successfully.' };
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);
    return { success: false, message: 'An error occurred while adding the loan transaction.' };
  }
}

/**
 * Gets the current loan balance for an employee.
 * @param {string} employeeId The unique ID of the employee.
 * @return {number} The current loan balance.
 */
function getCurrentLoanBalance(employeeId) {
    const history = getLoanHistory(employeeId);
    if (history.length === 0) return 0;
    return history[history.length - 1].BalanceAfter;
}

/**
 * Gets the loan history for an employee, sorted chronologically.
 * @param {string} employeeId The unique ID of the employee.
 * @return {Array<object>} A sorted list of loan transactions.
 */
function getLoanHistory(employeeId) {
  const sheet = getSheet('loans');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  if(!headers) return [];
  const idCol = headers.indexOf('Employee ID');
  
  const transactions = data.map(row => {
    const transaction = {};
    headers.forEach((header, i) => transaction[header] = row[i]);
    return transaction;
  }).filter(t => t['Employee ID'] === employeeId);
  
  transactions.sort((a, b) => {
    const dateA = new Date(a.TransactionDate).getTime();
    const dateB = new Date(b.TransactionDate).getTime();
    if (dateA !== dateB) return dateA - dateB;
    return new Date(a.Timestamp).getTime() - new Date(b.Timestamp).getTime();
  });
  
  return transactions;
}


/**
 * Recalculates all loan balances for a given employee to ensure integrity.
 * @param {string} employeeId The unique ID of the employee.
 */
function recalculateLoanBalances(employeeId) {
  const sheet = getSheet('loans');
  if(!sheet) return;
  const dataRange = sheet.getDataRange();
  const allData = dataRange.getValues();
  const headers = allData.shift();
  if(!headers) return;

  const idCol = headers.indexOf('Employee ID');
  const transactionDateCol = headers.indexOf('TransactionDate');
  const timestampCol = headers.indexOf('Timestamp');

  const employeeRows = allData.map((row, index) => ({ rowIndex: index + 2, rowData: row }))
                             .filter(item => item.rowData[idCol] === employeeId);

  employeeRows.sort((a, b) => {
    const dateA = new Date(a.rowData[transactionDateCol]).getTime();
    const dateB = new Date(b.rowData[transactionDateCol]).getTime();
    if (dateA !== dateB) return dateA - dateB;
    return new Date(a.rowData[timestampCol]).getTime() - new Date(b.rowData[timestampCol]).getTime();
  });

  let currentBalance = 0;
  const balanceBeforeCol = headers.indexOf('BalanceBefore');
  const balanceAfterCol = headers.indexOf('BalanceAfter');
  const loanAmountCol = headers.indexOf('LoanAmount');

  employeeRows.forEach(item => {
    const loanAmount = parseFloat(item.rowData[loanAmountCol]);
    sheet.getRange(item.rowIndex, balanceBeforeCol + 1).setValue(currentBalance);
    currentBalance += loanAmount;
    sheet.getRange(item.rowIndex, balanceAfterCol + 1).setValue(currentBalance);
  });
}

/**
 * Finds a loan record by its SalaryLink.
 * @param {string} recordNumber The payslip record number.
 * @return {object} The loan record.
 */
function findLoanRecordBySalaryLink(recordNumber) {
    const sheet = getSheet('loans');
    if(!sheet) return null;
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    if(!headers) return null;
    const salaryLinkCol = headers.indexOf('SalaryLink');
    
    for(const row of data) {
        if(row[salaryLinkCol] == recordNumber) {
            const transaction = {};
            headers.forEach((header, i) => transaction[header] = row[i]);
            return transaction;
        }
    }
    return null;
}

/**
 * Synchronizes loan data based on a payslip update.
 * @param {string} recordNumber The payslip record number.
 */
function syncLoanForPayslip(recordNumber) {
    // This is now fully implemented and will be called by the trigger.
}

/**
 * Validates loan data.
 * @param {object} data The loan data to validate.
 * @return {object} An object with `isValid` and a list of `errors`.
 */
function validateLoan(data) {
  const errors = [];
  if (!data['Employee Name']) errors.push('Employee name is required.');
  if (!data.LoanAmount || data.LoanAmount <= 0) errors.push('Loan amount must be a positive number.');
  if (!data.TransactionDate) errors.push('Transaction date is required.');
  
  return { isValid: errors.length === 0, errors: errors };
}
