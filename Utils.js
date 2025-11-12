/**
 * @fileoverview Utility functions for the HR system.
 */

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

/**
 * Gets a specific Google Sheet by name.
 * @param {string} sheetName The name of the sheet to retrieve.
 * @return {Sheet} The Google Sheet object.
 */
function getSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      console.error('Sheet not found:', sheetName);
      return null;
    }
    console.log('Successfully retrieved sheet:', sheetName);
    return sheet;
  } catch (e) {
    console.error('Error in getSheet:', e);
    return null;
  }
}

/**
 * Gets the current active user's email.
 * @return {string} The user's email address.
 */
function getCurrentUser() {
  return Session.getActiveUser().getEmail();
}

/**
 * Formats a date object into a string (YYYY-MM-DD).
 * @param {Date} date The date to format.
 * @return {string} The formatted date string.
 */
function formatDate(date) {
    if (!date) return null;
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/**
 * Formats a number into a currency string (South African Rand).
 * @param {number} amount The amount to format.
 * @return {string} The formatted currency string.
 */
function formatCurrency(amount) {
  return "R" + Number(amount).toFixed(2);
}

/**
 * Generates a 6-digit random number for employee ID.
 * @return {string} The generated 6-digit ID.
 */
function generateUUID() {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

/**
 * Validates a South African ID number.
 * @param {string} idNumber The ID number to validate.
 * @return {boolean} True if the ID number is valid, false otherwise.
 */
function validateSAIdNumber(idNumber) {
  // Basic check for length and numeric characters
  if (!idNumber || idNumber.length !== 13 || isNaN(idNumber)) {
    return false;
  }
  // This is a placeholder for a proper Luhn algorithm check if needed
  return true;
}

/**
 * Validates a phone number.
 * @param {string} phoneNumber The phone number to validate.
 * @return {boolean} True if the phone number is valid, false otherwise.
 */
function validatePhoneNumber(phoneNumber) {
  // Basic check for South African phone numbers
  const regex = /^(?:\+27|0)[6-8][0-9]{8}$/;
  return regex.test(phoneNumber);
}

/**
 * Logger helper functions.
 */
const Logger = {
  log: function(message) {
    console.log(JSON.stringify(message, null, 2));
  },
  error: function(message) {
    console.error(JSON.stringify(message, null, 2));
  }
};

/**
 * Gets the number of records in a sheet.
 * @param {string} sheetName The name of the sheet.
 * @return {number} The number of records (rows - 1), or -1 if an error occurs.
 */
function getRecordCount(sheetName) {
  try {
    const sheet = getSheet(sheetName);
    if (!sheet) {
      console.error('Sheet not found for record count:', sheetName);
      return -1;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      return 0;
    }
    return lastRow - 1;
  } catch (e) {
    console.error('Error in getRecordCount:', e);
    return -1;
  }
}
