/**
 * Voltran Extension - PIN Validation Google Apps Script (ROBUST VERSION)
 * 
 * This Google Apps Script provides PIN validation for the Voltran Chrome Extension.
 * While the extension now uses direct Google Sheets reading to eliminate CORS issues,
 * this script serves as a backup validation method and for legacy compatibility.
 * 
 * ROBUST SHEET DETECTION:
 * - Automatically detects PIN sheets: "PINs", "Sheet1", "PIN", "pins", "password", "codes", etc.
 * - Analyzes sheet content to identify PIN-like data automatically
 * - Works regardless of sheet naming conventions
 * - Creates a default PIN sheet if none found
 * 
 * PIN SHEET STRUCTURE (Auto-detected):
 * - Auto-detected sheet names: "PINs", "Sheet1", "PIN", "pins", "password", "codes"
 * - Column A: PIN codes (starting from row 1 or 2, headers auto-detected)
 * - Example:
 *   A1: PIN (header - optional, auto-detected)
 *   A2: 123456
 *   A3: 789012
 *   A4: ABC123
 * 
 * SETUP INSTRUCTIONS:
 * 1. Replace SHEET_ID below with your PIN Google Sheet ID (the sheet containing your PIN list)
 * 2. Deploy as web app with execute permissions for "Anyone"
 * 3. Copy the web app URL to your extension's PinValidationWebAppUrl configuration
 * 4. The script will automatically find your PIN data regardless of sheet name!
 */

// CONFIGURATION - UPDATE THESE VALUES
const SHEET_ID = 'YOUR_SHEET_ID_HERE_FROM_PIN_LIST_SHEET'; // Replace with the Google Sheet ID that contains your PIN list
const PREFERRED_PIN_SHEET_NAME = 'PINs'; // Preferred name for PIN sheet (will be created if no sheet found)

/**
 * Handles GET requests (CORS preflight and basic requests)
 */
function doGet(e) {
  Logger.log('Received GET request with parameters: ' + JSON.stringify(e.parameter));
  return handleCors();
}

/**
 * Handles POST requests for PIN validation
 */
function doPost(e) {
  try {
    // Handle CORS preflight requests
    if (e.parameter.method === 'OPTIONS' || !e.postData) {
      Logger.log('Received OPTIONS or empty postData request. Returning CORS headers.');
      return handleCors();
    }

    // Log the incoming request for debugging
    Logger.log('Received POST request. Post data: ' + e.postData.contents);

    // Parse the incoming JSON payload from the extension
    var data = JSON.parse(e.postData.contents);
    var pin = data.pin || '';
    Logger.log('Parsed PIN from payload: ' + pin);

    // Validate the PIN
    var validationResult = validatePinFromSheet(pin);
    
    // Return validation result
    var response = {
      status: validationResult.isValid ? 'success' : 'error',
      message: validationResult.message,
      pin: pin,
      timestamp: new Date().toISOString()
    };
    
    Logger.log('Validation response: ' + JSON.stringify(response));
    return createJsonResponse(response);

  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    var errorResponse = {
      status: 'error',
      message: 'Internal server error: ' + error.toString(),
      timestamp: new Date().toISOString()
    };
    return createJsonResponse(errorResponse);
  }
}

/**
 * Validates a PIN against the Google Sheet data
 * @param {string} pin - The PIN to validate
 * @returns {Object} - Validation result with isValid boolean and message
 */
function validatePinFromSheet(pin) {
  try {
    // Open the Google Spreadsheet by ID
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('Opened spreadsheet for PINs with ID: ' + SHEET_ID);
    
    // Use robust sheet detection to find the PIN sheet
    var sheet = getOrCreatePinSheet(spreadsheet);
    if (!sheet) {
      Logger.log('Error: No PIN sheet found and unable to create one.');
      return {
        isValid: false,
        message: 'PIN validation sheet not available.'
      };
    }
    
    Logger.log('Using PIN sheet: ' + sheet.getName());
    
    // Analyze the sheet structure to determine data range
    var structure = analyzePinSheetStructure(sheet);
    
    if (structure.endRow < structure.startRow) {
      Logger.log('No PIN data found in sheet.');
      return {
        isValid: false,
        message: 'No PIN data available.'
      };
    }
    
    // Get PIN data from the determined range
    var pinRange = sheet.getRange(structure.startRow, structure.dataColumn, 
                                  structure.endRow - structure.startRow + 1, 1);
    var pinList = pinRange.getDisplayValues(); // Use display values to handle formatting consistently
    
    Logger.log('Retrieved PIN list from ' + sheet.getName() + ' (rows ' + structure.startRow + '-' + structure.endRow + '): ' + pinList.length + ' entries');

    // Check if the provided PIN exists in the list
    var inputPin = pin.toString().trim();
    Logger.log('Validating PIN: [' + inputPin.length + ' characters]');
    
    for (var i = 0; i < pinList.length; i++) {
      var storedPin = pinList[i][0];
      if (storedPin) {
        storedPin = storedPin.toString().trim();
        
        // Try exact match first
        if (storedPin === inputPin) {
          Logger.log('Exact PIN match found at row ' + (structure.startRow + i));
          return {
            isValid: true,
            message: 'PIN validated successfully.'
          };
        }
        
        // Try case-insensitive match for alphanumeric PINs
        if (storedPin.toUpperCase() === inputPin.toUpperCase()) {
          Logger.log('Case-insensitive PIN match found at row ' + (structure.startRow + i));
          return {
            isValid: true,
            message: 'PIN validated successfully.'
          };
        }
      }
    }
    
    Logger.log('PIN validation result: Invalid - no match found among ' + pinList.length + ' PINs');
    return {
      isValid: false,
      message: 'Invalid PIN. Please check your PIN and try again.'
    };
    
  } catch (error) {
    Logger.log('Error in validatePinFromSheet: ' + error.toString());
    return {
      isValid: false,
      message: 'Error validating PIN: ' + error.toString()
    };
  }
}

/**
 * Handles CORS by returning appropriate headers
 * Note: Google Apps Script automatically handles CORS for web apps deployed as "Anyone" access
 */
function handleCors() {
  var output = ContentService.createTextOutput('CORS handled');
  output.setMimeType(ContentService.MimeType.JSON);
  
  return output;
}

/**
 * Creates a JSON response with proper CORS handling
 * Note: Google Apps Script automatically handles CORS for web apps deployed as "Anyone" access
 * @param {Object} data - The data to return as JSON
 */
function createJsonResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  
  return output;
}

/**
 * Test function to verify the script works correctly
 * Run this function in the Apps Script editor to test
 */
function testPinValidation() {
  Logger.log('Starting PIN validation test...');
  
  // Test with a known PIN (replace with actual PIN from your sheet)
  var testPin = '123456';
  var result = validatePinFromSheet(testPin);
  
  Logger.log('Test result for PIN "' + testPin + '":');
  Logger.log('- Valid: ' + result.isValid);
  Logger.log('- Message: ' + result.message);
  
  // Test with an invalid PIN
  var invalidPin = '999999';
  var invalidResult = validatePinFromSheet(invalidPin);
  
  Logger.log('Test result for invalid PIN "' + invalidPin + '":');
  Logger.log('- Valid: ' + invalidResult.isValid);
  Logger.log('- Message: ' + invalidResult.message);
}

/**
 * Utility function to get all PINs from the sheet (for debugging)
 */
function getAllPins() {
  try {
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    var sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      Logger.log('Sheet not found: ' + SHEET_NAME);
      return [];
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('No data in sheet');
      return [];
    }
    
    var pinRange = sheet.getRange('A2:A' + lastRow);
    var pinList = pinRange.getValues();
    
    var pins = pinList.map(function(row) {
      return row[0].toString().trim();
    }).filter(function(pin) {
      return pin.length > 0;
    });
    
    Logger.log('All PINs in sheet: ' + JSON.stringify(pins));
    return pins;
    
  } catch (error) {
    Logger.log('Error getting all PINs: ' + error.toString());
    return [];
  }
}

/**
 * ROBUST PIN SHEET DETECTION FUNCTIONS
 * These functions automatically detect PIN sheets regardless of naming conventions
 */

/**
 * Intelligently finds the PIN validation sheet by checking multiple possible names
 * @param {Spreadsheet} spreadsheet - The spreadsheet to search in
 * @returns {Sheet|null} - The found sheet or null
 */
function findPinSheet(spreadsheet) {
  // List of possible sheet names for PINs (case-insensitive)
  var possibleNames = [
    'PINs',
    'PIN',
    'Pins',
    'Pin',
    'Sheet1',
    'pins',
    'pin',
    'Password',
    'Passwords',
    'password',
    'passwords',
    'Access Codes',
    'AccessCodes',
    'access_codes',
    'access codes',
    'Codes',
    'codes',
    'Auth',
    'auth',
    'Authentication',
    'Key',
    'Keys',
    'keys'
  ];
  
  var sheets = spreadsheet.getSheets();
  Logger.log('Available sheets: ' + sheets.map(s => s.getName()).join(', '));
  
  // First try exact matches
  for (var i = 0; i < possibleNames.length; i++) {
    var sheet = spreadsheet.getSheetByName(possibleNames[i]);
    if (sheet) {
      // Verify this looks like a PIN sheet
      if (isPinSheet(sheet)) {
        Logger.log('Found PIN sheet (exact match): ' + possibleNames[i]);
        return sheet;
      }
    }
  }
  
  // Then try case-insensitive matches
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName().toLowerCase();
    for (var j = 0; j < possibleNames.length; j++) {
      if (sheetName === possibleNames[j].toLowerCase()) {
        if (isPinSheet(sheets[i])) {
          Logger.log('Found PIN sheet (case-insensitive): ' + sheets[i].getName());
          return sheets[i];
        }
      }
    }
  }
  
  // Look for sheets that contain PIN-like data by content analysis
  for (var i = 0; i < sheets.length; i++) {
    if (isPinSheet(sheets[i])) {
      Logger.log('Found PIN sheet by content analysis: ' + sheets[i].getName());
      return sheets[i];
    }
  }
  
  Logger.log('No PIN sheet found');
  return null;
}

/**
 * Checks if a sheet appears to contain PIN data
 * @param {Sheet} sheet - The sheet to analyze
 * @returns {boolean} - Whether this looks like a PIN sheet
 */
function isPinSheet(sheet) {
  try {
    if (sheet.getLastRow() < 1) return false;
    
    // Check for PIN-like headers in first few rows
    var maxRowsToCheck = Math.min(3, sheet.getLastRow());
    for (var row = 1; row <= maxRowsToCheck; row++) {
      if (sheet.getLastColumn() >= 1) {
        var possibleHeaders = sheet.getRange(row, 1, 1, Math.min(sheet.getLastColumn(), 3)).getValues()[0];
        var headerText = possibleHeaders.join('|').toLowerCase();
        
        if (headerText.includes('pin') || headerText.includes('password') || 
            headerText.includes('code') || headerText.includes('auth') ||
            headerText.includes('key') || headerText.includes('access')) {
          Logger.log('Found PIN-like headers in row ' + row + ': ' + headerText);
          return true;
        }
      }
    }
    
    // Check for PIN-like data in column A (numeric/alphanumeric codes)
    if (sheet.getLastRow() >= 1) {
      var startRow = 1;
      var maxDataRowsToCheck = Math.min(10, sheet.getLastRow());
      
      // Try different starting rows (in case row 1 is header)
      for (var startAttempt = 1; startAttempt <= Math.min(2, sheet.getLastRow()); startAttempt++) {
        var dataRange = sheet.getRange(startAttempt, 1, Math.min(maxDataRowsToCheck, sheet.getLastRow() - startAttempt + 1), 1);
        var data = dataRange.getValues();
        
        var pinLikeCount = 0;
        var validDataCount = 0;
        
        for (var i = 0; i < data.length; i++) {
          var value = data[i][0];
          if (value && value.toString().trim().length > 0) {
            validDataCount++;
            var strValue = value.toString().trim();
            
            // Check for PIN-like patterns
            if (/^\d{3,8}$/.test(strValue) || // 3-8 digit numeric codes
                /^[A-Za-z0-9]{4,12}$/.test(strValue) || // 4-12 character alphanumeric codes
                /^\d{4}-\d{4}$/.test(strValue) || // 4-4 digit patterns
                /^[A-Z]{2,4}\d{2,6}$/.test(strValue)) { // Letter+number patterns
              pinLikeCount++;
            }
          }
        }
        
        // If most valid data looks like PINs, consider this a PIN sheet
        if (validDataCount > 0 && pinLikeCount >= validDataCount * 0.6) {
          Logger.log('Found PIN-like data starting from row ' + startAttempt + ': ' + pinLikeCount + '/' + validDataCount + ' entries look like PINs');
          return true;
        }
      }
    }
    
    return false;
  } catch (e) {
    Logger.log('Error analyzing sheet ' + sheet.getName() + ': ' + e.toString());
    return false;
  }
}

/**
 * Gets or creates the PIN sheet with proper setup
 * @param {Spreadsheet} spreadsheet - The spreadsheet to work with
 * @returns {Sheet|null} - The PIN sheet or null if creation fails
 */
function getOrCreatePinSheet(spreadsheet) {
  var sheet = findPinSheet(spreadsheet);
  
  if (!sheet) {
    try {
      // Create new sheet with preferred name
      sheet = spreadsheet.insertSheet(PREFERRED_PIN_SHEET_NAME);
      Logger.log('Created new PIN sheet: ' + PREFERRED_PIN_SHEET_NAME);
      
      // Set up example PINs and header
      sheet.getRange('A1').setValue('PIN');
      sheet.getRange('A1').setFontWeight('bold');
      sheet.getRange('A1').setBackground('#f0f0f0');
      
      // Add some example PINs (these should be replaced with real ones)
      sheet.getRange('A2').setValue('123456');
      sheet.getRange('A3').setValue('789012');
      sheet.getRange('A4').setValue('000000');
      
      Logger.log('Set up example PINs in new PIN sheet');
    } catch (e) {
      Logger.log('Failed to create PIN sheet: ' + e.toString());
      return null;
    }
  }
  
  return sheet;
}

/**
 * Analyzes a PIN sheet to determine the correct data range
 * @param {Sheet} sheet - The PIN sheet to analyze
 * @returns {Object} - Object with startRow, endRow, and column information
 */
function analyzePinSheetStructure(sheet) {
  var result = {
    startRow: 1,
    endRow: sheet.getLastRow(),
    dataColumn: 1,
    hasHeader: false
  };
  
  if (sheet.getLastRow() < 1) {
    return result;
  }
  
  // Check if row 1 looks like a header
  try {
    var firstRowValue = sheet.getRange(1, 1).getValue();
    if (firstRowValue && typeof firstRowValue === 'string') {
      var headerText = firstRowValue.toString().toLowerCase().trim();
      if (headerText === 'pin' || headerText === 'pins' || 
          headerText === 'password' || headerText === 'code' ||
          headerText === 'auth' || headerText === 'key' ||
          headerText.includes('pin') || headerText.includes('code')) {
        result.hasHeader = true;
        result.startRow = 2;
        Logger.log('Detected header row: "' + firstRowValue + '"');
      }
    }
  } catch (e) {
    Logger.log('Error checking for header: ' + e.toString());
  }
  
  // Validate that we have actual data rows
  if (result.startRow > sheet.getLastRow()) {
    Logger.log('No data rows available after header');
    result.startRow = 1;
    result.hasHeader = false;
  }
  
  Logger.log('PIN sheet structure: startRow=' + result.startRow + ', endRow=' + result.endRow + ', hasHeader=' + result.hasHeader);
  return result;
}
