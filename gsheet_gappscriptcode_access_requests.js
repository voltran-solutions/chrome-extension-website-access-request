/**
 * Voltran Extension - Access Request Logging Google Apps Script
 * 
 * This Google Apps Script logs website access requests from the Voltran Chrome Extension.
 * It records PIN information, timestamps, and requested URLs for security auditing.
 * 
 * ROBUST SHEET DETECTION:
 * - Automatically detects access request sheets: "AccessRequests", "Sheet1", "access", "requests", etc.
 * - Automatically detects PIN validation sheets: "PINs", "Sheet1", "pins", "password", etc.
 * - Creates sheets with appropriate names if none found
 * 
 * ACCESS REQUEST SHEET STRUCTURE:
 * - Auto-detected sheet names: "AccessRequests", "Sheet1", "access", "requests"
 * - Columns: A=Timestamp, B=PIN Number, C=User Email, D=Title, E=URL, F=Media Type, G=Request Status, H=Access Link
 * - Note: Columns F, G, H are for manual review and will be filled by admins later
 * - Example:
 *   A1: Timestamp | B1: PIN Number | C1: User Email | D1: Title | E1: URL | F1: Media Type | G1: Request Status | H1: Access Link
 *   A2: Friday, Jun 06, 2025 10:30 AM | 1234 | user@company.com | Example Site | https://example.com | PENDING | PENDING | PENDING
 * 
 * SETUP INSTRUCTIONS:
 * 1. Replace SHEET_ID below with your ACCESS REQUEST Google Sheet ID (for logging user activity)
 * 2. Optionally set PIN_SHEET_ID if PIN validation is in a different spreadsheet
 * 3. Deploy as web app with execute permissions for "Anyone"
 * 4. Copy the web app URL to your extension's AccessRequestSheetWebAppUrl configuration
 */

// CONFIGURATION - UPDATE THESE VALUES
const SHEET_ID = 'YOUR_SHEET_ID_HERE_FROM_ACCESS_REQUEST_SHEET'; // Replace with the Google Sheet ID for logging access requests
const PIN_SHEET_ID = 'YOUR_SHEET_ID_HERE_FROM_PIN_LIST_SHEET'; // Replace with PIN sheet ID if different, or leave same for single spreadsheet
const PREFERRED_ACCESS_SHEET_NAME = 'AccessRequests'; // Preferred name for access request sheet (will be created if no sheet found)
const PREFERRED_PIN_SHEET_NAME = 'PINs'; // Preferred name for PIN sheet (will be created if no sheet found)

/**
 * Creates a human-readable timestamp in format: "Friday, Jun 06, 2025 10:30 AM"
 * @returns {string} - Formatted timestamp
 */
function createReadableTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "EEEE, MMM dd, yyyy hh:mm a");
}

/**
 * Handles GET requests (CORS preflight and data retrieval)
 * Note: Google Apps Script automatically handles CORS when deployed as web app with "Anyone" access
 */
function doGet(e) {
  Logger.log('Received GET request with parameters: ' + JSON.stringify(e.parameter));
  
  // Check if this is a data retrieval request
  if (e.parameter && e.parameter.action === 'getData') {
    return handleDataRetrieval(e.parameter);
  }
  
  // Default CORS handling for other GET requests
  return handleCors();
}

/**
 * Handles data retrieval requests for the AAP extension
 * @param {Object} params - Query parameters including userEmail
 * @returns {ContentService.TextOutput} - JSON response with user's data
 */
function handleDataRetrieval(params) {
  try {
    var userEmail = params.userEmail;
    
    Logger.log('Data retrieval request for user: ' + userEmail);
    
    if (!userEmail) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'userEmail parameter is required' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Get the access request sheet
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    var sheet = findAccessRequestSheet(spreadsheet);
    
    if (!sheet) {
      Logger.log('No access request sheet found');
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'No access request sheet found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    Logger.log('Found sheet: ' + sheet.getName());
    
    // Get all data from the sheet
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      Logger.log('No data rows found (only header row exists)');
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Get headers (row 1) and data (rows 2+)
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    var dataRows = dataRange.getValues();
    
    Logger.log('Found ' + dataRows.length + ' data rows');
    Logger.log('Headers: ' + headers.join(', '));
    
    // Convert to objects and filter by user email
    var userRows = [];
    
    for (var i = 0; i < dataRows.length; i++) {
      var row = dataRows[i];
      var rowData = {};
      
      // Map each column to header
      for (var j = 0; j < headers.length && j < row.length; j++) {
        var header = headers[j].toString().trim();
        var value = row[j];
        
        // Convert dates to strings if needed
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a");
        } else if (value !== null && value !== undefined) {
          value = value.toString().trim();
        } else {
          value = '';
        }
        
        rowData[header] = value;
      }
      
      // Check if this row belongs to the requesting user
      var rowUserEmail = '';
      
      // Try different possible email column names
      var emailColumns = ['User Email', 'userEmail', 'email', 'Email', 'user_email'];
      for (var k = 0; k < emailColumns.length; k++) {
        if (rowData[emailColumns[k]]) {
          rowUserEmail = rowData[emailColumns[k]].toLowerCase();
          break;
        }
      }
      
      if (rowUserEmail === userEmail.toLowerCase()) {
        userRows.push(rowData);
      }
    }
    
    Logger.log('Returning ' + userRows.length + ' rows for user: ' + userEmail);
    
    return ContentService
      .createTextOutput(JSON.stringify(userRows))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('Error in handleDataRetrieval: ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Failed to retrieve data: ' + error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles POST requests for logging access requests
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
    
    // Extract required fields (User ID removed)
    var url = data.url || '';
    var title = data.title || '';
    var timestamp = data.timestamp || createReadableTimestamp(); // Use readable timestamp format
    var userEmail = data.userEmail || '';
    var pin = data.pin || '';
    
    Logger.log('Parsed access request data:');
    Logger.log('- URL: ' + url);
    Logger.log('- Title: ' + title);
    Logger.log('- User Email: ' + userEmail);
    // Log the actual PIN number (masked in logs for security)
    Logger.log('- PIN: ' + (pin ? '[' + pin.length + ' characters]' : '[NOT PROVIDED]'));
    Logger.log('- Timestamp: ' + timestamp);

    // Validate PIN if provided (optional - logs all requests regardless)
    var isPinValid = false;
    if (pin) {
      isPinValid = validatePin(pin);
      Logger.log('PIN validation result: ' + (isPinValid ? 'VALID' : 'INVALID'));
    } else {
      Logger.log('No PIN provided - logging request without validation');
    }

    // Log the access request to the sheet (logs all requests, valid and invalid)
    var logResult = logAccessRequest({
      url: url,
      title: title,
      timestamp: timestamp,
      userEmail: userEmail,
      pin: pin,
      isPinValid: isPinValid
    });
    
    // Return success response with PIN validation status
    var response = {
      status: logResult.success ? 'success' : 'error',
      message: logResult.message,
      pinValid: isPinValid,
      pinProvided: !!pin,
      timestamp: createReadableTimestamp() // Use readable timestamp format
    };
    
    Logger.log('Access request response: ' + JSON.stringify(response));
    return createJsonResponse(response);

  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    var errorResponse = {
      status: 'error',
      message: 'Internal server error: ' + error.toString(),
      timestamp: createReadableTimestamp() // Use readable timestamp format
    };
    return createJsonResponse(errorResponse);
  }
}

/**
 * Logs an access request to the Google Sheet
 * @param {Object} requestData - The access request data
 * @returns {Object} - Result with success boolean and message
 */
function logAccessRequest(requestData) {
  try {
    // Open the Google Spreadsheet by ID
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('Opened spreadsheet for access requests with ID: ' + SHEET_ID);
    
    // Use robust sheet detection to find or create the access request sheet
    var sheet = getOrCreateAccessRequestSheet(spreadsheet);
    Logger.log('Using access request sheet: ' + sheet.getName());
    
    // Always ensure correct headers (create or update if needed)
    var expectedHeaders = ['Timestamp', 'PIN Number', 'User Email', 'Title', 'URL', 'Request Status', 'Media Type', 'Access Link'];
    var currentHeaders = [];
    
    try {
      if (sheet.getLastRow() >= 1) {
        var lastCol = sheet.getLastColumn();
        if (lastCol > 0) {
          currentHeaders = sheet.getRange(1, 1, 1, Math.min(lastCol, 8)).getValues()[0];
        }
      }
    } catch (e) {
      Logger.log('Could not read existing headers, will create new ones');
    }
    
    // Check if headers need to be updated
    var headersNeedUpdate = false;
    if (currentHeaders.length !== expectedHeaders.length) {
      headersNeedUpdate = true;
    } else {
      for (var i = 0; i < expectedHeaders.length; i++) {
        if (currentHeaders[i] !== expectedHeaders[i]) {
          headersNeedUpdate = true;
          break;
        }
      }
    }
    
    // Check for extra columns beyond H
    var lastCol = sheet.getLastColumn();
    if (lastCol > 8) {
      Logger.log('Found extra columns beyond H (up to column ' + lastCol + '), cleaning up');
      headersNeedUpdate = true;
    }
    
    if (headersNeedUpdate) {
      // Clear any extra columns beyond H first
      if (lastCol > 8) {
        var extraRange = sheet.getRange(1, 9, sheet.getMaxRows(), lastCol - 8);
        extraRange.clearContent();
        extraRange.clearFormat();
      }
      
      // Clear the entire first row to ensure clean state
      if (sheet.getLastColumn() > 0) {
        sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 8)).clearContent();
        sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 8)).clearFormat();
      }
      
      // Set the correct headers ONLY in A1:H1
      sheet.getRange('A1:H1').setValues([expectedHeaders]);
      
      // Format headers
      sheet.getRange('A1:H1').setFontWeight('bold');
      sheet.getRange('A1:H1').setBackground('#f0f0f0');
      
      Logger.log('Updated headers to: ' + expectedHeaders.join(', '));
      Logger.log('Cleaned up extra columns beyond H');
    }
    
    // Prepare the row data in the new column order (with review columns empty)
    var rowData = [
      requestData.timestamp || createReadableTimestamp(), // Timestamp (A)
      requestData.pin || 'NO PIN',    // PIN Number (B)
      requestData.userEmail,          // User Email (C)
      requestData.title,              // Title (D)
      requestData.url,                // URL (E)
      'PENDING',                      // Request Status (F) - to be filled by admin
      'PENDING',                      // Media Type (G) - to be filled by admin
      'PENDING'                       // Access Link (H) - to be filled by admin
    ];
    
    // Add the data to the next available row
    sheet.appendRow(rowData);
    
    Logger.log('Successfully logged access request for user: ' + requestData.userEmail + ' with PIN: ' + (requestData.pin || 'NO PIN'));
    
    // Auto-fix any timestamp formatting issues in the sheet
    var timestampFixResult = autoFixTimestampFormats(sheet);
    if (timestampFixResult.fixed > 0) {
      Logger.log('Auto-fixed ' + timestampFixResult.fixed + ' timestamp formats');
    }
    if (timestampFixResult.errors > 0) {
      Logger.log('Had ' + timestampFixResult.errors + ' errors while fixing timestamps');
    }
    
    // Auto-resize columns for better readability
    sheet.autoResizeColumns(1, 8);
    
    return {
      success: true,
      message: 'Access request logged successfully.'
    };
    
  } catch (error) {
    Logger.log('Error in logAccessRequest: ' + error.toString());
    return {
      success: false,
      message: 'Error logging access request: ' + error.toString()
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
 * Utility function to validate and fix sheet headers
 */
function validateAndFixHeaders() {
  try {
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    
    // Use robust sheet detection to find the access request sheet
    var sheet = getOrCreateAccessRequestSheet(spreadsheet);
    Logger.log('Using access request sheet for header validation: ' + sheet.getName());
    
    var expectedHeaders = ['Timestamp', 'PIN Number', 'User Email', 'Title', 'URL', 'Request Status', 'Media Type', 'Access Link'];
    var currentHeaders = [];
    
    try {
      if (sheet.getLastRow() >= 1) {
        // Read all columns to see what we have
        var lastCol = sheet.getLastColumn();
        if (lastCol > 0) {
          currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        }
      }
    } catch (e) {
      Logger.log('Could not read existing headers, will create new ones');
    }
    
    Logger.log('Current headers found:', currentHeaders);
    Logger.log('Current column count:', currentHeaders.length);
    Logger.log('Last column in sheet:', sheet.getLastColumn());
    
    // More aggressive cleanup of extra columns beyond H (column 8)
    var lastCol = sheet.getLastColumn();
    if (lastCol > 8) {
      Logger.log('Found extra columns beyond H, clearing columns I through ' + columnToLetter(lastCol));
      
      // Clear all content in extra columns
      var extraRange = sheet.getRange(1, 9, sheet.getMaxRows(), lastCol - 8);
      extraRange.clearContent();
      extraRange.clearFormat();
      
      // Also clear any data validation, notes, etc.
      extraRange.clearDataValidations();
      extraRange.clearNote();
      
      Logger.log('Cleared extra columns I through ' + columnToLetter(lastCol));
    }
    
    // Clear the entire first row first to ensure no remnants
    if (sheet.getLastColumn() > 0) {
      sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 8)).clearContent();
      sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 8)).clearFormat();
    }
    
    // Set the correct headers ONLY in A1:H1
    sheet.getRange('A1:H1').setValues([expectedHeaders]);
    sheet.getRange('A1:H1').setFontWeight('bold');
    sheet.getRange('A1:H1').setBackground('#f0f0f0');
    
    // Ensure no extra formatting beyond column H
    if (sheet.getLastColumn() > 8) {
      var extraFormatRange = sheet.getRange(1, 9, 1, sheet.getLastColumn() - 8);
      extraFormatRange.clearFormat();
    }
    
    Logger.log('Headers validated and set to: ' + expectedHeaders.join(', '));
    Logger.log('Previous headers were: ' + (currentHeaders.length > 0 ? currentHeaders.join(', ') : '[NONE]'));
    Logger.log('Final column count after cleanup:', sheet.getLastColumn());
    
    return { success: true, message: 'Headers validated and updated' };
    
  } catch (error) {
    Logger.log('Error validating headers: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

/**
 * Helper function to convert column number to letter (1=A, 2=B, etc.)
 */
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Test function to verify the script works correctly
 * Run this function in the Apps Script editor to test
 */
function testAccessRequestLogging() {
  Logger.log('Starting access request logging test...');
  
  // First validate headers
  var headerResult = validateAndFixHeaders();
  Logger.log('Header validation result: ' + JSON.stringify(headerResult));
  
  // Test PIN validation functionality
  Logger.log('\n=== Testing PIN Validation ===');
  var testPin = '123456';
  var pinValidResult = validatePin(testPin);
  Logger.log('PIN validation test for "' + testPin + '": ' + (pinValidResult ? 'VALID' : 'INVALID'));
  
  // Test data with actual PIN
  var testData = {
    url: 'https://example.com',
    title: 'Example Website',
    timestamp: createReadableTimestamp(), // Use readable timestamp format
    userEmail: 'test@example.com',
    pin: testPin,
    isPinValid: pinValidResult
  };
  
  Logger.log('\n=== Testing Access Request Logging ===');
  var result = logAccessRequest(testData);
  
  Logger.log('Test result:');
  Logger.log('- Success: ' + result.success);
  Logger.log('- Message: ' + result.message);
  Logger.log('- PIN logged: ' + testData.pin);
  Logger.log('- PIN was valid: ' + pinValidResult);
  
  // Test with invalid PIN
  Logger.log('\n=== Testing Invalid PIN ===');
  var invalidTestData = {
    url: 'https://invalid-test.com',
    title: 'Invalid PIN Test',
    timestamp: createReadableTimestamp(), // Use readable timestamp format
    userEmail: 'invalid-test@example.com',
    pin: '999999', // Likely invalid PIN
    isPinValid: validatePin('999999')
  };
  
  var invalidResult = logAccessRequest(invalidTestData);
  Logger.log('Invalid PIN test result:');
  Logger.log('- Success: ' + invalidResult.success);
  Logger.log('- Message: ' + invalidResult.message);
  Logger.log('- PIN validation: ' + (invalidTestData.isPinValid ? 'VALID' : 'INVALID'));
  
  // Test timestamp format fixing
  Logger.log('\n=== Testing Timestamp Format Fixing ===');
  var timestampFixResult = fixAllTimestampFormats();
  Logger.log('Timestamp fix test completed:');
  Logger.log('- Fixed: ' + timestampFixResult.fixed);
  Logger.log('- Errors: ' + timestampFixResult.errors);
}

/**
 * Test function specifically for diagnosing and fixing the column F issue
 */
function testFixColumnFIssue() {
  Logger.log('=== COLUMN F ISSUE DIAGNOSTIC TEST ===');
  
  try {
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    var sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      Logger.log('Sheet not found: ' + SHEET_NAME);
      return;
    }
    
    // Check current state
    Logger.log('Current sheet state:');
    Logger.log('- Last row: ' + sheet.getLastRow());
    Logger.log('- Last column: ' + sheet.getLastColumn());
    
    if (sheet.getLastColumn() > 0) {
      var allHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      Logger.log('- All current headers: [' + allHeaders.join(', ') + ']');
      
      // Check specifically for column F
      if (sheet.getLastColumn() >= 6) {
        var columnF = sheet.getRange('F1').getValue();
        Logger.log('- Column F value: "' + columnF + '"');
      }
    }
    
    // Run the fix
    Logger.log('\nRunning validateAndFixHeaders...');
    var result = validateAndFixHeaders();
    Logger.log('Fix result: ' + JSON.stringify(result));
    
    // Check state after fix
    Logger.log('\nState after fix:');
    Logger.log('- Last row: ' + sheet.getLastRow());
    Logger.log('- Last column: ' + sheet.getLastColumn());
    
    if (sheet.getLastColumn() > 0) {
      var headersAfter = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 10)).getValues()[0];
      Logger.log('- Headers after fix: [' + headersAfter.join(', ') + ']');
    }
    
    // Check if column F still exists
    if (sheet.getLastColumn() >= 6) {
      var columnFAfter = sheet.getRange('F1').getValue();
      Logger.log('- Column F after fix: "' + columnFAfter + '"');
      Logger.log('WARNING: Column F still exists!');
    } else {
      Logger.log('SUCCESS: Column F has been eliminated');
    }
    
    Logger.log('=== END DIAGNOSTIC TEST ===');
    
  } catch (error) {
    Logger.log('Error in diagnostic test: ' + error.toString());
  }
}

/**
 * Utility function to get recent access requests (for debugging)
 */
function getRecentAccessRequests(limit) {
  limit = limit || 10;
  
  try {
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    
    // Use robust sheet detection to find the access request sheet
    var sheet = getOrCreateAccessRequestSheet(spreadsheet);
    Logger.log('Getting recent requests from sheet: ' + sheet.getName());
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('No data in sheet');
      return [];
    }
    
    var startRow = Math.max(2, lastRow - limit + 1);
    var range = sheet.getRange(startRow, 1, lastRow - startRow + 1, 8);
    var data = range.getValues();
    
    Logger.log('Recent access requests: ' + JSON.stringify(data));
    return data;
    
  } catch (error) {
    Logger.log('Error getting recent access requests: ' + error.toString());
    return [];
  }
}

/**
 * Utility function to clean up old access requests (optional)
 * Keeps only the most recent N records
 */
function cleanupOldRequests(keepRecords) {
  keepRecords = keepRecords || 1000;
  
  try {
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    
    // Use robust sheet detection to find the access request sheet
    var sheet = getOrCreateAccessRequestSheet(spreadsheet);
    Logger.log('Cleaning up old requests from sheet: ' + sheet.getName());
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= keepRecords + 1) {
      Logger.log('No cleanup needed, only ' + (lastRow - 1) + ' records');
      return;
    }
    
    var rowsToDelete = lastRow - keepRecords - 1;
    sheet.deleteRows(2, rowsToDelete);
    
    Logger.log('Cleaned up ' + rowsToDelete + ' old access request records');
    
  } catch (error) {
    Logger.log('Error cleaning up old requests: ' + error.toString());
  }
}

/**
 * AGGRESSIVE FIX: Delete actual columns beyond H to completely eliminate extra column issues
 * This function will physically delete any columns beyond column H
 */
function aggressiveFixColumnStructure() {
  try {
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    
    // Use robust sheet detection to find the access request sheet
    var sheet = getOrCreateAccessRequestSheet(spreadsheet);
    Logger.log('Performing aggressive fix on sheet: ' + sheet.getName());
    
    Logger.log('Starting aggressive column structure fix...');
    Logger.log('Current last column: ' + sheet.getLastColumn());
    
    // Get the current maximum columns in the sheet
    var maxCols = sheet.getMaxColumns();
    Logger.log('Maximum columns in sheet: ' + maxCols);
    
    // Save existing data first (excluding headers)
    var existingData = [];
    if (sheet.getLastRow() > 1) {
      // Only get data from columns A-H, pad with PENDING if needed
      var currentCols = Math.min(sheet.getLastColumn(), 8);
      existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, currentCols).getValues();
      Logger.log('Saved ' + existingData.length + ' rows of existing data');
    }
    
    // Clear the entire sheet
    sheet.clear();
    Logger.log('Cleared entire sheet');
    
    // Set up the correct 8-column structure
    var expectedHeaders = ['Timestamp', 'PIN Number', 'User Email', 'Title', 'URL', 'Media Type', 'Request Status', 'Access Link'];
    sheet.getRange('A1:H1').setValues([expectedHeaders]);
    sheet.getRange('A1:H1').setFontWeight('bold');
    sheet.getRange('A1:H1').setBackground('#f0f0f0');
    
    Logger.log('Set correct headers: ' + expectedHeaders.join(', '));
    
    // Restore existing data
    if (existingData.length > 0) {
      // Ensure we only write 8 columns of data, pad with PENDING if needed
      var dataToRestore = existingData.map(function(row) {
        var restoredRow = row.slice(0, 8); // Only take first 8 columns
        // If row has less than 8 columns, pad with PENDING
        while (restoredRow.length < 8) {
          restoredRow.push('PENDING');
        }
        return restoredRow;
      });
      
      sheet.getRange(2, 1, dataToRestore.length, 8).setValues(dataToRestore);
      Logger.log('Restored ' + dataToRestore.length + ' rows of data');
    }
    
    // Check if there are still extra columns in the sheet structure
    maxCols = sheet.getMaxColumns();
    Logger.log('Sheet max columns after reset: ' + maxCols);
    
    // If there are more than 8 maximum columns, delete them
    if (maxCols > 8) {
      var colsToDelete = maxCols - 8;
      Logger.log('Deleting ' + colsToDelete + ' extra columns from sheet structure');
      
      for (var i = 0; i < colsToDelete; i++) {
        if (sheet.getMaxColumns() > 8) {
          sheet.deleteColumn(9); // Always delete column 9 (I)
        }
      }
    }
    
    // Auto-resize the 8 columns for better display
    sheet.autoResizeColumns(1, 8);
    
    Logger.log('Aggressive fix completed. Final column count: ' + sheet.getLastColumn());
    Logger.log('Final max columns: ' + sheet.getMaxColumns());
    
    return { success: true, message: 'Column structure aggressively fixed and data restored' };
    
  } catch (error) {
    Logger.log('Error in aggressive column fix: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

/**
 * ROBUST SHEET DETECTION FUNCTIONS
 * These functions automatically detect sheets regardless of naming conventions
 */

/**
 * Intelligently finds the access request sheet by checking multiple possible names
 * @param {Spreadsheet} spreadsheet - The spreadsheet to search in
 * @returns {Sheet|null} - The found sheet or null
 */
function findAccessRequestSheet(spreadsheet) {
  // List of possible sheet names for access requests (case-insensitive)
  var possibleNames = [
    'AccessRequests',
    'Access Requests', 
    'Access',
    'Requests',
    'Sheet1',
    'access',
    'requests',
    'accessrequests',
    'access_requests',
    'URL Requests',
    'url_requests',
    'urlrequests',
    'Logs',
    'logs',
    'Activity'
  ];
  
  var sheets = spreadsheet.getSheets();
  Logger.log('Available sheets: ' + sheets.map(s => s.getName()).join(', '));
  
  // First try exact matches
  for (var i = 0; i < possibleNames.length; i++) {
    var sheet = spreadsheet.getSheetByName(possibleNames[i]);
    if (sheet) {
      Logger.log('Found access request sheet: ' + possibleNames[i]);
      return sheet;
    }
  }
  
  // Then try case-insensitive matches
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName().toLowerCase();
    for (var j = 0; j < possibleNames.length; j++) {
      if (sheetName === possibleNames[j].toLowerCase()) {
        Logger.log('Found access request sheet (case-insensitive): ' + sheets[i].getName());
        return sheets[i];
      }
    }
  }
  
  // Check if any sheet contains access request headers
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    try {
      if (sheet.getLastRow() >= 1 && sheet.getLastColumn() >= 3) {
        var headers = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 8)).getValues()[0];
        var headerText = headers.join('|').toLowerCase();
        
        // Look for access request indicators in headers
        if (headerText.includes('timestamp') && 
            (headerText.includes('url') || headerText.includes('email') || headerText.includes('pin'))) {
          Logger.log('Found access request sheet by headers: ' + sheet.getName());
          return sheet;
        }
      }
    } catch (e) {
      // Skip sheets that can't be read
      continue;
    }
  }
  
  Logger.log('No access request sheet found, will create: ' + PREFERRED_ACCESS_SHEET_NAME);
  return null;
}

/**
 * Gets or creates the access request sheet with proper setup
 * @param {Spreadsheet} spreadsheet - The spreadsheet to work with
 * @returns {Sheet} - The access request sheet
 */
function getOrCreateAccessRequestSheet(spreadsheet) {
  var sheet = findAccessRequestSheet(spreadsheet);
  
  if (!sheet) {
    // Create new sheet with preferred name
    sheet = spreadsheet.insertSheet(PREFERRED_ACCESS_SHEET_NAME);
    Logger.log('Created new access request sheet: ' + PREFERRED_ACCESS_SHEET_NAME);
    
    // Set up headers immediately
    var expectedHeaders = ['Timestamp', 'PIN Number', 'User Email', 'Title', 'URL', 'Request Status', 'Media Type', 'Access Link'];
    sheet.getRange('A1:H1').setValues([expectedHeaders]);
    sheet.getRange('A1:H1').setFontWeight('bold');
    sheet.getRange('A1:H1').setBackground('#f0f0f0');
    Logger.log('Set up headers for new access request sheet');
  }
  
  return sheet;
}

/**
 * Intelligently finds the PIN validation sheet by checking multiple possible names
 * @param {Spreadsheet} spreadsheet - The spreadsheet to search in (may be different from access request sheet)
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
    'Codes',
    'codes',
    'Auth',
    'auth',
    'Authentication'
  ];
  
  var sheets = spreadsheet.getSheets();
  Logger.log('Available PIN sheets: ' + sheets.map(s => s.getName()).join(', '));
  
  // First try exact matches
  for (var i = 0; i < possibleNames.length; i++) {
    var sheet = spreadsheet.getSheetByName(possibleNames[i]);
    if (sheet) {
      // Verify this looks like a PIN sheet
      if (isPinSheet(sheet)) {
        Logger.log('Found PIN sheet: ' + possibleNames[i]);
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
  
  // Look for sheets that contain PIN-like data
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
    
    // Check for PIN-like headers
    if (sheet.getLastColumn() >= 1) {
      var possibleHeaders = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 3)).getValues()[0];
      var headerText = possibleHeaders.join('|').toLowerCase();
      
      if (headerText.includes('pin') || headerText.includes('password') || 
          headerText.includes('code') || headerText.includes('auth')) {
        return true;
      }
    }
    
    // Check for PIN-like data in column A (numeric codes OR text passwords)
    if (sheet.getLastRow() >= 2) {
      var dataRange = sheet.getRange(Math.min(2, sheet.getLastRow()), 1, Math.min(10, sheet.getLastRow() - 1), 1);
      var data = dataRange.getValues();
      
      var pinLikeCount = 0;
      for (var i = 0; i < data.length; i++) {
        var value = data[i][0];
        if (value && typeof value === 'number' && value >= 1000 && value <= 999999) {
          pinLikeCount++; // Numeric PIN
        } else if (value && typeof value === 'string') {
          var strValue = value.toString().trim();
          if (/^\d{4,6}$/.test(strValue)) {
            pinLikeCount++; // Numeric PIN as string
          } else if (strValue.length >= 3) {
            // Accept any text password 3+ characters long
            pinLikeCount++;
          }
        }
      }
      
      // If most values look like PINs or passwords, consider this a PIN sheet
      if (pinLikeCount >= data.length * 0.5) {
        return true;
      }
    }
    
    // If sheet has any non-empty data, give it a chance to be a PIN sheet
    if (sheet.getLastRow() >= 1) {
      var firstValue = sheet.getRange(1, 1).getValue();
      if (firstValue) {
        return true; // Be more permissive - let validatePin() handle the actual validation
      }
    }
    
    return false;
  } catch (e) {
    return false;
  }
}

/**
 * Validates a PIN against the PIN sheet with robust sheet detection
 * @param {string} pin - The PIN to validate
 * @returns {boolean} - Whether the PIN is valid
 */
function validatePin(pin) {
  if (!pin) {
    Logger.log('No PIN provided for validation');
    return false;
  }
  
  try {
    // Open the PIN spreadsheet (may be same as access request spreadsheet)
    var pinSpreadsheet = SpreadsheetApp.openById(PIN_SHEET_ID);
    Logger.log('Opened PIN spreadsheet for validation with ID: ' + PIN_SHEET_ID);
    
    // Find the PIN sheet using robust detection
    var pinSheet = findPinSheet(pinSpreadsheet);
    
    if (!pinSheet) {
      Logger.log('No PIN sheet found for validation');
      return false;
    }
    
    Logger.log('Using PIN sheet: ' + pinSheet.getName());
    
    var lastRow = pinSheet.getLastRow();
    if (lastRow < 1) {
      Logger.log('PIN sheet is empty');
      return false;
    }
    
    // Determine the correct range for PIN data
    var startRow = 1;
    var pinColumn = 1;
    
    // Check if row 1 contains headers
    if (lastRow >= 1) {
      var firstRowValue = pinSheet.getRange(1, 1).getValue();
      if (firstRowValue && typeof firstRowValue === 'string') {
        var headerText = firstRowValue.toString().toLowerCase();
        if (headerText.includes('pin') || headerText.includes('password') || 
            headerText.includes('code') || headerText.includes('auth') ||
            headerText === 'pin' || headerText === 'pins') {
          startRow = 2; // Skip header row
          Logger.log('Detected header row, starting PIN search from row 2');
        }
      }
    }
    
    if (lastRow < startRow) {
      Logger.log('No PIN data rows available');
      return false;
    }
    
    // Get PIN data
    var pinRange = pinSheet.getRange(startRow, pinColumn, lastRow - startRow + 1, 1);
    var pinList = pinRange.getDisplayValues();
    
    Logger.log('Retrieved PIN list for validation, checking ' + pinList.length + ' entries');
    
    // Convert input PIN to string for comparison
    var inputPin = pin.toString().trim();
    Logger.log('Validating PIN: [' + inputPin.length + ' characters]');
    
    // Compare with each PIN in the list
    for (var i = 0; i < pinList.length; i++) {
      var storedPin = pinList[i][0];
      if (storedPin) {
        storedPin = storedPin.toString().trim();
        if (storedPin === inputPin) {
          Logger.log('Valid PIN found at row ' + (startRow + i));
          return true;
        }
      }
    }
    
    Logger.log('PIN validation failed - no match found');
    return false;
    
  } catch (error) {
    Logger.log('Error during PIN validation: ' + error.toString());
    return false;
  }
}

/**
 * Debug function to troubleshoot PIN validation issues
 * Run this function in the Google Apps Script console to see what's happening
 */
function debugPinValidation() {
  Logger.log('=== PIN VALIDATION DEBUG ===');
  
  try {
    var pinSpreadsheet = SpreadsheetApp.openById(PIN_SHEET_ID);
    Logger.log('✅ Successfully opened PIN spreadsheet: ' + PIN_SHEET_ID);
    
    // List all sheets
    var sheets = pinSpreadsheet.getSheets();
    Logger.log('📋 Available sheets in PIN spreadsheet:');
    sheets.forEach(function(sheet, index) {
      Logger.log('  ' + (index + 1) + '. "' + sheet.getName() + '" (Rows: ' + sheet.getLastRow() + ', Cols: ' + sheet.getLastColumn() + ')');
    });
    
    // Try to find PIN sheet
    var pinSheet = findPinSheet(pinSpreadsheet);
    
    if (!pinSheet) {
      Logger.log('❌ No PIN sheet found using findPinSheet()');
      
      // Check each sheet manually
      Logger.log('\n🔍 Manual sheet inspection:');
      sheets.forEach(function(sheet) {
        Logger.log('\n--- Analyzing sheet: "' + sheet.getName() + '" ---');
        Logger.log('isPinSheet result: ' + isPinSheet(sheet));
        
        if (sheet.getLastRow() >= 1) {
          // Show first few rows
          var maxRows = Math.min(5, sheet.getLastRow());
          var maxCols = Math.min(3, sheet.getLastColumn());
          if (maxRows > 0 && maxCols > 0) {
            var data = sheet.getRange(1, 1, maxRows, maxCols).getValues();
            Logger.log('First ' + maxRows + ' rows:');
            data.forEach(function(row, i) {
              Logger.log('  Row ' + (i + 1) + ': [' + row.join(', ') + ']');
            });
          }
        }
      });
      
      return;
    }
    
    Logger.log('✅ Found PIN sheet: "' + pinSheet.getName() + '"');
    
    // Show PIN sheet contents
    var lastRow = pinSheet.getLastRow();
    Logger.log('📊 PIN sheet has ' + lastRow + ' rows');
    
    if (lastRow >= 1) {
      var maxRows = Math.min(10, lastRow);
      var data = pinSheet.getRange(1, 1, maxRows, 1).getValues();
      
      Logger.log('\n📝 PIN sheet contents (first ' + maxRows + ' rows):');
      data.forEach(function(row, i) {
        var value = row[0];
        var type = typeof value;
        var display = value ? value.toString() : '(empty)';
        Logger.log('  Row ' + (i + 1) + ': "' + display + '" (type: ' + type + ')');
      });
    }
    
    // Test common passwords
    var testPins = ['password', 'Password', '1234', '5678'];
    Logger.log('\n🔑 Testing PIN validation:');
    
    testPins.forEach(function(pin) {
      Logger.log('Testing "' + pin + '": ' + validatePin(pin));
    });
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
  }
  
  Logger.log('\n=== DEBUG COMPLETE ===');
}

/**
 * Comprehensive test function to demonstrate robust sheet detection
 * This function tests the script's ability to find sheets regardless of naming
 */
function testRobustSheetDetection() {
  Logger.log('=== ROBUST SHEET DETECTION TEST ===');
  
  try {
    // Test access request sheet detection
    Logger.log('\n--- Testing Access Request Sheet Detection ---');
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('Available sheets: ' + spreadsheet.getSheets().map(s => s.getName()).join(', '));
    
    var accessSheet = findAccessRequestSheet(spreadsheet);
    if (accessSheet) {
      Logger.log('✅ Found access request sheet: ' + accessSheet.getName());
    } else {
      Logger.log('❌ No access request sheet found - will create: ' + PREFERRED_ACCESS_SHEET_NAME);
    }
    
    // Test getting or creating access sheet
    var finalAccessSheet = getOrCreateAccessRequestSheet(spreadsheet);
    Logger.log('✅ Final access sheet: ' + finalAccessSheet.getName());
    
    // Test PIN sheet detection
    Logger.log('\n--- Testing PIN Sheet Detection ---');
    var pinSpreadsheet = SpreadsheetApp.openById(PIN_SHEET_ID);
    Logger.log('Available PIN sheets: ' + pinSpreadsheet.getSheets().map(s => s.getName()).join(', '));
    
    var pinSheet = findPinSheet(pinSpreadsheet);
    if (pinSheet) {
      Logger.log('✅ Found PIN sheet: ' + pinSheet.getName());
      
      // Test PIN validation
      Logger.log('\n--- Testing PIN Validation ---');
      if (pinSheet.getLastRow() >= 2) {
        // Get a sample PIN from the sheet for testing
        var samplePin = pinSheet.getRange(2, 1).getValue();
        if (samplePin) {
          samplePin = samplePin.toString().trim();
          Logger.log('Testing with sample PIN from sheet: [' + samplePin.length + ' characters]');
          var isValid = validatePin(samplePin);
          Logger.log('Sample PIN validation result: ' + (isValid ? 'VALID ✅' : 'INVALID ❌'));
        }
      }
      
      // Test with known invalid PIN
      var invalidTest = validatePin('000000');
      Logger.log('Invalid PIN test (000000): ' + (invalidTest ? 'VALID ✅' : 'INVALID ❌'));
      
    } else {
      Logger.log('❌ No PIN sheet found');
      Logger.log('Available sheets for PIN validation:');
      pinSpreadsheet.getSheets().forEach(function(sheet) {
        Logger.log('  - ' + sheet.getName() + ' (appears to be PIN sheet: ' + (isPinSheet(sheet) ? 'YES' : 'NO') + ')');
      });
    }
    
    Logger.log('\n=== DETECTION TEST COMPLETE ===');
    
  } catch (error) {
    Logger.log('Error in robust sheet detection test: ' + error.toString());
  }
}

/**
 * Test function specifically for different sheet naming scenarios
 */
function testSheetNamingScenarios() {
  Logger.log('=== SHEET NAMING SCENARIOS TEST ===');
  
  try {
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    var sheets = spreadsheet.getSheets();
    
    Logger.log('Testing each available sheet for compatibility:');
    
    sheets.forEach(function(sheet, index) {
      Logger.log('\n--- Sheet ' + (index + 1) + ': "' + sheet.getName() + '" ---');
      
      // Check if it could be an access request sheet
      var accessSheetResult = (findAccessRequestSheet(spreadsheet) && 
                              findAccessRequestSheet(spreadsheet).getName() === sheet.getName());
      Logger.log('Access request sheet candidate: ' + (accessSheetResult ? 'YES ✅' : 'NO'));
      
      // Check if it could be a PIN sheet
      var pinSheetResult = isPinSheet(sheet);
      Logger.log('PIN sheet candidate: ' + (pinSheetResult ? 'YES ✅' : 'NO'));
      
      // Show basic sheet info
      Logger.log('- Rows: ' + sheet.getLastRow());
      Logger.log('- Columns: ' + sheet.getLastColumn());
      
      if (sheet.getLastRow() >= 1 && sheet.getLastColumn() >= 1) {
        try {
          var firstRow = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 5)).getValues()[0];
          Logger.log('- First row sample: [' + firstRow.slice(0, 3).join(', ') + '...]');
        } catch (e) {
          Logger.log('- Could not read first row');
        }
      }
    });
    
    Logger.log('\n=== NAMING SCENARIOS TEST COMPLETE ===');
    
  } catch (error) {
    Logger.log('Error in sheet naming scenarios test: ' + error.toString());
  }
}

/**
 * Automatically detects and fixes timestamp formatting issues in the sheet
 * This function will convert any ISO timestamps or incorrectly formatted timestamps 
 * to the human-readable format
 * @param {Sheet} sheet - The sheet to check and fix
 * @returns {Object} - Result with details of what was fixed
 */
function autoFixTimestampFormats(sheet) {
  try {
    var result = {
      fixed: 0,
      errors: 0,
      details: []
    };
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('No data rows to check for timestamp format');
      return result;
    }
    
    Logger.log('Checking timestamp formats in ' + (lastRow - 1) + ' rows...');
    
    // Get all timestamp values from column A (excluding header)
    var timestampRange = sheet.getRange(2, 1, lastRow - 1, 1);
    var timestampValues = timestampRange.getValues();
    var needsUpdate = false;
    
    // Check and fix each timestamp
    for (var i = 0; i < timestampValues.length; i++) {
      var currentValue = timestampValues[i][0];
      var rowNumber = i + 2; // +2 because we start from row 2 and array is 0-indexed
      
      try {
        var fixedTimestamp = detectAndFixTimestamp(currentValue);
        
        if (fixedTimestamp !== currentValue) {
          timestampValues[i][0] = fixedTimestamp;
          needsUpdate = true;
          result.fixed++;
          result.details.push('Row ' + rowNumber + ': Fixed "' + currentValue + '" → "' + fixedTimestamp + '"');
          Logger.log('Fixed timestamp in row ' + rowNumber + ': "' + currentValue + '" → "' + fixedTimestamp + '"');
        }
      } catch (error) {
        result.errors++;
        result.details.push('Row ' + rowNumber + ': Error fixing "' + currentValue + '" - ' + error.toString());
        Logger.log('Error fixing timestamp in row ' + rowNumber + ': ' + error.toString());
      }
    }
    
    // Update the sheet if any changes were made
    if (needsUpdate) {
      timestampRange.setValues(timestampValues);
      Logger.log('Updated ' + result.fixed + ' timestamp formats in the sheet');
      
      // Set proper text format for the timestamp column to prevent auto-conversion
      timestampRange.setNumberFormat('@'); // @ means text format
    }
    
    return result;
    
  } catch (error) {
    Logger.log('Error in autoFixTimestampFormats: ' + error.toString());
    return {
      fixed: 0,
      errors: 1,
      details: ['Error in autoFixTimestampFormats: ' + error.toString()]
    };
  }
}

/**
 * Detects various timestamp formats and converts them to human-readable format
 * @param {*} value - The timestamp value to check and fix
 * @returns {string} - The properly formatted timestamp
 */
function detectAndFixTimestamp(value) {
  if (!value) {
    return createReadableTimestamp(); // Return current readable timestamp if empty
  }
  
  var stringValue = value.toString().trim();
  
  // If it's already in our desired format, keep it
  if (isReadableTimestampFormat(stringValue)) {
    return stringValue;
  }
  
  var date;
  
  try {
    // Try to parse various timestamp formats
    
    // ISO 8601 format (e.g., "2025-06-06T17:22:23.598Z")
    if (stringValue.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{3})?Z?$/)) {
      date = new Date(stringValue);
      Logger.log('Detected ISO 8601 format: ' + stringValue);
    }
    // Date object (from Google Sheets date cells)
    else if (value instanceof Date) {
      date = value;
      Logger.log('Detected Date object: ' + stringValue);
    }
    // Unix timestamp (numeric)
    else if (!isNaN(value) && value > 1000000000) {
      date = new Date(Number(value) * (value.toString().length === 10 ? 1000 : 1)); // Handle both seconds and milliseconds
      Logger.log('Detected Unix timestamp: ' + stringValue);
    }
    // Other date formats that JavaScript can parse
    else {
      date = new Date(stringValue);
      Logger.log('Attempting to parse as generic date: ' + stringValue);
    }
    
    // Check if the date is valid
    if (isNaN(date.getTime())) {
      Logger.log('Could not parse timestamp, keeping original: ' + stringValue);
      return stringValue; // Keep original if we can't parse it
    }
    
    // Convert to our readable format
    var readableFormat = Utilities.formatDate(date, Session.getScriptTimeZone(), "EEEE, MMM dd, yyyy hh:mm a");
    Logger.log('Converted to readable format: ' + readableFormat);
    return readableFormat;
    
  } catch (error) {
    Logger.log('Error parsing timestamp "' + stringValue + '": ' + error.toString());
    return stringValue; // Keep original if there's an error
  }
}

/**
 * Checks if a string is already in the desired readable timestamp format
 * @param {string} value - The string to check
 * @returns {boolean} - Whether it's already in correct format
 */
function isReadableTimestampFormat(value) {
  // Pattern for "Friday, Jun 06, 2025 10:30 AM" format
  var pattern = /^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday), (Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{2}, \d{4} \d{1,2}:\d{2} (AM|PM)$/;
  return pattern.test(value);
}

/**
 * Standalone function to fix all timestamp formats in the sheet
 * This can be run manually to fix existing data
 */
function fixAllTimestampFormats() {
  try {
    Logger.log('=== FIXING ALL TIMESTAMP FORMATS ===');
    
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    var sheet = getOrCreateAccessRequestSheet(spreadsheet);
    
    Logger.log('Fixing timestamps in sheet: ' + sheet.getName());
    
    var result = autoFixTimestampFormats(sheet);
    
    Logger.log('=== TIMESTAMP FIX RESULTS ===');
    Logger.log('- Fixed: ' + result.fixed + ' timestamps');
    Logger.log('- Errors: ' + result.errors);
    Logger.log('- Details:');
    
    for (var i = 0; i < result.details.length; i++) {
      Logger.log('  ' + result.details[i]);
    }
    
    if (result.fixed > 0) {
      Logger.log('SUCCESS: Fixed ' + result.fixed + ' timestamp formats to human-readable format');
    } else if (result.errors === 0) {
      Logger.log('SUCCESS: All timestamps are already in correct format');
    } else {
      Logger.log('WARNING: Had ' + result.errors + ' errors while fixing timestamps');
    }
    
    return result;
    
  } catch (error) {
    Logger.log('ERROR in fixAllTimestampFormats: ' + error.toString());
    return {
      fixed: 0,
      errors: 1,
      details: ['Error: ' + error.toString()]
    };
  }
}

/**
 * Test function specifically for timestamp format detection and conversion
 */
function testTimestampFormatDetection() {
  Logger.log('=== TESTING TIMESTAMP FORMAT DETECTION ===');
  
  var testTimestamps = [
    '2025-06-06T17:22:23.598Z',           // ISO format
    '2025-06-06T17:22:23Z',               // ISO without milliseconds
    new Date(),                           // Date object
    'Friday, Jun 06, 2025 05:22 PM',      // Already correct format
    '6/6/2025 5:22:23 PM',                // US format
    '2025-06-06 17:22:23',                // SQL format
    1717693343,                           // Unix timestamp (seconds)
    1717693343598,                        // Unix timestamp (milliseconds)
    'Jun 6, 2025 5:22 PM',                // Partial format
    '',                                   // Empty
    null,                                 // Null
    'invalid date string'                 // Invalid
  ];
  
  for (var i = 0; i < testTimestamps.length; i++) {
    var testValue = testTimestamps[i];
    try {
      var result = detectAndFixTimestamp(testValue);
      Logger.log('Test ' + (i + 1) + ':');
      Logger.log('  Input:  ' + (testValue === null ? 'null' : testValue.toString()));
      Logger.log('  Output: ' + result);
      Logger.log('  Valid:  ' + isReadableTimestampFormat(result));
      Logger.log('');
    } catch (error) {
      Logger.log('Test ' + (i + 1) + ' ERROR: ' + error.toString());
    }
  }
  
  Logger.log('=== TIMESTAMP FORMAT DETECTION TEST COMPLETE ===');
}
