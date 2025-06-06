function doGet(e) {
  return handleCors();
}

function doPost(e) {
  try {
    // Handle CORS preflight requests
    if (e.parameter.method === 'OPTIONS' || !e.postData) {
      Logger.log('Received OPTIONS or empty postData request. Returning CORS headers.');
      return handleCors();
    }
    // Log the incoming request for debugging
    Logger.log('Received POST request for URL logging. Post data: ' + e.postData.contents);
    // Parse the incoming JSON payload from the extension
    var data = JSON.parse(e.postData.contents);
    var url = data.url || 'N/A';
    var title = data.title || 'Untitled Page';
    // Parse and format timestamp to PST
    var rawTimestamp = data.timestamp || new Date().toISOString();
    var timestampDate = new Date(rawTimestamp);
    var timestamp = Utilities.formatDate(timestampDate, "America/Los_Angeles", "MM/dd/yyyy h:mm:ss a"); // Replace with your preferred TimeZone and Date & Time Format
    Logger.log('Raw timestamp (UTC): ' + rawTimestamp + ', Formatted timestamp (PST): ' + timestamp);
    var userEmail = data.userEmail || 'Unknown';
    var userId = data.userId || 'Unknown';
    var pin = data.pin || 'Not Provided';
    Logger.log('Parsed data - Timestamp: ' + timestamp + ', Email: ' + userEmail + ', Title: ' + title + ', URL: ' + url + ', PIN: ' + pin);
    // Open the Google Spreadsheet for URLs (Request Form)
    var spreadsheet = SpreadsheetApp.openById('YOUR_SHEET_ID_HERE_FROM_ACCESS_SHEET'); // Replace with your Google Sheet ID for Request Form
    Logger.log('Opened spreadsheet for URLs with ID: YOUR_SHEET_ID_HERE_FROM_ACCESS_SHEET');
    var sheet = spreadsheet.getSheetByName('Sheet1'); // Adjust sheet name if different
    if (!sheet) {
      Logger.log('Error: Sheet named "Sheet1" not found.');
      throw new Error('Sheet named "Sheet1" not found.');
    }
    // Check if headers exist in the first row and match the expected headers; if not, insert a new row at the top
    var firstRowRange = sheet.getRange(1, 1, 1, 6);
    var firstRowValues = firstRowRange.getValues()[0];
    var expectedHeaders = ['timeStamp', 'userEmail', 'websiteTitle', 'websiteRequested', 'pinNumber', 'pinStatus'];
    var headersMatch = firstRowValues.length === expectedHeaders.length &&
                       firstRowValues.every((value, index) =>
                         value.toString().trim().toLowerCase() === expectedHeaders[index].toLowerCase());
    if (!headersMatch) {
      Logger.log('Headers missing or do not match expected values. Current first row: ' + JSON.stringify(firstRowValues));
      Logger.log('Inserting new row at the top with headers: ' + expectedHeaders.join(', '));
      sheet.insertRowBefore(1); // Insert a new row at row 1, shifting existing data down
      sheet.getRange(1, 1, 1, 6).setValues([expectedHeaders]); // Set headers in the new first row
      Logger.log('Headers successfully inserted in the new first row.');
    } else {
      Logger.log('Headers already exist and match expected values: ' + firstRowValues.join(', '));
    }
    // Validate PIN by accessing the PIN spreadsheet
    var isValidPin = false;
    try {
      var pinSpreadsheet = SpreadsheetApp.openById('YOUR_SHEET_ID_HERE_FROM_PIN_ACCESS_SHEET'); // Replace with your Google Sheet ID for PINs
      Logger.log('Opened PIN spreadsheet for validation with ID: YOUR_SHEET_ID_HERE_FROM_PIN_ACCESS_SHEET');
      var pinSheet = pinSpreadsheet.getSheetByName('Sheet1'); // Case-sensitive, adjust if different
      if (pinSheet) {
        Logger.log('Successfully accessed sheet named "Sheet1".');
        var lastPinRow = pinSheet.getLastRow();
        Logger.log('Last row in PIN sheet: ' + lastPinRow);
        // Adjust range if header is in row 1 and data starts at row 2
        var pinRange = lastPinRow > 1 ? pinSheet.getRange('A2:A' + lastPinRow) : null;
        var pinList = pinRange ? pinRange.getDisplayValues() : []; // Use getDisplayValues to avoid formatting issues
        Logger.log('Retrieved PIN list for validation (display values): ' + JSON.stringify(pinList));
        // Convert input PIN to string for comparison
        var inputPin = pin.toString().trim();
        Logger.log('Input PIN (trimmed as string): "' + inputPin + '" (length: ' + inputPin.length + ')');
        for (var i = 0; i < pinList.length; i++) {
          var storedPin = pinList[i][0].toString().trim();
          Logger.log('Comparing with stored PIN at row ' + (i + 2) + ': "' + storedPin + '" (length: ' + storedPin.length + ', type: ' + typeof pinList[i][0] + ')');
          if (storedPin && storedPin === inputPin) {
            isValidPin = true;
            Logger.log('PIN match found at row ' + (i + 2) + ': "' + storedPin + '" matches input "' + inputPin + '"');
            break;
          } else {
            Logger.log('No match: stored "' + storedPin + '" ≠ input "' + inputPin + '"');
          }
        }
        // Hardcoded check for debugging specific PINs
        if (inputPin === "9789") {
          Logger.log('Hardcoded check: Input PIN is one of the test PINs (1234, 1111, 9789). Should be valid if in sheet.');
        }
      } else {
        Logger.log('Error: PIN sheet named "Sheet1" not found. Available sheets: ' + pinSpreadsheet.getSheets().map(s => s.getName()).join(', '));
      }
    } catch (pinError) {
      Logger.log('Error accessing PIN spreadsheet: ' + pinError.toString());
    }
    Logger.log('PIN validation result during URL logging: ' + (isValidPin ? 'Valid' : 'Invalid'));
    // Check for duplicates within the last 5 minutes (only for valid PINs)
    var lastRow = sheet.getLastRow();
    Logger.log('Last row in URL sheet: ' + lastRow);
    var lastRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 5).getValues() : [];
    var currentTime = new Date(timestampDate).getTime(); // Use the parsed Date object for consistency
    var cooldownMs = 5 * 60 * 1000; // 5 minutes in milliseconds
    Logger.log('Checking for duplicates within cooldown period of 5 minutes.');
    var isDuplicate = false;
    if (isValidPin) {
      for (var i = lastRows.length - 1; i >= 0; i--) {
        var rowUrl = lastRows[i][0];
        var rowTimeStr = lastRows[i][1];
        if (!rowTimeStr) continue;
        var rowTime = new Date(rowTimeStr).getTime();
        if (isNaN(rowTime)) {
          Logger.log('Invalid timestamp in row ' + (i + 2) + ': ' + rowTimeStr);
          continue;
        }
        if (currentTime - rowTime > cooldownMs) break;
        if (rowUrl === url) {
          isDuplicate = true;
          Logger.log('Duplicate URL found in row ' + (i + 2) + ': ' + url);
          break;
        }
      }
    }
    // Log all attempts (valid and invalid PINs, with status)
    var status = isValidPin ? (isDuplicate ? 'Valid PIN' : 'Valid PIN') : 'Invalid PIN';
    sheet.appendRow([timestamp, userEmail, title, url, pin, status]);
    Logger.log('Appended new row with data: ' + [timestamp, userEmail, title, url, pin, status].join(', '));
    // Return response based on PIN validity and duplicate check
    if (!isValidPin) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'failure', message: 'Invalid PIN/Password' }))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeader('Access-Control-Allow-Origin', '*')
        .setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
        .setHeader('Access-Control-Allow-Headers', 'Content-Type');
    } else if (isDuplicate) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'duplicate', message: 'URL already requested within last 1 minute.' }))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeader('Access-Control-Allow-Origin', '*')
        .setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
        .setHeader('Access-Control-Allow-Headers', 'Content-Type');
    } else {
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeader('Access-Control-Allow-Origin', '*')
        .setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
        .setHeader('Access-Control-Allow-Headers', 'Content-Type');
    }
  } catch (error) {
    // Log the error for debugging
    Logger.log('Error in URL logging script: ' + error.toString());
    // Return error response if something goes wrong
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*')
      .setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
      .setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
}

// Helper function to handle CORS preflight requests
function handleCors() {
  Logger.log('Handling CORS preflight request for URL logging.');
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type')
    .setStatusCode(204); // No Content for preflight
}