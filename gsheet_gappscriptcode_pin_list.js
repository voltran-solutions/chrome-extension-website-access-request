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
    Logger.log('Received POST request. Post data: ' + e.postData.contents);

    // Parse the incoming JSON payload from the extension
    var data = JSON.parse(e.postData.contents);
    var pin = data.pin || '';
    Logger.log('Parsed PIN from payload: ' + pin);

    // Open the Google Spreadsheet by ID
    var spreadsheet = SpreadsheetApp.openById('YOUR_SHEET_ID_HERE_FROM_PIN_ACCESS_SHEET'); // Replace with your Google Sheet ID for PINs
    Logger.log('Opened spreadsheet for PINs with ID: YOUR_SHEET_ID_HERE_FROM_PIN_ACCESS_SHEET');
    var sheet = spreadsheet.getSheetByName('PINs'); // Adjust sheet name if different
    if (!sheet) {
      Logger.log('Error: Sheet named "PINs" not found.');
      throw new Error('Sheet named "PINs" not found.');
    }
    var lastRow = sheet.getLastRow();
    Logger.log('Last row in PINs sheet: ' + lastRow);
    var pinRange = lastRow > 1 ? sheet.getRange('A2:A' + lastRow) : null;
    var pinList = pinRange ? pinRange.getValues() : [];
    Logger.log('Retrieved PIN list (raw values): ' + JSON.stringify(pinList));

    // Check if the provided PIN exists in the list
    var isValid = false;
    var inputPin = pin.toString().trim();
    Logger.log('Input PIN (trimmed as string): "' + inputPin + '"');
    for (var i = 0; i < pinList.length; i++) {
      var storedPin = pinList[i][0].toString().trim();
      Logger.log('Comparing with stored PIN at row ' + (i + 2) + ': "' + storedPin + '" (type: ' + typeof pinList[i][0] + ')');
      if (storedPin && storedPin === inputPin) {
        isValid = true;
        Logger.log('PIN match found at row ' + (i + 2) + ': "' + storedPin + '" matches input "' + inputPin + '"');
        break;
      } else {
        Logger.log('No match: stored "' + storedPin + '" ≠ input "' + inputPin + '"');
      }
    }
    Logger.log('PIN validation result: ' + (isValid ? 'Valid' : 'Invalid'));

    // Return validation result
    return ContentService.createTextOutput(JSON.stringify({ status: isValid ? 'success' : 'failure', message: isValid ? 'PIN valid' : 'Invalid PIN/Password' }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*')
      .setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
      .setHeader('Access-Control-Allow-Headers', 'Content-Type');
  } catch (error) {
    // Log the error for debugging
    Logger.log('Error in PIN validation script: ' + error.toString());
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
  Logger.log('Handling CORS preflight request.');
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type')
    .setStatusCode(204); // No Content for preflight
}
