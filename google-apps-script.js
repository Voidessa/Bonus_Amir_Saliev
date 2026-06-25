/**
 * Google Apps Script Web App Connector for Questionnaire CRM DB
 *
 * How to deploy:
 * 1. Open Google Sheets.
 * 2. Click Extensions > Apps Script.
 * 3. Delete any code in the editor and paste this code.
 * 4. Click Save.
 * 5. Click Deploy > New deployment.
 * 6. Select type "Web app".
 * 7. Set Description: "Blogger CRM Database API"
 * 8. Set Execute as: "Me"
 * 9. Set Who has access: "Anyone"
 * 10. Click Deploy, authorize permissions, and copy the "Web app URL".
 * 11. Paste this URL into app.js as SHEET_SCRIPT_URL.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // 10 seconds lock to prevent concurrent writing issues
  
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Self-initialize headers if the sheet is empty
    if (sheet.getLastRow() === 0) {
      var headers = [
        "Timestamp", 
        "Full Name", 
        "Phone Number", 
        "Telegram Username", 
        "Age & Profile", 
        "Motivation / Goal", 
        "Previous Courses", 
        "Income Level", 
        "Readiness to Study"
      ];
      
      sheet.appendRow(headers);
      
      // Format headers with brand background (Gold) and Bold white text
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setFontColor("#FFFFFF");
      headerRange.setBackground("#D4AF37"); // Blogger brand Gold color
      headerRange.setHorizontalAlignment("center");
      headerRange.setFontSize(11);
      sheet.setFrozenRows(1);
    }
    
    // Parse parameters
    var params = e.parameter;
    
    // Extract values. If values already start with single quote, Google Sheets treats it as plain text and strips the visible quote.
    // However, if some parameters are undefined, set default empty string.
    var rowData = [
      params.timestamp || new Date().toISOString(),
      params.name || "",
      params.phone || "",
      params.telegram || "",
      params.age_profile || "",
      params.motivation || "",
      params.courses || "",
      params.income || "",
      params.readiness || ""
    ];
    
    // Append the row
    sheet.appendRow(rowData);
    
    // Auto-resize columns to keep it organized
    var lastColumn = sheet.getLastColumn();
    for (var col = 1; col <= lastColumn; col++) {
      sheet.autoResizeColumn(col);
    }
    
    // Return CORS-friendly success output
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Amir Saliev Questionnaire API is active. Send POST requests to submit data.")
    .setMimeType(ContentService.MimeType.TEXT);
}
