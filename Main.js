// Main function to add emails to Google Sheet
function addEmailsToSheetRechtstreeks() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = sheets.getSheetByName("ontvangenMails");
  var existingEmailIds = new Set(sourceSheet.getRange(1, 4, sourceSheet.getLastRow()).getValues().flat());

  var targetSpreadsheet = SpreadsheetApp.openById("Your-Spreadsheet-ID");
  var targetSheet = targetSpreadsheet.getSheetByName("_inkomendeVragen");

  var nameLookupSheet = targetSpreadsheet.getSheetByName("vanWie");
  var nameLookupRange = nameLookupSheet.getRange("A:B").getValues();
  var nameLookupMap = new Map(nameLookupRange.map(([name, email]) => [email, name]));

  var setupSheet = sheets.getSheetByName("setup");
  var allowedEmails = new Set(setupSheet.getRange("A:A").getValues().flat());

  var threads = GmailApp.search("to:helpdesk@godk.be");
  var parentFolderId = "Your-Shared-Drive-Folder-ID";

  for (var i in threads) {
    processThread(threads[i], existingEmailIds, nameLookupMap, allowedEmails, sourceSheet, targetSheet, parentFolderId);
  }
}

// Function to clear content in specific sheets
function clearSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = ["ontvangenMails", "verwerkMailBody"];

  for (var i = 0; i < sheetNames.length; i++) {
    var sheet = ss.getSheetByName(sheetNames[i]);
    sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
  }
}
