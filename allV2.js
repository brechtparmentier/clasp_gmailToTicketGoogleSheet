// function addEmailsToSheetRechtstreeks() {
//   var sheets = SpreadsheetApp.getActiveSpreadsheet();
//   var sourceSheet = sheets.getSheetByName("ontvangenMails");
//   var existingEmailIds = new Set(sourceSheet.getRange(1, 4, sourceSheet.getLastRow()).getValues().flat());
//   var targetSpreadsheet = SpreadsheetApp.openById("19EApoPk-o7tRaYx1EE8bwPl1Cg0V7KmlOjehPL0DZDI");
//   var targetSheet = targetSpreadsheet.getSheetByName("_inkomendeVragen");
//   var nameLookupSheet = targetSpreadsheet.getSheetByName("vanWie");
//   var setupSheet = sheets.getSheetByName("setup");
//   var nameLookupRange = nameLookupSheet.getRange("A:B").getValues();
//   var allowedEmails = new Set(setupSheet.getRange("A:A").getValues().flat());
//   var threads = GmailApp.search("to:helpdesk@godk.be");
//   var parentFolderId = "1k8ZJhnXKPjWCiAllnUfBrswlRKySt0RB"; // Shared Drive map ID

//   for (var i in threads) {
//     processThread(threads[i], existingEmailIds, nameLookupRange, allowedEmails, sourceSheet, targetSheet, parentFolderId);
//   }
// }

// function clearSheets() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var sheetNames = ["ontvangenMails", "verwerkMailBody"];

//   for (var i = 0; i < sheetNames.length; i++) {
//     var sheet = ss.getSheetByName(sheetNames[i]);
//     sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
//   }
// }

// function processThread(thread, existingEmailIds, nameLookupRange, allowedEmails, sourceSheet, targetSheet, parentFolderId) {
//   var messages = thread.getMessages();
//   var threadId = thread.getId();
//   var folder = createOrGetFolder(threadId, parentFolderId);

//   var isNewFolder = !doesFolderExist(threadId, parentFolderId); // Nieuw toegevoegd

//   for (var j in messages) {
//     var message = messages[j];
//     var emailId = message.getId();
//     var subject = message.getSubject();
//     var body = message.getPlainBody();
//     var emailType = isForwarded(subject, body) ? "forwarded" : "rechtstreeks";

//     if (!existingEmailIds.has(emailId)) {
//       processMessage(message, isNewFolder, subject, body, emailId, emailType, nameLookupRange, allowedEmails, sourceSheet, targetSheet, folder);
//       existingEmailIds.add(emailId);
//     }
//   }
//   thread.markRead();
// }

// function processMessage(message, isNewFolder, subject, body, emailId, emailType, nameLookupRange, allowedEmails, sourceSheet, targetSheet, folder) {
//   var processedData = []; // Initialize to an empty array

//   var date = message.getDate();
//   var sender = message.getFrom();
//   var emailMatch = sender.match(/<(.+)>/);
//   var emailOnly = emailMatch ? emailMatch[1] : sender;
//   var senderName = getSenderName(emailOnly, nameLookupRange);
//   var attachments = message.getAttachments();
//   Logger.log("Processing email from " + sender + " with subject " + subject + " and emailId " + emailId + "with number of attachments " + attachments.length);

//   if (allowedEmails.has(emailOnly)) {
//     processedData = processEmailBody(body); // Reassign the value
//     appendToSheets(date, subject, body, emailId, emailType, senderName, sourceSheet, targetSheet, processedData);
//   }

//   if (isNewFolder && attachments.length > 0) {
//     // Save only if the folder is new
//     var attachmentUrls = saveAttachmentsToFolder(attachments, folder);
//     appendToSheets(date, subject, body, emailId, emailType, senderName, sourceSheet, targetSheet, processedData, attachmentUrls);
//   } else {
//     appendToSheets(date, subject, body, emailId, emailType, senderName, sourceSheet, targetSheet, processedData);
//   }
// }

// function findRowOfMaxValue(sheet, column) {
//   var data = sheet.getRange(column + "1:" + column + sheet.getLastRow()).getValues();
//   var maxVal = -Infinity;
//   var maxRow = -1;
//   for (var i = 0; i < data.length; i++) {
//     if (data[i][0] > maxVal) {
//       maxVal = data[i][0];
//       maxRow = i + 1;
//     }
//   }
//   return maxRow;
// }

// function createOrGetFolder(threadId, parentFolderId) {
//   var parentFolder = DriveApp.getFolderById(parentFolderId);
//   var folders = parentFolder.getFoldersByName(threadId);
//   var folder;
//   if (folders.hasNext()) {
//     folder = folders.next();
//   } else {
//     folder = parentFolder.createFolder(threadId);
//   }
//   return folder;
// }

// function saveAttachmentsToFolder(attachments, folder) {
//   var attachmentUrls = [];
//   for (var i = 0; i < attachments.length; i++) {
//     var blob = attachments[i].copyBlob();
//     var file = folder.createFile(blob);
//     attachmentUrls.push(file.getUrl());
//   }
//   return attachmentUrls;
// }

// function isForwarded(subject, body) {
//   if (subject.startsWith("Fwd:") || subject.startsWith("FW:")) {
//     return true;
//   }
//   if (body.includes("Forwarded Message") || body.includes("--- Forwarded message ---")) {
//     return true;
//   }
//   return false;
// }

// function getSenderName(emailOnly, nameLookupRange) {
//   var senderName = null;
//   for (var k = 0; k < nameLookupRange.length; k++) {
//     if (nameLookupRange[k][1] === emailOnly) {
//       senderName = nameLookupRange[k][0];
//       break;
//     }
//   }
//   return senderName || "Onbekend";
// }

// function appendToSheets(date, subject, body, emailId, emailType, senderName, sourceSheet, targetSheet, processedData, attachmentUrls) {
//   sourceSheet.appendRow([date, subject, body, emailId, emailType]);
//   var rowOfMaxValue = findRowOfMaxValue(targetSheet, "A") + 1;

//   var values = [
//     [date], // B
//     ["helpdesk"], // C
//     [senderName], // E
//     [subject + "\n" + body], // G
//   ];

//   targetSheet.getRange(rowOfMaxValue, 2, 4, 1).setValues(values);
//   targetSheet.getRange(rowOfMaxValue, 8).setValue(emailId); // Kolom H voor emailId
//   targetSheet.getRange(rowOfMaxValue, 9).setValue(emailType); // Kolom I voor emailType

//   var newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("verwerkMailBody") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("verwerkMailBody");
//   newSheet.appendRow([...processedData, emailId, emailType]); // Hier voeg je processedData toe aan de nieuwe sheet
//   if (attachmentUrls && attachmentUrls.length > 0) {
//     targetSheet.getRange("H" + targetSheet.getLastRow()).setValue(attachmentUrls.join(", "));
//   }
// }

// function processEmailBody(body) {
//   var fromRegex = /Van: (.+)<(.+)>/;
//   var dateRegex = /Date: (.+)/;
//   var subjectRegex = /Subject: (.+)/;
//   var toRegex = /To: (.+)<(.+)>/;

//   var from = (body.match(fromRegex) || [])[1];
//   var date = (body.match(dateRegex) || [])[1];
//   var subject = (body.match(subjectRegex) || [])[1];
//   var to = (body.match(toRegex) || [])[1];

//   // Voeg verder verwerking toe indien nodig

//   return [from, date, subject, to];
// }

// function doesFolderExist(threadId, parentFolderId) {
//   var parentFolder = DriveApp.getFolderById(parentFolderId);
//   var folders = parentFolder.getFoldersByName(threadId);
//   return folders.hasNext();
// }
