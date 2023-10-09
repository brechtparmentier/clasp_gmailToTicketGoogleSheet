function addEmailsToSheetRechtstreeks() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = sheets.getSheetByName("ontvangenMails");
  var existingEmailIds = sourceSheet.getRange(1, 4, sourceSheet.getLastRow()).getValues().flat();
  var targetSpreadsheet = SpreadsheetApp.openById("19EApoPk-o7tRaYx1EE8bwPl1Cg0V7KmlOjehPL0DZDI");
  var targetSheet = targetSpreadsheet.getSheetByName("_inkomendeVragen");
  var nameLookupSheet = targetSpreadsheet.getSheetByName("vanWie");
  var setupSheet = sheets.getSheetByName("setup");
  var nameLookupRange = nameLookupSheet.getRange("A:B").getValues();
  var allowedEmails = setupSheet.getRange("A:A").getValues().flat();
  var threads = GmailApp.search("to:helpdesk@godk.be");
  var parentFolderId = "1k8ZJhnXKPjWCiAllnUfBrswlRKySt0RB"; // Shared Drive map ID

  for (var i in threads) {
    processThread(threads[i], existingEmailIds, nameLookupRange, allowedEmails, sourceSheet, targetSheet, parentFolderId);
  }
}

function processThread(thread, existingEmailIds, nameLookupRange, allowedEmails, sourceSheet, targetSheet, parentFolderId) {
  var messages = thread.getMessages();
  var threadId = thread.getId();
  var folder = createOrGetFolder(threadId, parentFolderId);

  for (var j in messages) {
    var message = messages[j];
    var emailId = message.getId();
    var subject = message.getSubject();
    var body = message.getPlainBody();
    var emailType = isForwarded(subject, body) ? "forwarded" : "rechtstreeks";

    if (existingEmailIds.indexOf(emailId) === -1) {
      processMessage(message, subject, body, emailId, emailType, nameLookupRange, allowedEmails, sourceSheet, targetSheet, folder);
      existingEmailIds.push(emailId);
    }
  }
  thread.markRead();
}

function processMessage(message, subject, body, emailId, emailType, nameLookupRange, allowedEmails, sourceSheet, targetSheet, folder) {
  var date = message.getDate();
  var sender = message.getFrom();
  var emailMatch = sender.match(/<(.+)>/);
  var emailOnly = emailMatch ? emailMatch[1] : sender;
  var senderName = getSenderName(emailOnly, nameLookupRange);
  var attachments = message.getAttachments();

  if (allowedEmails.indexOf(emailOnly) !== -1) {
appendToSheets(date, subject, body, emailId, emailType, senderName, sourceSheet, targetSheet);
  }

  if (attachments.length > 0) {
    var attachmentUrls = saveAttachmentsToFolder(attachments, folder);
    targetSheet.getRange("H" + targetSheet.getLastRow()).setValue(attachmentUrls.join(", "));
  }
}

function findRowOfMaxValue(sheet, column) {
  var data = sheet.getRange(column + "1:" + column + sheet.getLastRow()).getValues();
  var maxVal = -Infinity;
  var maxRow = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] > maxVal) {
      maxVal = data[i][0];
      maxRow = i + 1;
    }
  }
  return maxRow;
}

function createOrGetFolder(threadId, parentFolderId) {
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var folders = parentFolder.getFoldersByName(threadId);
  var folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = parentFolder.createFolder(threadId);
  }
  return folder;
}

function saveAttachmentsToFolder(attachments, folder) {
  var attachmentUrls = [];
  for (var i = 0; i < attachments.length; i++) {
    var blob = attachments[i].copyBlob();
    var file = folder.createFile(blob);
    attachmentUrls.push(file.getUrl());
  }
  return attachmentUrls;
}

function isForwarded(subject, body) {
  if (subject.startsWith("Fwd:") || subject.startsWith("FW:")) {
    return true;
  }
  if (body.includes("Forwarded Message") || body.includes("--- Forwarded message ---")) {
    return true;
  }
  return false;
}

function getSenderName(emailOnly, nameLookupRange) {
  var senderName = null;
  for (var k = 0; k < nameLookupRange.length; k++) {
    if (nameLookupRange[k][1] === emailOnly) {
      senderName = nameLookupRange[k][0];
      break;
    }
  }
  return senderName || "Onbekend";
}

function appendToSheets(date, subject, body, emailId, emailType, senderName, sourceSheet, targetSheet) {
  sourceSheet.appendRow([date, subject, body, emailId, emailType]);
  var rowOfMaxValue = findRowOfMaxValue(targetSheet, "A") + 1;
  targetSheet.getRange("B" + rowOfMaxValue).setValue(date);
  targetSheet.getRange("C" + rowOfMaxValue).setValue("helpdesk");
  targetSheet.getRange("E" + rowOfMaxValue).setValue(senderName);
  targetSheet.getRange("G" + rowOfMaxValue).setValue(subject + "\n" + body);
}