function addEmailsToSheetRechtstreeks() {
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ontvangenMails");
  var lastRow = sourceSheet.getLastRow();
  var existingEmailIds = sourceSheet.getRange(1, 4, lastRow).getValues().flat();
  var targetSpreadsheet = SpreadsheetApp.openById("19EApoPk-o7tRaYx1EE8bwPl1Cg0V7KmlOjehPL0DZDI");
  var targetSheet = targetSpreadsheet.getSheetByName("_inkomendeVragen");
  var nameLookupSheet = targetSpreadsheet.getSheetByName("vanWie");
  var setupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("setup");
  var nameLookupRange = nameLookupSheet.getRange("A:B").getValues();
  var allowedEmails = setupSheet.getRange("A:A").getValues().flat();
  var threads = GmailApp.search("to:helpdesk@godk.be");
  var parentFolderId = "1k8ZJhnXKPjWCiAllnUfBrswlRKySt0RB"; // Shared Drive map ID

  for (var i = 0; i < threads.length; i++) {
    var threadId = threads[i].getId();
    var folder = createOrGetFolder(threadId, parentFolderId);
    var messages = threads[i].getMessages();

    for (var j = 0; j < messages.length; j++) {
      var emailId = messages[j].getId();
      var emailType = isForwarded(subject, body) ? "forwarded" : "rechtstreeks";
      if (existingEmailIds.indexOf(emailId) === -1) {
        var attachments = messages[j].getAttachments();
        var subject = messages[j].getSubject();
        var body = messages[j].getPlainBody();
        var date = messages[j].getDate();
        var sender = messages[j].getFrom();
        var emailMatch = sender.match(/<(.+)>/);
        var emailOnly = emailMatch ? emailMatch[1] : sender;
        var senderName = null;
        for (var k = 0; k < nameLookupRange.length; k++) {
          if (nameLookupRange[k][1] === emailOnly) {
            senderName = nameLookupRange[k][0];
            break;
          }
        }
        if (allowedEmails.indexOf(emailOnly) !== -1) {
          sourceSheet.appendRow([date, subject, body, emailId, emailType]); // Extra kolom voor emailType
          existingEmailIds.push(emailId);
          var rowOfMaxValue = findRowOfMaxValue(targetSheet, "A") + 1;
          targetSheet.getRange("B" + rowOfMaxValue).setValue(date);
          targetSheet.getRange("C" + rowOfMaxValue).setValue("helpdesk");
          targetSheet.getRange("E" + rowOfMaxValue).setValue(senderName || "Onbekend");
          targetSheet.getRange("G" + rowOfMaxValue).setValue(subject + "\n" + body);
          if (attachments.length > 0 && allowedEmails.indexOf(emailOnly) !== -1) {
            var attachmentString = attachmentUrls.join(", ");
            targetSheet.getRange("H" + rowOfMaxValue).setValue(attachmentString);
          }
        }
        if (attachments.length > 0) {
          var folder = createOrGetFolder(threadId, parentFolderId);
          var attachmentUrls = saveAttachmentsToFolder(attachments, folder);
          var attachmentString = attachmentUrls.join(", ");
        }
        targetSheet.getRange("I" + rowOfMaxValue).setValue(isForwarded);
      }
    }
    threads[i].markRead();
  }
}

function findFirstEmptyRow(sheet, column) {
  var data = sheet.getRange(column + "1:" + column + sheet.getLastRow()).getValues();
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) {
      return i + 1;
    }
  }
  return data.length + 1;
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

function processEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ontvangenMails");
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("verwerkMailBody") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("verwerkMailBody");
  var emails = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var newData = [["Van", "Datum", "Onderwerp", "Aan", "Bericht", "Attachments", "Extra1", "Extra2"]];

  for (var i = 0; i < emails.length; i++) {
    var email = emails[i];
    var [date, subject, body, emailId, emailType, attachmentUrls] = email;
    var fromRegex = /Van: (.+)<(.+)>/;
    var dateRegex = /Date: (.+)/;
    var subjectRegex = /Subject: (.+)/;
    var toRegex = /To: (.+)<(.+)>/;
    var from = (body.match(fromRegex) || [])[1];
    var dateFromEmail = (body.match(dateRegex) || [])[1];
    var subjectFromEmail = (body.match(subjectRegex) || [])[1];
    var to = (body.match(toRegex) || [])[1];
    var message = "";
    newData.push([from, dateFromEmail, subjectFromEmail, to, message, attachmentUrls, emailType, ""]);
  }
  newSheet.getRange(1, 1, newData.length, 8).setValues(newData);
}
