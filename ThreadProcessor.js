// Function to process each thread
function processThread(thread, existingEmailIds, nameLookupMap, allowedEmails, sourceSheet, inkomendeVragenSheet, parentFolderId) {
  var messages = thread.getMessages();
  var threadId = thread.getId();

  // Use destructuring to get folder and isNewFolder
  var [folder, isNewFolder] = createOrGetFolder(threadId, parentFolderId);

  for (var j in messages) {
    var message = messages[j];
    var emailId = message.getId();
    var subject = message.getSubject();
    var body = message.getPlainBody();
    var emailType = isForwarded(subject, body) ? "forwarded" : "rechtstreeks";

    if (!existingEmailIds.has(emailId)) {
      processMessage(message, isNewFolder, subject, body, emailId, emailType, nameLookupMap, allowedEmails, sourceSheet, inkomendeVragenSheet, folder);
      existingEmailIds.add(emailId);
    }
  }
  thread.markRead();
}

// Function to process each message within a thread
function processMessage(message, isNewFolder, subject, body, emailId, emailType, nameLookupMap, allowedEmails, sourceSheet, inkomendeVragenSheet, folder) {
  Logger.log("start processMessage");
  var date = message.getDate();
  var sender = message.getFrom();
  Logger.log("sender: " + sender);
  Logger.log("body: " + body);
  var emailMatch = sender.match(/<(.+)>/);
  var emailOnly = emailMatch ? emailMatch[1] : sender;
  var senderName = getSenderName(emailOnly, nameLookupMap);

  if (allowedEmails.has(emailOnly)) {
    Logger.log("Processing email from " + senderName + " (" + emailOnly + ")");
    var processedData = processEmailBody(body);
    var attachmentUrls = [];

    if (isNewFolder && message.getAttachments().length > 0) {
      attachmentUrls = saveAttachmentsToFolder(message.getAttachments(), folder);
    }

    appendToSheets(date, subject, body, emailId, emailType, senderName, sourceSheet, inkomendeVragenSheet, processedData, attachmentUrls);

    var verwerkMailBodySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("verwerkMailBody") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("verwerkMailBody");
    Logger.log("verwerkMailBodySheet: " + verwerkMailBodySheet.getName());
    verwerkMailBodySheet.appendRow([...processedData, emailId, emailType]); // Hier voeg je processedData toe aan de nieuwe sheet
    Logger.log("append " + processedData + " to verwerkMailBodySheet" + emailId + " " + emailType + emailType + " " + emailType);
    if (attachmentUrls && attachmentUrls.length > 0) {
      Logger.log("attachmentUrls: " + attachmentUrls);
      verwerkMailBodySheet.getRange("H" + inkomendeVragenSheet.getLastRow()).setValue(attachmentUrls.join(", "));
    }
  } else {
    Logger.log("Email from " + senderName + " (" + emailOnly + ") is not allowed. Skipping.");
  }
}
