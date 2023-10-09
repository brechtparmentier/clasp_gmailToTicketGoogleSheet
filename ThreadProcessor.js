// Function to process each thread
function processThread(thread, existingEmailIds, nameLookupMap, allowedEmails, sourceSheet, targetSheet, parentFolderId) {
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
      processMessage(message, isNewFolder, subject, body, emailId, emailType, nameLookupMap, allowedEmails, sourceSheet, targetSheet, folder);
      existingEmailIds.add(emailId);
    }
  }
  thread.markRead();
}

// Function to process each message within a thread
function processMessage(message, isNewFolder, subject, body, emailId, emailType, nameLookupMap, allowedEmails, sourceSheet, targetSheet, folder) {
  var processedData = [];
  var date = message.getDate();
  var sender = message.getFrom();
  var emailMatch = sender.match(/<(.+)>/);
  var emailOnly = emailMatch ? emailMatch[1] : sender;
  var senderName = getSenderName(emailOnly, nameLookupMap);
  var attachments = message.getAttachments();

  if (allowedEmails.has(emailOnly)) {
    processedData = processEmailBody(body);
    appendToSheets(date, subject, body, emailId, emailType, senderName, sourceSheet, targetSheet, processedData);
  }

  if (isNewFolder && attachments.length > 0) {
    var attachmentUrls = saveAttachmentsToFolder(attachments, folder);
    appendToSheets(date, subject, body, emailId, emailType, senderName, sourceSheet, targetSheet, processedData, attachmentUrls);
  } else {
    appendToSheets(date, subject, body, emailId, emailType, senderName, sourceSheet, targetSheet, processedData);
  }
}
