// Function to create or get a folder in Google Drive
function createOrGetFolder(threadId, parentFolderId) {
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var folders = parentFolder.getFoldersByName(threadId);
  var folder;
  var isNewFolder = false;

  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = parentFolder.createFolder(threadId);
    isNewFolder = true;
  }
  return [folder, isNewFolder];
}

// Function to save attachments to Google Drive folder
function saveAttachmentsToFolder(attachments, folder) {
  var attachmentUrls = [];
  for (var i = 0; i < attachments.length; i++) {
    var blob = attachments[i].copyBlob();
    var file = folder.createFile(blob);
    attachmentUrls.push(file.getUrl());
  }
  return attachmentUrls;
}

// Function to check if folder exists in Google Drive
function doesFolderExist(threadId, parentFolderId) {
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var folders = parentFolder.getFoldersByName(threadId);
  return folders.hasNext();
}
