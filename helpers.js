// Deze functie vindt de eerste lege rij in een kolom
function findFirstEmptyRow(sheet, column) {
  var data = sheet.getRange(column + "1:" + column + sheet.getLastRow()).getValues();
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) {
      return i + 1;
    }
  }
  return data.length + 1;
}

// Deze functie vindt de rij van het grootste getal in een kolom
function findRowOfMaxValue(sheet, column) {
  var data = sheet.getRange(column + "1:" + column + sheet.getLastRow()).getValues();
  var maxVal = -Infinity;
  var maxRow = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] > maxVal) {
      maxVal = data[i][0];
      maxRow = i + 1; // Rijen in Google Sheets beginnen bij 1
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
