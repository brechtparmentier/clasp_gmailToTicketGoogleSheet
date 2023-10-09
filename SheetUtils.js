// Function to append data to both source and target sheets
function appendToSheets(date, subject, body, emailId, emailType, senderName, sourceSheet, targetSheet, processedData, attachmentUrls) {
  sourceSheet.appendRow([date, subject, body, emailId, emailType]);
  var rowOfMaxValue = findRowOfMaxValue(targetSheet, "A") + 1;

  var values = [
    [date], // B
    ["helpdesk"], // C
    [senderName], // E
    [subject + "\n" + body], // G
  ];

  targetSheet.getRange(rowOfMaxValue, 2, 4, 1).setValues(values);
  targetSheet.getRange(rowOfMaxValue, 8).setValue(emailId); // Column H for emailId
  targetSheet.getRange(rowOfMaxValue, 9).setValue(emailType); // Column I for emailType

  if (attachmentUrls && attachmentUrls.length > 0) {
    targetSheet.getRange("H" + targetSheet.getLastRow()).setValue(attachmentUrls.join(", "));
  }
}

// Function to find the row index of the maximum value in a column
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
