var headers = ["ID", "Datum vraag", "Aangemaakt Door", "Voor wie", "Van wie", "Prior?", "Vraag", "Oplossing", "link naar Doc", "Datum oplossing", "Status", "Ok?", "Nodige Tijd", "Extra uren", "Link", "Mail verstuurd op"];

// Function to append data to both source and target sheets
function appendToSheets(date, subject, body, emailId, emailType, senderName, sourceSheet, inkomendeVragenSheet, processedData, attachmentUrls) {
  sourceSheet.appendRow([date, subject, body, emailId, emailType]);
  var rowOfMaxValue = findRowOfMaxValue(inkomendeVragenSheet, "A") + 1;

  // Map values to headers
  var headerToValue = {
    "Datum vraag": date,
    "Aangemaakt Door": "helpdesk",
    "Van wie": senderName,
    Vraag: subject + "\n" + body,
  };

  var values = headers.map(function (header) {
    return headerToValue[header] || ""; // return the mapped value or an empty string
  });

  inkomendeVragenSheet.getRange(rowOfMaxValue, 1, 1, headers.length).setValues([values]); // Place the values into the correct columns

  if (attachmentUrls && attachmentUrls.length > 0) {
    inkomendeVragenSheet.getRange("H" + inkomendeVragenSheet.getLastRow()).setValue(attachmentUrls.join(", "));
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
