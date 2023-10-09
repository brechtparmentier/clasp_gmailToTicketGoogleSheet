function processEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ontvangenMails");
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("verwerkMailBody") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("verwerkMailBody");

  // Kolommen uitlezen
  var emails = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();

  // Voorbereiden van data voor nieuwe sheet
  var newData = [["Van", "Datum", "Onderwerp", "Aan", "Bericht", "Attachments", "Extra1", "Extra2"]];

  for (var i = 0; i < emails.length; i++) {
    var email = emails[i];
    var [date, subject, body, emailId, emailType, attachmentUrls] = email; // Verander dit naar jouw kolomstructuur

    var fromRegex = /Van: (.+)<(.+)>/;
    var dateRegex = /Date: (.+)/;
    var subjectRegex = /Subject: (.+)/;
    var toRegex = /To: (.+)<(.+)>/;

    var from = (body.match(fromRegex) || [])[1];
    var dateFromEmail = (body.match(dateRegex) || [])[1];
    var subjectFromEmail = (body.match(subjectRegex) || [])[1];
    var to = (body.match(toRegex) || [])[1];

    var message = ""; // hier verwerk je 'body' verder

    newData.push([from, dateFromEmail, subjectFromEmail, to, message, attachmentUrls, emailType, ""]);
  }

  // In één keer alle data invullen in de nieuwe sheet
  newSheet.getRange(1, 1, newData.length, 8).setValues(newData);
}












// function processEmails() {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ontvangenMails");
//   var newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("verwerkMailBody") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("verwerkMailBody");

//   var emails = sheet.getRange("C1:C" + sheet.getLastRow()).getValues();

//   // Definieer de headers voor de nieuwe sheet
//   newSheet.getRange("A1:H1").setValues([["Van", "Datum", "Onderwerp", "Aan", "Bericht", "Attachments", "Extra1", "Extra2"]]);

//   for (var i = 0; i < emails.length; i++) {
//     var email = emails[i][0];

//     // RegExp om informatie te extraheren
//     var fromRegex = /Van: (.+)<(.+)>/;
//     var dateRegex = /Date: (.+)/;
//     var subjectRegex = /Subject: (.+)/;
//     var toRegex = /To: (.+)<(.+)>/;

//     var from = (email.match(fromRegex) || [])[1];
//     var date = (email.match(dateRegex) || [])[1];
//     var subject = (email.match(subjectRegex) || [])[1];
//     var to = (email.match(toRegex) || [])[1];


//     // De rest van het bericht
//     var splitEmail = email.split("To:");
//     if (splitEmail.length > 1) {
//       var message = splitEmail[1];
//       message = message.substring(message.indexOf("\n") + 1).trim();

//     } else {
//       var message = ""; // of een andere standaardwaarde
//     }


//     // Vul de nieuwe sheet
//     newSheet.appendRow([from, date, subject, to, message, "", "", ""]);
//   }
// }
