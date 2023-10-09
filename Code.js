function addEmailsToSheetRechtstreeks() {
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ontvangenMails');
  var lastRow = sourceSheet.getLastRow();
  var existingEmailIds = sourceSheet.getRange(1, 4, lastRow).getValues().flat();
  var targetSpreadsheet = SpreadsheetApp.openById('19EApoPk-o7tRaYx1EE8bwPl1Cg0V7KmlOjehPL0DZDI');
  var targetSheet = targetSpreadsheet.getSheetByName('_inkomendeVragen');
  var nameLookupSheet = targetSpreadsheet.getSheetByName('vanWie');
  var setupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setup');
  var nameLookupRange = nameLookupSheet.getRange('A:B').getValues();
  var allowedEmails = setupSheet.getRange('A:A').getValues().flat();
  var threads = GmailApp.search('to:helpdesk@godk.be');
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
          var rowOfMaxValue = findRowOfMaxValue(targetSheet, 'A') + 1;
          targetSheet.getRange('B' + rowOfMaxValue).setValue(date);
          targetSheet.getRange('C' + rowOfMaxValue).setValue("helpdesk");
          targetSheet.getRange('E' + rowOfMaxValue).setValue(senderName || "Onbekend");
          targetSheet.getRange('G' + rowOfMaxValue).setValue(subject + "\n" + body);

          if (attachments.length > 0 && allowedEmails.indexOf(emailOnly) !== -1) {
            var attachmentString = attachmentUrls.join(", ");
            targetSheet.getRange('H' + rowOfMaxValue).setValue(attachmentString);
          }
        }

        if (attachments.length > 0) {
  var folder = createOrGetFolder(threadId, parentFolderId);
  var attachmentUrls = saveAttachmentsToFolder(attachments, folder);
  var attachmentString = attachmentUrls.join(", ");
}
        targetSheet.getRange('I' + rowOfMaxValue).setValue(isForwarded);

      }
    }
    threads[i].markRead();
  }
}


// function addEmailsToSheetRechtstreeks() {
//   var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ontvangenMails');
//   var targetSpreadsheet = SpreadsheetApp.openById('19EApoPk-o7tRaYx1EE8bwPl1Cg0V7KmlOjehPL0DZDI');
//   var targetSheet = targetSpreadsheet.getSheetByName('_inkomendeVragen');
//   var nameLookupSheet = targetSpreadsheet.getSheetByName('vanWie');
//   var setupSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setup');
//     var lastRow = sourceSheet.getLastRow();
//   var existingEmailIds = sourceSheet.getRange(1, 4, lastRow).getValues().flat(); // Stel dat de e-mail ID's in kolom 4 worden opgeslagen


//   // Haal de lijst met toegestane e-mailadressen op
//   var allowedEmails = setupSheet.getRange('A:A').getValues().flat();

//   // Haal de lijst met namen en e-mailadressen op uit 'vanWie'
//   var nameLookupRange = nameLookupSheet.getRange('A:B').getValues();
//   Logger.log(nameLookupRange)

//   // Start logging
//   Logger.log("Start het verwerken van e-mail threads...");

//   var threads = GmailApp.search('to:helpdesk@godk.be');

//   for (var i = 0; i < threads.length; i++) {
//     var messages = threads[i].getMessages();

//     // Log informatie over de huidige thread
//     //Logger.log("Verwerken van thread " + (i + 1) + " met " + messages.length + " berichten.");

//     for (var j = 0; j < messages.length; j++) {
//       var emailId = messages[j].getId();
//       if (existingEmailIds.indexOf(emailId) === -1) { // Controleer of de e-mail ID al bestaat

//       var subject = messages[j].getSubject();
//       var body = messages[j].getPlainBody();
//       var date = messages[j].getDate();
//       var sender = messages[j].getFrom();
//       var emailMatch = sender.match(/<(.+)>/);
//       var emailOnly = emailMatch ? emailMatch[1] : sender;
//       //Logger.log("Afzender van de e-mail: " + emailOnly);


//       // Log het onderwerp van de huidige e-mail
//       //Logger.log("Onderwerp van het bericht: " + subject);

//       // Controleer of de afzender in de toegestane lijst staat
//       if (allowedEmails.indexOf(emailOnly) !== -1) {
//         var senderName = null;
//         for (var k = 0; k < nameLookupRange.length; k++) {
//           if (nameLookupRange[k][1] === sender) {
//             senderName = nameLookupRange[k][0];
//             Logger.log('senderName: ' + senderName)
//             break;
//           }
//         }

//         // Voeg gegevens toe aan de bron-sheet
//         sourceSheet.appendRow([date, subject, body]);

//         //var firstEmptyRow = findFirstEmptyRow(targetSheet, 'G');
//         var rowOfMaxValue = findRowOfMaxValue(targetSheet, 'A') + 1;



//         // Vind de volgende lege rij in het doeltabblad
//         var lastRow = targetSheet.getLastRow() + 1;

//         // Voeg gegevens toe aan specifieke cellen in het doeltabblad
//         targetSheet.getRange('C' + rowOfMaxValue).setValue("helpdesk");
//         targetSheet.getRange('E' + rowOfMaxValue).setValue(senderName || "Onbekend"); // Gebruik 'Onbekend' als er geen naam gevonden is
//         targetSheet.getRange('G' + rowOfMaxValue).setValue(subject + "\n" + body);


//         // Log dat de rij succesvol is toegevoegd
//         Logger.log("Rij succesvol toegevoegd aan doeltabblad.");
//       }
//     }
//     }
//     threads[i].markRead();
//   }

//   // Einde logging
//   Logger.log("Verwerking van e-mail threads voltooid.");
// }



