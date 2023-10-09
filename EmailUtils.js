// Function to check if the email is forwarded
function isForwarded(subject, body) {
  return subject.startsWith("Fwd:") || subject.startsWith("FW:") || body.includes("Forwarded Message") || body.includes("--- Forwarded message ---");
}

// Function to get sender's name from the lookup map
function getSenderName(emailOnly, nameLookupMap) {
  return nameLookupMap.get(emailOnly) || "Onbekend";
}

// Function to process the body of the email and extract relevant information
function processEmailBody(body) {
  var fromRegex = /Van: (.+)<(.+)>/;
  var dateRegex = /Date: (.+)/;
  var subjectRegex = /Subject: (.+)/;
  var toRegex = /To: (.+)<(.+)>/;

  var from = (body.match(fromRegex) || [])[1];
  var date = (body.match(dateRegex) || [])[1];
  var subject = (body.match(subjectRegex) || [])[1];
  var to = (body.match(toRegex) || [])[1];

  return [from, date, subject, to];
}
