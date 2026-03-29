/**
 * Kindoo Step 3: Automated Notifications
 * This script should be pasted into the Apps Script editor of the GOOGLE SHEET.
 */

function onFormSubmitTrigger(e) {
  // 1. Get the data from the form submission
  var responses = e.namedValues;
  
  var requesterName = responses['Requester Name'][0];
  var requesterEmail = responses['Requester Email'][0];
  var building = responses['Building Location'][0];
  var ward = responses['Requester\'s Ward'][0];  

  // --- HUMAN READABLE TIME FIX ---
  // We convert the string to a Date object, then format it
  var rawStart = new Date(responses['Access Start (Date & Time)'][0]);
  var rawEnd = new Date(responses['Access End (Date & Time)'][0]);
  
  // Format: M/d/yyyy h:mm a (e.g., 3/23/2026 1:01 PM)
  var start = Utilities.formatDate(rawStart, Session.getScriptTimeZone(), "M/d/yyyy h:mm a");
  var end = Utilities.formatDate(rawEnd, Session.getScriptTimeZone(), "M/d/yyyy h:mm a");

  // 2. Map Wards to Bishop Emails
  // UPDATE THESE with the actual email addresses
  var bishopEmails = {
    '1st Ward': 'WARD_1_EMAIL',
    '2nd Ward': 'WARD_2_EMAIL',
    '4th Ward': 'WARD_4_EMAIL',
    '5th Ward': 'WARD_5_EMAIL',
    '7th Ward': 'WARD_7_EMAIL'
  };

  var targetBishop = bishopEmails[ward] || "STAKE_TECHNOLOGY_SPECIALIST"; //Fall back email.

  // 3. Email 1: The Bishop's FYI
  var bishopSubject = "FYI: Building Access Request - " + ward;
  var bishopBody = "Bishop,\n\nThis is an automated notification that a building access request has been scheduled by a member of your ward.\n\n" +
                   "Requester: " + requesterName + "\n" +
                   "Building: " + building + "\n" +
                   "Time: " + start + " to " + end + "\n\n" +
                   "No action is required on your part. This has been vetted by the Building Scheduler.";
  
  MailApp.sendEmail(targetBishop, bishopSubject, bishopBody);

  // 4. Email 2: The Requester's Confirmation
  var requesterSubject = "Kindoo Access Confirmed: " + building;
  var requesterBody = "Hello " + requesterName + ",\n\n" +
                      "Your access request for the " + building + " has been logged.\n\n" +
                      "Scheduled Time: " + start + " to " + end + "\n\n" +
                      "IMPORTANT: Your digital key will be assigned to your Kindoo app 7 days before the event begins. " +
                      "Please ensure your Kindoo app email matches your Member Tools email (" + requesterEmail + ").";

  MailApp.sendEmail(requesterEmail, requesterSubject, requesterBody);

  // 5. Ledger Documentation
  // The data is automatically added to the sheet by Google Forms. 
  // We can add a "Status" note to the last column if desired.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, sheet.getLastColumn() + 1).setValue("Vetted and Scheduled");
}