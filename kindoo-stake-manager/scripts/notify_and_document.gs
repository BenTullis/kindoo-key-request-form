/**
 * Kindoo Step 3: Automated Notifications
 * This script should be pasted into the Apps Script editor of the GOOGLE SHEET.
 */

function getSecret_(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error('Missing script property: ' + key);
  }
  return value;
}

function getBishopEmails_() {
  return {
    '1st Ward': getSecret_('WARD_1_EMAIL'),
    '2nd Ward': getSecret_('WARD_2_EMAIL'),
    '4th Ward': getSecret_('WARD_4_EMAIL'),
    '5th Ward': getSecret_('WARD_5_EMAIL'),
    '7th Ward': getSecret_('WARD_7_EMAIL')
  };
}

function getSingleResponseValue_(responses, key) {
  var value = responses[key];
  if (Array.isArray(value) && value.length > 0) {
    return String(value[0]).trim();
  }
  if (typeof value === 'string') {
    return value.trim();
  }
  return '';
}

function getValueFromEventRow_(e, headerName) {
  var sheet = e.range.getSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerIndex = headers.indexOf(headerName);
  if (headerIndex === -1) {
    return null;
  }

  var rowValues = sheet.getRange(e.range.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
  return rowValues[headerIndex];
}

function getSubmissionValue_(e, responses, possibleHeaders) {
  for (var i = 0; i < possibleHeaders.length; i++) {
    var fromNamedValues = getSingleResponseValue_(responses, possibleHeaders[i]);
    if (fromNamedValues) {
      return fromNamedValues;
    }
  }

  for (var j = 0; j < possibleHeaders.length; j++) {
    var fromRow = getValueFromEventRow_(e, possibleHeaders[j]);
    if (fromRow !== null && fromRow !== '') {
      return fromRow;
    }
  }

  return '';
}

function getRequesterEmail_(e, responses) {
  var requesterEmail = getSubmissionValue_(e, responses, [
    'Requester Email',
    'Email Address',
    'Email'
  ]);

  if (!requesterEmail) {
    throw new Error('Requester email is missing from the submission row.');
  }

  return String(requesterEmail).trim();
}

function getOrCreateStatusColumn_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statusHeader = 'Status';
  var statusIndex = headers.indexOf(statusHeader);

  if (statusIndex !== -1) {
    return statusIndex + 1;
  }

  var newColumn = sheet.getLastColumn() + 1;
  sheet.getRange(1, newColumn).setValue(statusHeader);
  return newColumn;
}

function onFormSubmitTrigger(e) {
  // 1. Get the data from the form submission
  var responses = e.namedValues;
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  
  var requesterName = String(getSubmissionValue_(e, responses, ['Requester Name'])).trim();
  var requesterEmail = getRequesterEmail_(e, responses);
  var building = String(getSubmissionValue_(e, responses, ['Building Location'])).trim();
  var ward = String(getSubmissionValue_(e, responses, ["Requester's Ward"])).trim();

  // --- HUMAN READABLE TIME FIX ---
  // We resolve timestamps from namedValues first, then fall back to the row values.
  var rawStartValue = getSubmissionValue_(e, responses, ['Access Start (Date & Time)']);
  var rawEndValue = getSubmissionValue_(e, responses, ['Access End (Date & Time)']);
  var rawStart = new Date(rawStartValue);
  var rawEnd = new Date(rawEndValue);

  if (isNaN(rawStart.getTime()) || isNaN(rawEnd.getTime())) {
    throw new Error('Access start/end date is missing or invalid in the submission row.');
  }
  
  // Format: M/d/yyyy h:mm a (e.g., 3/23/2026 1:01 PM)
  var start = Utilities.formatDate(rawStart, Session.getScriptTimeZone(), "M/d/yyyy h:mm a");
  var end = Utilities.formatDate(rawEnd, Session.getScriptTimeZone(), "M/d/yyyy h:mm a");

  var statusColumn = getOrCreateStatusColumn_(sheet);
  var existingStatus = String(sheet.getRange(row, statusColumn).getValue() || '').trim();
  var isUpdatedSubmission = existingStatus !== '';

  // 2. Map wards to bishop emails via Script Properties.
  var bishopEmails = getBishopEmails_();
  var targetBishop = bishopEmails[ward] || getSecret_('STAKE_TECHNOLOGY_SPECIALIST_EMAIL');

  // 3. Email 1: The Bishop's FYI
  var bishopSubject = isUpdatedSubmission
    ? "FYI: Updated Building Access Request - " + ward
    : "FYI: Building Access Request - " + ward;
  var bishopBody = "Bishop,\n\nThis is an automated notification that a building access request has been " +
                   (isUpdatedSubmission ? "updated" : "scheduled") +
                   " by a member of your ward.\n\n" +
                   "Requester: " + requesterName + "\n" +
                   "Building: " + building + "\n" +
                   "Time: " + start + " to " + end + "\n\n" +
                   "No action is required on your part. This has been vetted by the Building Scheduler.";
  
  MailApp.sendEmail(targetBishop, bishopSubject, bishopBody);

  // 4. Email 2: The Requester's Confirmation
  var requesterSubject = isUpdatedSubmission
    ? "Kindoo Access Updated: " + building
    : "Kindoo Access Confirmed: " + building;
  var requesterBody = "Hello " + requesterName + ",\n\n" +
                      "Your access request for the " + building + " has been " +
                      (isUpdatedSubmission ? "updated" : "completed") +
                      ".\n\n" +
                      "Scheduled Time: " + start + " to " + end + "\n\n" +
                      "IMPORTANT: Your digital key will be assigned to your Kindoo app a few days before the event begins. " +
                      "Please ensure your Kindoo app email matches your Member Tools email (" + requesterEmail + ").";

  MailApp.sendEmail(requesterEmail, requesterSubject, requesterBody);

  // 5. Ledger Documentation
  sheet.getRange(row, statusColumn).setValue(
    isUpdatedSubmission ? "Updated and Scheduled" : "Vetted and Scheduled"
  );
}
