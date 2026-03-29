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

function getStakeManagerEmails_() {
  var raw = getSecret_('STAKE_MANAGER_EMAILS');
  return raw.split(',').map(function(email) {
    return email.trim();
  }).filter(function(email) {
    return email !== '';
  });
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

function getLedgerSheet_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = PropertiesService.getScriptProperties().getProperty('LEDGER_SHEET_NAME');

  if (sheetName) {
    var namedSheet = spreadsheet.getSheetByName(sheetName);
    if (!namedSheet) {
      throw new Error('Missing ledger sheet: ' + sheetName);
    }
    return namedSheet;
  }

  return spreadsheet.getSheets()[0];
}

function getHeaderMap_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerMap = {};

  for (var i = 0; i < headers.length; i++) {
    if (headers[i]) {
      headerMap[String(headers[i]).trim()] = i + 1;
    }
  }

  return headerMap;
}

function getOrCreateColumnByHeader_(sheet, headerName) {
  var headerMap = getHeaderMap_(sheet);
  if (headerMap[headerName]) {
    return headerMap[headerName];
  }

  var newColumn = sheet.getLastColumn() + 1;
  sheet.getRange(1, newColumn).setValue(headerName);
  return newColumn;
}

function getOrCreateStatusColumn_(sheet) {
  return getOrCreateColumnByHeader_(sheet, 'Status');
}

function getCellString_(rowValues, headerMap, headerName) {
  var column = headerMap[headerName];
  if (!column) {
    return '';
  }

  var value = rowValues[column - 1];
  return value == null ? '' : String(value).trim();
}

function getCellValue_(rowValues, headerMap, headerName) {
  var column = headerMap[headerName];
  if (!column) {
    return null;
  }

  return rowValues[column - 1];
}

function parseSheetDate_(value) {
  if (!value) {
    return null;
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }

  var parsed = new Date(value);
  if (isNaN(parsed.getTime())) {
    return null;
  }
  return parsed;
}

function isSameLocalDate_(left, right, timeZone) {
  return Utilities.formatDate(left, timeZone, 'yyyy-MM-dd') === Utilities.formatDate(right, timeZone, 'yyyy-MM-dd');
}

function shouldSendManagerAlert_(rowValues, headerMap, now, timeZone) {
  var status = getCellString_(rowValues, headerMap, 'Status');
  var claimStatus = getCellString_(rowValues, headerMap, 'Manager Claim Status');
  var lastAlertSentAt = parseSheetDate_(getCellValue_(rowValues, headerMap, 'Manager Alert Last Sent At'));
  var startDate = parseSheetDate_(getCellValue_(rowValues, headerMap, 'Access Start (Date & Time)'));

  if (!startDate) {
    return false;
  }

  if (claimStatus.toLowerCase() === 'claimed') {
    return false;
  }

  if (status !== 'Vetted and Scheduled' && status !== 'Updated and Scheduled') {
    return false;
  }

  if (startDate.getTime() <= now.getTime()) {
    return false;
  }

  var sevenDaysFromNow = new Date(now.getTime() + (7 * 24 * 60 * 60 * 1000));
  if (startDate.getTime() > sevenDaysFromNow.getTime()) {
    return false;
  }

  if (lastAlertSentAt && isSameLocalDate_(lastAlertSentAt, now, timeZone)) {
    return false;
  }

  return true;
}

function sendStakeManagerAlert_(details) {
  var timeZone = Session.getScriptTimeZone();
  var subject = 'Kindoo Access Needs Assignment: ' + details.building + ' - ' + details.requesterName;
  var body = 'Stake Managers,\n\n' +
             'A Kindoo access request is within the next 7 days and needs to be claimed.\n\n' +
             'Requester: ' + details.requesterName + '\n' +
             'Requester Email: ' + details.requesterEmail + '\n' +
             'Requester Phone: ' + details.requesterPhone + '\n' +
             'Building: ' + details.building + '\n' +
             'Ward: ' + details.ward + '\n' +
             'Access Start: ' + Utilities.formatDate(details.startDate, timeZone, 'M/d/yyyy h:mm a') + '\n' +
             'Access End: ' + Utilities.formatDate(details.endDate, timeZone, 'M/d/yyyy h:mm a') + '\n' +
             'Ledger Row: ' + details.row + '\n\n' +
             'This request will continue to alert daily until it is marked Claimed in the ledger.';

  MailApp.sendEmail(getStakeManagerEmails_().join(','), subject, body);
}

function runUpcomingAccessScan() {
  var sheet = getLedgerSheet_();
  var statusColumn = getOrCreateStatusColumn_(sheet);
  var claimStatusColumn = getOrCreateColumnByHeader_(sheet, 'Manager Claim Status');
  var alertStatusColumn = getOrCreateColumnByHeader_(sheet, 'Manager Alert Status');
  var alertLastSentColumn = getOrCreateColumnByHeader_(sheet, 'Manager Alert Last Sent At');
  var alertCountColumn = getOrCreateColumnByHeader_(sheet, 'Manager Alert Count');
  var headerMap = getHeaderMap_(sheet);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow < 2) {
    return;
  }

  var rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  var now = new Date();
  var timeZone = Session.getScriptTimeZone();

  for (var i = 0; i < rows.length; i++) {
    var rowValues = rows[i];
    var rowNumber = i + 2;

    if (!shouldSendManagerAlert_(rowValues, headerMap, now, timeZone)) {
      continue;
    }

    var details = {
      row: rowNumber,
      requesterName: getCellString_(rowValues, headerMap, 'Requester Name'),
      requesterEmail: getCellString_(rowValues, headerMap, 'Requester Email'),
      requesterPhone: getCellString_(rowValues, headerMap, 'Requester Phone Number'),
      building: getCellString_(rowValues, headerMap, 'Building Location'),
      ward: getCellString_(rowValues, headerMap, "Requester's Ward"),
      startDate: parseSheetDate_(getCellValue_(rowValues, headerMap, 'Access Start (Date & Time)')),
      endDate: parseSheetDate_(getCellValue_(rowValues, headerMap, 'Access End (Date & Time)'))
    };

    sendStakeManagerAlert_(details);

    var existingAlertCount = parseInt(sheet.getRange(rowNumber, alertCountColumn).getValue(), 10);
    sheet.getRange(rowNumber, claimStatusColumn).setValue(
      String(sheet.getRange(rowNumber, claimStatusColumn).getValue() || '').trim() || 'Unclaimed'
    );
    sheet.getRange(rowNumber, alertStatusColumn).setValue('Alerted');
    sheet.getRange(rowNumber, alertLastSentColumn).setValue(now);
    sheet.getRange(rowNumber, alertCountColumn).setValue(isNaN(existingAlertCount) ? 1 : existingAlertCount + 1);
    sheet.getRange(rowNumber, statusColumn).setValue(
      String(sheet.getRange(rowNumber, statusColumn).getValue() || '').trim()
    );
  }
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
