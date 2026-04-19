/**
 * Kindoo notification, ledger, trigger, and web-app workflow.
 * This script is intended to run from the spreadsheet-bound Apps Script project.
 */

// Reads a required Script Property and fails fast with a clear message when it
// has not been configured in the live Apps Script project.
function getSecret_(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error('Missing script property: ' + key);
  }
  return value;
}

// Returns the ward-to-bishop mapping from Script Properties so bishop emails
// stay out of source control.
function getBishopEmails_() {
  return {
    '1st Ward': getSecret_('WARD_1_EMAIL'),
    '2nd Ward': getSecret_('WARD_2_EMAIL'),
    '4th Ward': getSecret_('WARD_4_EMAIL'),
    '5th Ward': getSecret_('WARD_5_EMAIL'),
    '7th Ward': getSecret_('WARD_7_EMAIL')
  };
}

// Parses the comma-separated stake manager email list from Script Properties.
function getStakeManagerEmails_() {
  var raw = getSecret_('STAKE_MANAGER_EMAILS');
  return raw.split(',').map(function(email) {
    return email.trim();
  }).filter(function(email) {
    return email !== '';
  });
}

// Returns the secret used to sign claim links.
function getClaimLinkSecret_() {
  return getSecret_('CLAIM_LINK_SECRET');
}

// Returns the secret used to sign issued links.
function getIssuedLinkSecret_() {
  return getSecret_('ISSUED_LINK_SECRET');
}

// Normalizes a namedValues entry from the form submit event into one trimmed
// string, or returns an empty string when the event value is missing.
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

// Reads a response value directly from the submitted sheet row when namedValues
// are incomplete, such as after an edited submission.
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

// Resolves one submission field by trying the form event first and then the
// actual spreadsheet row as a fallback.
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

// Resolves the requester email from the submission and throws a clear error if
// the response does not contain one.
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

// Opens the ledger spreadsheet using an explicit Script Property when present,
// or falls back to the active spreadsheet in bound-script scenarios.
function getLedgerSpreadsheet_() {
  var spreadsheetId = PropertiesService.getScriptProperties().getProperty('LEDGER_SPREADSHEET_ID');

  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet) {
    return activeSpreadsheet;
  }

  throw new Error('Missing ledger spreadsheet. Set Script Property LEDGER_SPREADSHEET_ID.');
}

// Returns the active ledger tab by name when configured, otherwise the first
// sheet in the ledger spreadsheet.
function getLedgerSheet_() {
  var spreadsheet = getLedgerSpreadsheet_();
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

// Removes existing installable triggers for one handler so trigger setup can be
// safely re-run without creating duplicates.
function deleteTriggersByHandler_(handlerName) {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// Creates the spreadsheet form-submit trigger that drives submit-time emails
// and ledger updates.
function createKindooSpreadsheetTrigger() {
  var spreadsheet = getLedgerSpreadsheet_();

  deleteTriggersByHandler_('onFormSubmitTrigger');

  ScriptApp.newTrigger('onFormSubmitTrigger')
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();

  Logger.log('Created spreadsheet form-submit trigger for spreadsheet: ' + spreadsheet.getId());
}

// Creates the once-daily trigger that scans upcoming requests for manager
// action.
function createKindooDailyScanTrigger() {
  deleteTriggersByHandler_('runUpcomingAccessScan');

  ScriptApp.newTrigger('runUpcomingAccessScan')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('Created daily time-driven trigger for runUpcomingAccessScan().');
}

// Convenience entrypoint that recreates both required installable triggers.
function createKindooTriggers() {
  createKindooSpreadsheetTrigger();
  createKindooDailyScanTrigger();
  Logger.log('Created Kindoo triggers successfully.');
}

// Builds a leftmost-wins header map so duplicate helper columns do not confuse
// later lookups.
function getHeaderMap_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerMap = {};

  for (var i = 0; i < headers.length; i++) {
    if (headers[i] && !headerMap[String(headers[i]).trim()]) {
      headerMap[String(headers[i]).trim()] = i + 1;
    }
  }

  return headerMap;
}

// Returns every column index whose header matches the requested name.
function getColumnsByHeader_(sheet, headerName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var matches = [];

  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim() === headerName) {
      matches.push(i + 1);
    }
  }

  return matches;
}

// Reuses the first matching helper column or creates it if the header does not
// exist yet.
function getOrCreateColumnByHeader_(sheet, headerName) {
  var matchingColumns = getColumnsByHeader_(sheet, headerName);
  if (matchingColumns.length > 0) {
    return matchingColumns[0];
  }

  var newColumn = sheet.getLastColumn() + 1;
  sheet.getRange(1, newColumn).setValue(headerName);
  return newColumn;
}

// Returns the canonical status column.
function getOrCreateStatusColumn_(sheet) {
  return getOrCreateColumnByHeader_(sheet, 'Status');
}

// Returns the canonical request ID column.
function getOrCreateRequestIdColumn_(sheet) {
  return getOrCreateColumnByHeader_(sheet, 'Request ID');
}

// Reads one cell as a trimmed string using the header map.
function getCellString_(rowValues, headerMap, headerName) {
  var column = headerMap[headerName];
  if (!column) {
    return '';
  }

  var value = rowValues[column - 1];
  return value == null ? '' : String(value).trim();
}

// Reads one raw cell value using the header map.
function getCellValue_(rowValues, headerMap, headerName) {
  var column = headerMap[headerName];
  if (!column) {
    return null;
  }

  return rowValues[column - 1];
}

// Safely parses sheet values into a Date or returns null when parsing fails.
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

// Compares two dates by local date only in the script timezone.
function isSameLocalDate_(left, right, timeZone) {
  return Utilities.formatDate(left, timeZone, 'yyyy-MM-dd') === Utilities.formatDate(right, timeZone, 'yyyy-MM-dd');
}

// Converts the HMAC byte array into a lowercase hex string for URLs.
function toHexString_(bytes) {
  return bytes.map(function(byteValue) {
    var normalized = byteValue < 0 ? byteValue + 256 : byteValue;
    var hex = normalized.toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

// Normalizes email addresses for token generation and equality checks.
function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

// Builds a signed token for a specific action, request, and actor email.
function buildSignedToken_(action, requestId, actorEmail, secret) {
  var signature = Utilities.computeHmacSha256Signature(
    action + ':' + requestId + ':' + normalizeEmail_(actorEmail),
    secret
  );
  return toHexString_(signature);
}

// Creates the signed token used by claim links.
function buildClaimToken_(requestId, actorEmail) {
  return buildSignedToken_('claim', requestId, actorEmail, getClaimLinkSecret_());
}

// Creates the signed token used by issued links.
function buildIssuedToken_(requestId, actorEmail) {
  return buildSignedToken_('issued', requestId, actorEmail, getIssuedLinkSecret_());
}

// Returns the deployed web-app URL from Script Properties, with Apps Script's
// service URL as a fallback.
function getClaimWebAppUrl_() {
  var configuredUrl = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL');
  if (configuredUrl) {
    return configuredUrl;
  }

  var deployedUrl = ScriptApp.getService().getUrl();
  return deployedUrl || '';
}

// Builds the personalized claim URL for one manager email.
function buildClaimUrl_(requestId, actorEmail) {
  var baseUrl = getClaimWebAppUrl_();
  if (!baseUrl) {
    return '';
  }

  return baseUrl +
    '?action=claim' +
    '&requestId=' + encodeURIComponent(requestId) +
    '&actor=' + encodeURIComponent(normalizeEmail_(actorEmail)) +
    '&token=' + encodeURIComponent(buildClaimToken_(requestId, actorEmail));
}

// Builds the personalized issued URL for the claiming manager.
function buildIssuedUrl_(requestId, actorEmail) {
  var baseUrl = getClaimWebAppUrl_();
  if (!baseUrl) {
    return '';
  }

  return baseUrl +
    '?action=issued' +
    '&requestId=' + encodeURIComponent(requestId) +
    '&actor=' + encodeURIComponent(normalizeEmail_(actorEmail)) +
    '&token=' + encodeURIComponent(buildIssuedToken_(requestId, actorEmail));
}

// Finds the next numeric request sequence by scanning existing request IDs.
function findNextRequestSequence_(sheet, requestIdColumn) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 1;
  }

  var values = sheet.getRange(2, requestIdColumn, lastRow - 1, 1).getValues();
  var maxNumber = 0;

  for (var i = 0; i < values.length; i++) {
    var match = String(values[i][0] || '').trim().match(/^REQ-(\d+)$/);
    if (!match) {
      continue;
    }

    var parsed = parseInt(match[1], 10);
    if (!isNaN(parsed) && parsed > maxNumber) {
      maxNumber = parsed;
    }
  }

  return maxNumber + 1;
}

// Ensures the given ledger row has a stable request ID and returns it.
function ensureRequestIdForRow_(sheet, row) {
  var requestIdColumn = getOrCreateRequestIdColumn_(sheet);
  var existingRequestId = String(sheet.getRange(row, requestIdColumn).getValue() || '').trim();

  if (existingRequestId) {
    return existingRequestId;
  }

  var requestId = 'REQ-' + ('0000' + findNextRequestSequence_(sheet, requestIdColumn)).slice(-4);
  sheet.getRange(row, requestIdColumn).setValue(requestId);
  return requestId;
}

// One-time maintenance helper that merges duplicate Request ID columns into the
// leftmost canonical column and removes the extras.
function cleanupDuplicateRequestIdColumns() {
  var sheet = getLedgerSheet_();
  var requestIdColumns = getColumnsByHeader_(sheet, 'Request ID');

  if (requestIdColumns.length <= 1) {
    Logger.log('No duplicate Request ID columns found on sheet: ' + sheet.getName());
    return;
  }

  var canonicalColumn = requestIdColumns[0];
  var duplicateColumns = requestIdColumns.slice(1);
  var lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    var canonicalValues = sheet.getRange(2, canonicalColumn, lastRow - 1, 1).getValues();

    for (var i = 0; i < duplicateColumns.length; i++) {
      var duplicateColumn = duplicateColumns[i];
      var duplicateValues = sheet.getRange(2, duplicateColumn, lastRow - 1, 1).getValues();

      for (var rowIndex = 0; rowIndex < duplicateValues.length; rowIndex++) {
        var canonicalValue = String(canonicalValues[rowIndex][0] || '').trim();
        var duplicateValue = String(duplicateValues[rowIndex][0] || '').trim();

        if (!canonicalValue && duplicateValue) {
          canonicalValues[rowIndex][0] = duplicateValue;
        }
      }
    }

    sheet.getRange(2, canonicalColumn, lastRow - 1, 1).setValues(canonicalValues);
  }

  for (var j = duplicateColumns.length - 1; j >= 0; j--) {
    sheet.deleteColumn(duplicateColumns[j]);
  }

  Logger.log(
    'Merged duplicate Request ID columns into column ' +
    canonicalColumn +
    ' and removed ' +
    duplicateColumns.length +
    ' duplicate column(s) on sheet ' +
    sheet.getName()
  );
}

// Determines whether a row should trigger an alert or reminder during the
// current scan run.
function shouldSendManagerAlert_(rowValues, headerMap, now, timeZone) {
  var status = getCellString_(rowValues, headerMap, 'Status');
  var claimStatus = getCellString_(rowValues, headerMap, 'Manager Claim Status');
  var claimedBy = getCellString_(rowValues, headerMap, 'Claimed By');
  var keyStatus = getCellString_(rowValues, headerMap, 'Kindoo Key Status');
  var lastAlertSentAt = parseSheetDate_(getCellValue_(rowValues, headerMap, 'Manager Alert Last Sent At'));
  var startDate = parseSheetDate_(getCellValue_(rowValues, headerMap, 'Access Start (Date & Time)'));

  if (!startDate) {
    return false;
  }

  if (keyStatus.toLowerCase() === 'issued') {
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

  if (claimStatus.toLowerCase() === 'claimed' && !claimedBy) {
    return false;
  }

  return true;
}

// Sends the initial daily assignment alert to every configured stake manager.
function sendStakeManagerAlert_(details) {
  var timeZone = Session.getScriptTimeZone();
  var formattedStart = Utilities.formatDate(details.startDate, timeZone, 'M/d/yyyy h:mm a');
  var formattedEnd = Utilities.formatDate(details.endDate, timeZone, 'M/d/yyyy h:mm a');
  var managerEmails = getStakeManagerEmails_();

  for (var i = 0; i < managerEmails.length; i++) {
    var managerEmail = managerEmails[i];
    var claimUrl = buildClaimUrl_(details.requestId, managerEmail);
    var subject = 'Kindoo Access Needs Assignment: ' + details.building + ' - ' + details.requesterName;
    var body = 'Stake Kindoo Managers,\n\n' +
               'A Kindoo access request is within the next 7 days and needs to be claimed.\n\n' +
               'Once you claim it, the other Stake Kindoo Managers will be notified that you have taken ownership. At that point, it becomes your responsibility to schedule the Kindoo access request. You will continue to receive a daily reminder until the key is actually scheduled.\n\n' +
               'Claim this request: ' + (claimUrl || 'Publish the web app and set WEB_APP_URL to enable claims.') + '\n\n' +
               'Request ID: ' + details.requestId + '\n' +
               'Requester: ' + details.requesterName + '\n' +
               'Requester Email: ' + details.requesterEmail + '\n' +
               'Requester Phone: ' + details.requesterPhone + '\n' +
               'Building: ' + details.building + '\n' +
               'Ward: ' + details.ward + '\n' +
               'Access Start: ' + formattedStart + '\n' +
               'Access End: ' + formattedEnd + '\n' +
               'Ledger Row: ' + details.row + '\n\n' +
               'This request will continue to alert daily until it is claimed.';
    var htmlBody = '<p>Stake Kindoo Managers,</p>' +
                   '<p>A Kindoo access request is within the next 7 days and needs to be claimed.</p>' +
                   '<p>Once you claim it, the other Stake Kindoo Managers will be notified that you have taken ownership. At that point, it becomes your responsibility to schedule the Kindoo access request. You will continue to receive a daily reminder until the key is actually scheduled.</p>' +
                   (claimUrl
                     ? '<p><a href="' + claimUrl + '" style="display:inline-block;padding:10px 16px;background:#1a73e8;color:#ffffff;text-decoration:none;border-radius:4px;">Claim this request</a></p>'
                     : '<p><strong>Claim link unavailable.</strong> Publish the web app and set <code>WEB_APP_URL</code> to enable claims.</p>') +
                   '<p><strong>Request ID:</strong> ' + details.requestId + '<br>' +
                   '<strong>Requester:</strong> ' + details.requesterName + '<br>' +
                   '<strong>Requester Email:</strong> ' + details.requesterEmail + '<br>' +
                   '<strong>Requester Phone:</strong> ' + details.requesterPhone + '<br>' +
                   '<strong>Building:</strong> ' + details.building + '<br>' +
                   '<strong>Ward:</strong> ' + details.ward + '<br>' +
                   '<strong>Access Start:</strong> ' + formattedStart + '<br>' +
                   '<strong>Access End:</strong> ' + formattedEnd + '<br>' +
                   '<strong>Ledger Row:</strong> ' + details.row + '</p>' +
                   '<p>This request will continue to alert daily until it is claimed.</p>';

    MailApp.sendEmail({
      to: managerEmail,
      subject: subject,
      body: body,
      htmlBody: htmlBody
    });
  }
}

// Sends the once-per-day reminder to the claiming manager until the key is
// actually issued.
function sendClaimedManagerReminder_(claimerEmail, details) {
  if (!claimerEmail || claimerEmail.indexOf('@') === -1) {
    return;
  }

  var timeZone = Session.getScriptTimeZone();
  var formattedStart = Utilities.formatDate(details.startDate, timeZone, 'M/d/yyyy h:mm a');
  var formattedEnd = Utilities.formatDate(details.endDate, timeZone, 'M/d/yyyy h:mm a');
  var issueUrl = buildIssuedUrl_(details.requestId, claimerEmail);
  var subject = 'Reminder: Schedule Kindoo Key for ' + details.requestId;
  var body = 'This is a reminder that you claimed request ' + details.requestId + ' and still need to schedule the Kindoo access key.\n\n' +
             'Requester: ' + details.requesterName + '\n' +
             'Requester Email: ' + details.requesterEmail + '\n' +
             'Requester Phone: ' + details.requesterPhone + '\n' +
             'Building: ' + details.building + '\n' +
             'Ward: ' + details.ward + '\n' +
             'Access Start: ' + formattedStart + '\n' +
             'Access End: ' + formattedEnd + '\n\n' +
             'Only click the issued button after you have actually scheduled the Kindoo access key.\n\n' +
             'When the Kindoo key has been issued, click:\n' + issueUrl;
  var htmlBody = '<p>This is a reminder that you claimed request <strong>' + details.requestId + '</strong> and still need to schedule the Kindoo access key.</p>' +
                 '<p><strong>Requester:</strong> ' + details.requesterName + '<br>' +
                 '<strong>Requester Email:</strong> ' + details.requesterEmail + '<br>' +
                 '<strong>Requester Phone:</strong> ' + details.requesterPhone + '<br>' +
                 '<strong>Building:</strong> ' + details.building + '<br>' +
                 '<strong>Ward:</strong> ' + details.ward + '<br>' +
                 '<strong>Access Start:</strong> ' + formattedStart + '<br>' +
                 '<strong>Access End:</strong> ' + formattedEnd + '</p>' +
                 '<p><strong>Only click the button below after you have actually scheduled the Kindoo access key.</strong></p>' +
                 '<p><a href="' + issueUrl + '" style="display:inline-block;padding:10px 16px;background:#188038;color:#ffffff;text-decoration:none;border-radius:4px;">Kindoo Key Issued</a></p>';

  MailApp.sendEmail({
    to: claimerEmail,
    subject: subject,
    body: body,
    htmlBody: htmlBody
  });
}

// Builds a simple branded HTML response for claim and issued web-app actions.
function buildClaimResponseHtml_(title, body, accentColor) {
  var color = accentColor || '#1a73e8';
  var html = '<html><body style="font-family:Arial,sans-serif;padding:24px;line-height:1.5;">' +
             '<div style="max-width:800px;margin:0 auto;border-left:8px solid ' + color + ';padding-left:20px;">' +
             '<h2>' + title + '</h2>' +
             '<p>' + body + '</p>' +
             '</div></body></html>';

  return HtmlService.createHtmlOutput(html).setTitle(title);
}

// Reads the requester-facing details used in manager emails and final success
// notifications for one ledger row.
function getRequesterDetailsForRow_(sheet, row) {
  var headerMap = getHeaderMap_(sheet);
  var rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  return {
    requesterName: getCellString_(rowValues, headerMap, 'Requester Name'),
    requesterEmail: getCellString_(rowValues, headerMap, 'Requester Email'),
    requesterPhone: getCellString_(rowValues, headerMap, 'Requester Phone Number'),
    building: getCellString_(rowValues, headerMap, 'Building Location'),
    ward: getCellString_(rowValues, headerMap, "Requester's Ward"),
    startDate: parseSheetDate_(getCellValue_(rowValues, headerMap, 'Access Start (Date & Time)')),
    endDate: parseSheetDate_(getCellValue_(rowValues, headerMap, 'Access End (Date & Time)'))
  };
}

// Sends the immediate follow-up email to the claiming manager with the issued
// button.
function sendKeyIssuancePromptEmail_(claimerEmail, requestId, details) {
  if (!claimerEmail || claimerEmail.indexOf('@') === -1) {
    return;
  }

  var timeZone = Session.getScriptTimeZone();
  var issueUrl = buildIssuedUrl_(requestId, claimerEmail);
  var subject = 'Kindoo Key Assignment Needed: ' + requestId;
  var body = 'You claimed request ' + requestId + '.\n\n' +
             'Next step: schedule the Kindoo access key for this request.\n\n' +
             'Only click the issued button after you have actually scheduled the Kindoo access key.\n\n' +
             'Requester: ' + details.requesterName + '\n' +
             'Building: ' + details.building + '\n' +
             'Ward: ' + details.ward + '\n' +
             'Access Start: ' + Utilities.formatDate(details.startDate, timeZone, 'M/d/yyyy h:mm a') + '\n' +
             'Access End: ' + Utilities.formatDate(details.endDate, timeZone, 'M/d/yyyy h:mm a') + '\n\n' +
             'After you issue the Kindoo key, click this link:\n' + issueUrl;
  var htmlBody = '<p>You claimed request <strong>' + requestId + '</strong>.</p>' +
                 '<p><strong>Next step:</strong> schedule the Kindoo access key for this request.</p>' +
                 '<p><strong>Only click the button below after you have actually scheduled the Kindoo access key.</strong></p>' +
                 '<p><strong>Requester:</strong> ' + details.requesterName + '<br>' +
                 '<strong>Building:</strong> ' + details.building + '<br>' +
                 '<strong>Ward:</strong> ' + details.ward + '<br>' +
                 '<strong>Access Start:</strong> ' + Utilities.formatDate(details.startDate, timeZone, 'M/d/yyyy h:mm a') + '<br>' +
                 '<strong>Access End:</strong> ' + Utilities.formatDate(details.endDate, timeZone, 'M/d/yyyy h:mm a') + '</p>' +
                 '<p><a href="' + issueUrl + '" style="display:inline-block;padding:10px 16px;background:#188038;color:#ffffff;text-decoration:none;border-radius:4px;">Kindoo Key Issued</a></p>';

  MailApp.sendEmail({
    to: claimerEmail,
    subject: subject,
    body: body,
    htmlBody: htmlBody
  });
}

// Notifies the non-claiming managers that the key has been issued and no
// further action is needed from them.
function sendIssuedNotificationToOtherManagers_(issuerEmail, requestId, details) {
  var recipients = getStakeManagerEmails_().filter(function(email) {
    return email && email.toLowerCase() !== String(issuerEmail || '').toLowerCase();
  });

  if (recipients.length === 0) {
    return;
  }

  var timeZone = Session.getScriptTimeZone();
  var issuerLabel = issuerEmail || 'A manager';
  var subject = 'Kindoo Key Issued: ' + requestId;
  var body = issuerLabel + ' has issued the Kindoo key for request ' + requestId + '.\n\n' +
             'Requester: ' + details.requesterName + '\n' +
             'Building: ' + details.building + '\n' +
             'Ward: ' + details.ward + '\n' +
             'Access Start: ' + Utilities.formatDate(details.startDate, timeZone, 'M/d/yyyy h:mm a') + '\n' +
             'Access End: ' + Utilities.formatDate(details.endDate, timeZone, 'M/d/yyyy h:mm a');
  var htmlBody = '<p><strong>' + issuerLabel + '</strong> has issued the Kindoo key for request <strong>' + requestId + '</strong>.</p>' +
                 '<p><strong>Requester:</strong> ' + details.requesterName + '<br>' +
                 '<strong>Building:</strong> ' + details.building + '<br>' +
                 '<strong>Ward:</strong> ' + details.ward + '<br>' +
                 '<strong>Access Start:</strong> ' + Utilities.formatDate(details.startDate, timeZone, 'M/d/yyyy h:mm a') + '<br>' +
                 '<strong>Access End:</strong> ' + Utilities.formatDate(details.endDate, timeZone, 'M/d/yyyy h:mm a') + '</p>';

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    body: body,
    htmlBody: htmlBody
  });
}

// Sends the final success email to the member once the key has actually been
// issued.
function sendFinalSuccessEmailToMember_(details) {
  if (!details.requesterEmail || details.requesterEmail.indexOf('@') === -1) {
    return;
  }

  var timeZone = Session.getScriptTimeZone();
  var subject = 'Kindoo Access Confirmed: ' + details.building;
  var body = 'Hello ' + details.requesterName + ',\n\n' +
             'Your Kindoo key access request for the ' + details.building + ' has been completed.\n\n' +
             'Scheduled Time: ' + Utilities.formatDate(details.startDate, timeZone, 'M/d/yyyy h:mm a') + ' to ' +
             Utilities.formatDate(details.endDate, timeZone, 'M/d/yyyy h:mm a') + '\n\n' +
             'Your Kindoo digital key has now been issued.';
  var htmlBody = '<p>Hello ' + details.requesterName + ',</p>' +
                 '<p>Your Kindoo key access request for the <strong>' + details.building + '</strong> has been completed.</p>' +
                 '<p><strong>Scheduled Time:</strong> ' +
                 Utilities.formatDate(details.startDate, timeZone, 'M/d/yyyy h:mm a') + ' to ' +
                 Utilities.formatDate(details.endDate, timeZone, 'M/d/yyyy h:mm a') + '</p>' +
                 '<p>Your Kindoo digital key has now been issued.</p>';

  MailApp.sendEmail({
    to: details.requesterEmail,
    subject: subject,
    body: body,
    htmlBody: htmlBody
  });
}

// Notifies the other managers once one manager has claimed the request, after
// which only the claimer keeps receiving reminders.
function sendClaimedNotificationToOtherManagers_(claimerEmail, requestId, details) {
  var recipients = getStakeManagerEmails_().filter(function(email) {
    return email && email.toLowerCase() !== String(claimerEmail || '').toLowerCase();
  });

  if (recipients.length === 0) {
    return;
  }

  var timeZone = Session.getScriptTimeZone();
  var claimerLabel = claimerEmail || 'A manager';
  var subject = 'Kindoo Request Claimed: ' + requestId;
  var body = claimerLabel + ' has claimed request ' + requestId + '.\n\n' +
             'Requester: ' + details.requesterName + '\n' +
             'Building: ' + details.building + '\n' +
             'Ward: ' + details.ward + '\n' +
             'Access Start: ' + Utilities.formatDate(details.startDate, timeZone, 'M/d/yyyy h:mm a') + '\n' +
             'Access End: ' + Utilities.formatDate(details.endDate, timeZone, 'M/d/yyyy h:mm a');
  var htmlBody = '<p><strong>' + claimerLabel + '</strong> has claimed request <strong>' + requestId + '</strong>.</p>' +
                 '<p><strong>Requester:</strong> ' + details.requesterName + '<br>' +
                 '<strong>Building:</strong> ' + details.building + '<br>' +
                 '<strong>Ward:</strong> ' + details.ward + '<br>' +
                 '<strong>Access Start:</strong> ' + Utilities.formatDate(details.startDate, timeZone, 'M/d/yyyy h:mm a') + '<br>' +
                 '<strong>Access End:</strong> ' + Utilities.formatDate(details.endDate, timeZone, 'M/d/yyyy h:mm a') + '</p>';

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    body: body,
    htmlBody: htmlBody
  });
}

// Finds a ledger row by its stable request ID.
function findRowByRequestId_(sheet, requestId, requestIdColumn) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }

  var values = sheet.getRange(2, requestIdColumn, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === requestId) {
      return i + 2;
    }
  }

  return -1;
}

// Handles the manager claim web-app action, records ownership in the ledger,
// and fans out the next notification emails.
function claimRequest_(requestId, actorEmail, token) {
  if (!requestId || !actorEmail || !token) {
    return buildClaimResponseHtml_('Claim Failed', 'The claim link is missing required information.', '#d93025');
  }

  if (token !== buildClaimToken_(requestId, actorEmail)) {
    return buildClaimResponseHtml_('Claim Failed', 'This claim link is invalid.', '#d93025');
  }

  var sheet = getLedgerSheet_();
  var requestIdColumn = getOrCreateRequestIdColumn_(sheet);
  var claimStatusColumn = getOrCreateColumnByHeader_(sheet, 'Manager Claim Status');
  var alertStatusColumn = getOrCreateColumnByHeader_(sheet, 'Manager Alert Status');
  var claimedByColumn = getOrCreateColumnByHeader_(sheet, 'Claimed By');
  var claimedAtColumn = getOrCreateColumnByHeader_(sheet, 'Claimed At');
  var row = findRowByRequestId_(sheet, requestId, requestIdColumn);

  if (row === -1) {
    return buildClaimResponseHtml_('Claim Failed', 'No ledger row was found for request ID ' + requestId + '.', '#d93025');
  }

  var existingClaimStatus = String(sheet.getRange(row, claimStatusColumn).getValue() || '').trim();
  var existingClaimedBy = String(sheet.getRange(row, claimedByColumn).getValue() || '').trim();
  var existingClaimedAt = parseSheetDate_(sheet.getRange(row, claimedAtColumn).getValue());
  var timeZone = Session.getScriptTimeZone();
  var details = getRequesterDetailsForRow_(sheet, row);

  if (existingClaimStatus.toLowerCase() === 'claimed') {
    var alreadyClaimedMessage = 'Request ' + requestId + ' was already claimed';
    if (existingClaimedBy) {
      alreadyClaimedMessage += ' by ' + existingClaimedBy;
    }
    if (existingClaimedAt) {
      alreadyClaimedMessage += ' on ' + Utilities.formatDate(existingClaimedAt, timeZone, 'M/d/yyyy h:mm a');
    }
    alreadyClaimedMessage += '.';

    return buildClaimResponseHtml_('Already Claimed', alreadyClaimedMessage);
  }

  var claimedBy = normalizeEmail_(actorEmail);
  var now = new Date();

  sheet.getRange(row, claimStatusColumn).setValue('Claimed');
  sheet.getRange(row, alertStatusColumn).setValue('Claimed');
  sheet.getRange(row, claimedByColumn).setValue(claimedBy);
  sheet.getRange(row, claimedAtColumn).setValue(now);
  sendKeyIssuancePromptEmail_(claimedBy, requestId, details);
  sendClaimedNotificationToOtherManagers_(claimedBy, requestId, details);

  return buildClaimResponseHtml_(
    'Request Claimed',
    'Request ' + requestId + ' has been claimed successfully by ' + claimedBy + '. Next step: schedule the Kindoo access key. A follow-up email has been sent with a button to mark the key as issued.',
    '#188038'
  );
}

// Handles the issued web-app action, marks the ledger issued, and sends final
// notifications.
function markRequestIssued_(requestId, actorEmail, token) {
  if (!requestId || !actorEmail || !token) {
    return buildClaimResponseHtml_('Issue Update Failed', 'The issued link is missing required information.', '#d93025');
  }

  if (token !== buildIssuedToken_(requestId, actorEmail)) {
    return buildClaimResponseHtml_('Issue Update Failed', 'This issued link is invalid.', '#d93025');
  }

  var sheet = getLedgerSheet_();
  var requestIdColumn = getOrCreateRequestIdColumn_(sheet);
  var claimStatusColumn = getOrCreateColumnByHeader_(sheet, 'Manager Claim Status');
  var alertStatusColumn = getOrCreateColumnByHeader_(sheet, 'Manager Alert Status');
  var issuedByColumn = getOrCreateColumnByHeader_(sheet, 'Issued By');
  var issuedAtColumn = getOrCreateColumnByHeader_(sheet, 'Issued At');
  var keyStatusColumn = getOrCreateColumnByHeader_(sheet, 'Kindoo Key Status');
  var row = findRowByRequestId_(sheet, requestId, requestIdColumn);

  if (row === -1) {
    return buildClaimResponseHtml_('Issue Update Failed', 'No ledger row was found for request ID ' + requestId + '.', '#d93025');
  }

  var existingIssuedAt = parseSheetDate_(sheet.getRange(row, issuedAtColumn).getValue());
  var existingIssuedBy = String(sheet.getRange(row, issuedByColumn).getValue() || '').trim();
  var issuerEmail = normalizeEmail_(actorEmail) || existingIssuedBy || String(sheet.getRange(row, getOrCreateColumnByHeader_(sheet, 'Claimed By')).getValue() || '').trim();
  var details = getRequesterDetailsForRow_(sheet, row);

  if (existingIssuedAt) {
    var alreadyIssuedMessage = 'Request ' + requestId + ' was already marked issued';
    if (existingIssuedBy) {
      alreadyIssuedMessage += ' by ' + existingIssuedBy;
    }
    alreadyIssuedMessage += ' on ' + Utilities.formatDate(existingIssuedAt, Session.getScriptTimeZone(), 'M/d/yyyy h:mm a') + '.';
    return buildClaimResponseHtml_('Already Issued', alreadyIssuedMessage, '#d93025');
  }

  var now = new Date();
  sheet.getRange(row, claimStatusColumn).setValue('Claimed');
  sheet.getRange(row, alertStatusColumn).setValue('Issued');
  sheet.getRange(row, keyStatusColumn).setValue('Issued');
  sheet.getRange(row, issuedByColumn).setValue(issuerEmail || 'Issued via web app');
  sheet.getRange(row, issuedAtColumn).setValue(now);

  sendIssuedNotificationToOtherManagers_(issuerEmail, requestId, details);
  sendFinalSuccessEmailToMember_(details);

  return buildClaimResponseHtml_(
    'Kindoo Key Issued',
    'Request ' + requestId + ' has been marked as issued. The member and the other managers have been notified.',
    '#188038'
  );
}

// Web-app entrypoint for claim and issued actions.
function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : '';

  if (action === 'claim') {
    return claimRequest_(e.parameter.requestId, e.parameter.actor, e.parameter.token);
  }

  if (action === 'issued') {
    return markRequestIssued_(e.parameter.requestId, e.parameter.actor, e.parameter.token);
  }

  return buildClaimResponseHtml_('Kindoo Claim App', 'This web app is running. Use a claim link from a manager alert email.');
}

// Daily scanner that routes unclaimed requests to all managers and claimed
// requests to the claiming manager only until issuance.
function runUpcomingAccessScan() {
  var sheet = getLedgerSheet_();
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

    var requestId = ensureRequestIdForRow_(sheet, rowNumber);
    var claimStatus = getCellString_(rowValues, headerMap, 'Manager Claim Status');
    var claimedBy = getCellString_(rowValues, headerMap, 'Claimed By');
    var details = {
      row: rowNumber,
      requestId: requestId,
      requesterName: getCellString_(rowValues, headerMap, 'Requester Name'),
      requesterEmail: getCellString_(rowValues, headerMap, 'Requester Email'),
      requesterPhone: getCellString_(rowValues, headerMap, 'Requester Phone Number'),
      building: getCellString_(rowValues, headerMap, 'Building Location'),
      ward: getCellString_(rowValues, headerMap, "Requester's Ward"),
      startDate: parseSheetDate_(getCellValue_(rowValues, headerMap, 'Access Start (Date & Time)')),
      endDate: parseSheetDate_(getCellValue_(rowValues, headerMap, 'Access End (Date & Time)'))
    };

    if (claimStatus.toLowerCase() === 'claimed' && claimedBy) {
      sendClaimedManagerReminder_(claimedBy, details);
    } else {
      sendStakeManagerAlert_(details);
    }

    var existingAlertCount = parseInt(sheet.getRange(rowNumber, alertCountColumn).getValue(), 10);
    sheet.getRange(rowNumber, claimStatusColumn).setValue(
      String(sheet.getRange(rowNumber, claimStatusColumn).getValue() || '').trim() || 'Unclaimed'
    );
    sheet.getRange(rowNumber, alertStatusColumn).setValue(
      claimStatus.toLowerCase() === 'claimed' && claimedBy ? 'Claimed Reminder Sent' : 'Alerted'
    );
    sheet.getRange(rowNumber, alertLastSentColumn).setValue(now);
    sheet.getRange(rowNumber, alertCountColumn).setValue(isNaN(existingAlertCount) ? 1 : existingAlertCount + 1);
  }
}

// Spreadsheet form-submit trigger that sends the bishop FYI and requester
// confirmation emails, assigns request IDs, and records the submission status.
function onFormSubmitTrigger(e) {
  // 1. Get the data from the form submission
  var responses = e.namedValues;
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  ensureRequestIdForRow_(sheet, row);
  
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
