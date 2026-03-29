/**
 * Kindoo Building Access: Create/Update Existing Form
 * This script updates your LIVE form by ID to match the implementation plan.
 */

// 1. SET YOUR LIVE FORM ID HERE
var LIVE_FORM_ID = '1ukbTtHgnHgpsce9_NAp196HlUBELamRkzRKHsBRxj9Y'; 

function updateExistingKindooForm() {
  var form = FormApp.openById(LIVE_FORM_ID);
  
  // 1. CLEAR THE SLATE
  var items = form.getItems();
  items.forEach(function(item) {
    form.deleteItem(item);
  });
  
/**
 * Paste START
 */
  form.setCollectEmail(false); // Disables the "Verified Login" requirement
  form.setAllowResponseEdits(true); // Allows Schedulers to fix typos
  form.setLimitOneResponsePerUser(false); // Ensures they can submit multiple requests

  // --- AUTOMATED PRESENTATION SETTINGS ---
  form.setProgressBar(true);                  // Show progress bar
  form.setShuffleQuestions(false);            // Shuffle question order: OFF
  form.setShowLinkToRespondAgain(true);       // Show link to submit another: ON
  form.setPublishingSummary(false);           // View results summary: OFF
  
  // Sets your custom confirmation message automatically
  form.setConfirmationMessage('Thank you. An email has been sent to the requester and the Bishop of the selected ward.  The request has been logged on the Kindoo ledger, and the key will be assigned 7 days before the event. Thank you.');


  form.setTitle('Kindoo Building Access Request Form');
  form.setDescription('Only Building Schedulers are authorized to fill out this form.');

  var cb = form.addCheckboxItem().setTitle('Scheduler Verification').setRequired(true);
  var cbVal = FormApp.createCheckboxValidation().requireSelectExactly(5).setHelpText('Select all boxes').build();
  cb.setChoices([cb.createChoice('I have confirmed that the building is available.'), cb.createChoice('I have confirmed the requester knows how Kindoo works.'), cb.createChoice('I have confirmed the requester has or will download the Kindoo app.'), cb.createChoice('I have confirmed the requester knows that Kindoo and the Member Tools app must have the same email.'), cb.createChoice('I have scheduled the event.')]).setValidation(cbVal);

  form.addPageBreakItem().setTitle('Requester Details');

// 1. Create the validation rule (No numbers allowed)
  var val_RequesterName = FormApp.createTextValidation()
    .requireTextMatchesPattern('^([^0-9]*)$')
    .setHelpText('Numbers are not allowed in the name field.')
    .build();

  // 2. Apply it to the Text Item
  form.addTextItem()
      .setTitle('Requester Name')
      .setValidation(val_RequesterName)
      .setRequired(true);

// 1. Create the email validation rule
  var val_RequesterEmail = FormApp.createTextValidation()
    .requireTextIsEmail()
    .setHelpText('Please enter a valid email address.')
    .build();

  // 2. Apply it to the Text Item
  form.addTextItem()
      .setTitle('Requester Email')
      .setHelpText('Must match Member Tools and the Kindoo app')
      .setValidation(val_RequesterEmail)
      .setRequired(true);

// 1. Create the phone validation rule (Forces digits and hyphens only)
  var val_RequesterPhoneNumber = FormApp.createTextValidation()
    .requireTextMatchesPattern('^[0-9]{3}-[0-9]{3}-[0-9]{4}$')
    .setHelpText('Please enter the phone number exactly as 123-123-1234.')
    .build();

  // 2. Apply it to the Text Item
  form.addTextItem()
      .setTitle('Requester Phone Number')
      .setHelpText('Does not need to match Member Tools nor the Kindoo app')
      .setValidation(val_RequesterPhoneNumber)
      .setRequired(true);

// 1. Building Location (Multiple Choice)
  form.addMultipleChoiceItem()
      .setTitle('Building Location')
      .setChoiceValues(['Stake Center', 'South Building'])
      .setRequired(true);

  // 2. Ward Selection (Dropdown)
  form.addListItem()
      .setTitle('Requester\'s Ward')
      .setChoiceValues(['1st Ward', '2nd Ward', '4th Ward', '5th Ward', '7th Ward'])
      .setRequired(true);

  // 3. Schedule Disclaimer & Dates
  form.addSectionHeaderItem()
      .setTitle('Important Schedule Notice')
      .setHelpText('Note: There is no form validation to ensure that a date is not set in the past, or that the end date is after the start date. Please take care when filling out.');

  form.addDateTimeItem().setTitle('Access Start (Date & Time)').setRequired(true);
  form.addDateTimeItem().setTitle('Access End (Date & Time)').setRequired(true);

/**
 * Paste END
 */
 

  Logger.log('Form updated successfully: ' + form.getEditUrl());
}
