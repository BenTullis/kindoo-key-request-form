/**
 * Kindoo Building Access form management.
 *
 * Use updateExistingKindooForm() for normal repair work. It updates known
 * questions in place, recreates missing ones, and removes unexpected items
 * without recreating the canonical response-bound questions, which avoids
 * bloating the linked response sheet.
 *
 * Use hardResetKindooForm() only when you intentionally want a full wipe and
 * rebuild. That mode recreates all items and should be paired with a fresh
 * response destination.
 */

var LIVE_FORM_ID = '1ukbTtHgnHgpsce9_NAp196HlUBELamRkzRKHsBRxj9Y';

var KINDOO_FORM_SPEC = {
  title: 'Kindoo Building Access Request Form',
  description: 'Only Building Schedulers are authorized to fill out this form.',
  confirmationMessage: 'Thank you. An email has been sent to the requester and the Bishop of the selected ward.  The request has been logged on the Kindoo ledger, and the key will be assigned 7 days before the event. Thank you.',
  settings: {
    collectEmail: false,
    allowResponseEdits: true,
    limitOneResponsePerUser: false,
    progressBar: true,
    shuffleQuestions: false,
    showLinkToRespondAgain: true,
    publishingSummary: false
  },
  items: [
    {
      type: FormApp.ItemType.CHECKBOX,
      title: 'Scheduler Verification'
    },
    {
      type: FormApp.ItemType.PAGE_BREAK,
      title: 'Requester Details'
    },
    {
      type: FormApp.ItemType.TEXT,
      title: 'Requester Name'
    },
    {
      type: FormApp.ItemType.TEXT,
      title: 'Requester Email'
    },
    {
      type: FormApp.ItemType.TEXT,
      title: 'Requester Phone Number'
    },
    {
      type: FormApp.ItemType.MULTIPLE_CHOICE,
      title: 'Building Location'
    },
    {
      type: FormApp.ItemType.LIST,
      title: 'Requester\'s Ward'
    },
    {
      type: FormApp.ItemType.SECTION_HEADER,
      title: 'Important Schedule Notice'
    },
    {
      type: FormApp.ItemType.DATETIME,
      title: 'Access Start (Date & Time)'
    },
    {
      type: FormApp.ItemType.DATETIME,
      title: 'Access End (Date & Time)'
    }
  ]
};

// Safe repair mode for the live form. This updates the existing form in place
// so routine repairs do not create new response columns in the linked ledger.
function updateExistingKindooForm() {
  var form = FormApp.openById(LIVE_FORM_ID);

  applyKindooFormSettings_(form);
  removeUnexpectedItems_(form);
  ensureCanonicalItems_(form);
  reorderItemsToMatchSpec_(form);

  Logger.log('Form repaired successfully without recreating response columns: ' + form.getEditUrl());
}

// Full wipe-and-rebuild mode. This recreates all questions from scratch and can
// bloat the current linked ledger, so reconnect the form to a fresh response
// sheet before collecting any new responses after using this function.
function hardResetKindooForm() {
  var form = FormApp.openById(LIVE_FORM_ID);
  var items = form.getItems();

  items.forEach(function(item) {
    form.deleteItem(item);
  });

  applyKindooFormSettings_(form);
  createCanonicalItems_(form);

  Logger.log('Form hard reset completed. Reconnect to a fresh response sheet before collecting new responses: ' + form.getEditUrl());
}

// Applies the canonical form-level settings such as title, description,
// confirmation message, and response behavior.
function applyKindooFormSettings_(form) {
  form.setCollectEmail(KINDOO_FORM_SPEC.settings.collectEmail);
  form.setAllowResponseEdits(KINDOO_FORM_SPEC.settings.allowResponseEdits);
  form.setLimitOneResponsePerUser(KINDOO_FORM_SPEC.settings.limitOneResponsePerUser);
  form.setProgressBar(KINDOO_FORM_SPEC.settings.progressBar);
  form.setShuffleQuestions(KINDOO_FORM_SPEC.settings.shuffleQuestions);
  form.setShowLinkToRespondAgain(KINDOO_FORM_SPEC.settings.showLinkToRespondAgain);
  form.setPublishingSummary(KINDOO_FORM_SPEC.settings.publishingSummary);
  form.setConfirmationMessage(KINDOO_FORM_SPEC.confirmationMessage);
  form.setTitle(KINDOO_FORM_SPEC.title);
  form.setDescription(KINDOO_FORM_SPEC.description);
}

// Deletes items that are not part of the canonical form definition so the live
// form can recover from ad hoc manual additions.
function removeUnexpectedItems_(form) {
  var items = form.getItems();

  for (var i = items.length - 1; i >= 0; i--) {
    var item = items[i];
    if (!isExpectedItem_(item)) {
      form.deleteItem(item);
    }
  }
}

// Returns true when the given live form item matches a known canonical item by
// both type and title.
function isExpectedItem_(item) {
  for (var i = 0; i < KINDOO_FORM_SPEC.items.length; i++) {
    var spec = KINDOO_FORM_SPEC.items[i];
    if (item.getType() === spec.type && item.getTitle() === spec.title) {
      return true;
    }
  }
  return false;
}

// Ensures every canonical item exists and applies the expected configuration to
// it without deleting and recreating existing response-bound questions.
function ensureCanonicalItems_(form) {
  for (var i = 0; i < KINDOO_FORM_SPEC.items.length; i++) {
    var spec = KINDOO_FORM_SPEC.items[i];
    var item = findItemByTypeAndTitle_(form, spec.type, spec.title);

    if (!item) {
      item = createItemForSpec_(form, spec);
    }

    configureItemFromSpec_(item, spec);
  }
}

// Creates the full canonical item set from scratch. This is used only by the
// hard reset path.
function createCanonicalItems_(form) {
  for (var i = 0; i < KINDOO_FORM_SPEC.items.length; i++) {
    var spec = KINDOO_FORM_SPEC.items[i];
    var item = createItemForSpec_(form, spec);
    configureItemFromSpec_(item, spec);
  }
}

// Moves items into the expected order so the live form matches the blueprint
// after in-place repair work.
function reorderItemsToMatchSpec_(form) {
  for (var targetIndex = 0; targetIndex < KINDOO_FORM_SPEC.items.length; targetIndex++) {
    var spec = KINDOO_FORM_SPEC.items[targetIndex];
    var item = findItemByTypeAndTitle_(form, spec.type, spec.title);

    if (!item) {
      continue;
    }

    var currentIndex = getItemIndexById_(form, item.getId());
    if (currentIndex !== -1 && currentIndex !== targetIndex) {
      form.moveItem(currentIndex, targetIndex);
    }
  }
}

// Finds the current index of a form item by its stable Google Forms item ID.
function getItemIndexById_(form, itemId) {
  var items = form.getItems();
  for (var i = 0; i < items.length; i++) {
    if (items[i].getId() === itemId) {
      return i;
    }
  }
  return -1;
}

// Finds a live form item using the canonical type and title pair.
function findItemByTypeAndTitle_(form, itemType, title) {
  var items = form.getItems(itemType);
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === title) {
      return items[i];
    }
  }
  return null;
}

// Creates one new form item for the given canonical spec entry.
function createItemForSpec_(form, spec) {
  switch (spec.type) {
    case FormApp.ItemType.CHECKBOX:
      return form.addCheckboxItem().setTitle(spec.title);
    case FormApp.ItemType.PAGE_BREAK:
      return form.addPageBreakItem().setTitle(spec.title);
    case FormApp.ItemType.TEXT:
      return form.addTextItem().setTitle(spec.title);
    case FormApp.ItemType.MULTIPLE_CHOICE:
      return form.addMultipleChoiceItem().setTitle(spec.title);
    case FormApp.ItemType.LIST:
      return form.addListItem().setTitle(spec.title);
    case FormApp.ItemType.SECTION_HEADER:
      return form.addSectionHeaderItem().setTitle(spec.title);
    case FormApp.ItemType.DATETIME:
      return form.addDateTimeItem().setTitle(spec.title);
    default:
      throw new Error('Unsupported item type in form spec: ' + spec.type);
  }
}

// Applies canonical configuration to an existing or newly created item based on
// its spec definition.
function configureItemFromSpec_(item, spec) {
  switch (spec.type) {
    case FormApp.ItemType.CHECKBOX:
      configureSchedulerVerification_(item.asCheckboxItem());
      return;
    case FormApp.ItemType.PAGE_BREAK:
      item.asPageBreakItem().setTitle(spec.title);
      return;
    case FormApp.ItemType.TEXT:
      configureTextItem_(item.asTextItem(), spec.title);
      return;
    case FormApp.ItemType.MULTIPLE_CHOICE:
      item.asMultipleChoiceItem()
        .setTitle(spec.title)
        .setChoiceValues(['Stake Center', 'South Building'])
        .setRequired(true);
      return;
    case FormApp.ItemType.LIST:
      item.asListItem()
        .setTitle(spec.title)
        .setChoiceValues(['1st Ward', '2nd Ward', '4th Ward', '5th Ward', '7th Ward'])
        .setRequired(true);
      return;
    case FormApp.ItemType.SECTION_HEADER:
      item.asSectionHeaderItem()
        .setTitle(spec.title)
        .setHelpText('Note: There is no form validation to ensure that a date is not set in the past, or that the end date is after the start date. Please take care when filling out.');
      return;
    case FormApp.ItemType.DATETIME:
      item.asDateTimeItem()
        .setTitle(spec.title)
        .setRequired(true);
      return;
    default:
      throw new Error('Unsupported item type in form config: ' + spec.type);
  }
}

// Configures the required scheduler verification checkbox section, including
// the exact five required acknowledgements.
function configureSchedulerVerification_(checkboxItem) {
  var validation = FormApp.createCheckboxValidation()
    .requireSelectExactly(5)
    .setHelpText('Select all boxes')
    .build();

  checkboxItem
    .setTitle('Scheduler Verification')
    .setRequired(true)
    .setChoices([
      checkboxItem.createChoice('I have confirmed that the building is available.'),
      checkboxItem.createChoice('I have confirmed the requester knows how Kindoo works.'),
      checkboxItem.createChoice('I have confirmed the requester has or will download the Kindoo app.'),
      checkboxItem.createChoice('I have confirmed the requester knows that Kindoo and the Member Tools app must have the same email.'),
      checkboxItem.createChoice('I have scheduled the event.')
    ])
    .setValidation(validation);
}

// Configures each canonical text field with the correct validation and help
// text based on its title.
function configureTextItem_(textItem, title) {
  if (title === 'Requester Name') {
    var requesterNameValidation = FormApp.createTextValidation()
      .requireTextMatchesPattern('^([^0-9]*)$')
      .setHelpText('Numbers are not allowed in the name field.')
      .build();

    textItem
      .setTitle(title)
      .setValidation(requesterNameValidation)
      .setRequired(true);
    return;
  }

  if (title === 'Requester Email') {
    var requesterEmailValidation = FormApp.createTextValidation()
      .requireTextIsEmail()
      .setHelpText('Please enter a valid email address.')
      .build();

    textItem
      .setTitle(title)
      .setHelpText('Must match Member Tools and the Kindoo app')
      .setValidation(requesterEmailValidation)
      .setRequired(true);
    return;
  }

  if (title === 'Requester Phone Number') {
    var requesterPhoneValidation = FormApp.createTextValidation()
      .requireTextMatchesPattern('^[0-9]{3}-[0-9]{3}-[0-9]{4}$')
      .setHelpText('Please enter the phone number exactly as 123-123-1234.')
      .build();

    textItem
      .setTitle(title)
      .setHelpText('Does not need to match Member Tools nor the Kindoo app')
      .setValidation(requesterPhoneValidation)
      .setRequired(true);
      return;
  }

  throw new Error('Unsupported text item title in form spec: ' + title);
}
