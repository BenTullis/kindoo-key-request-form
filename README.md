# Kindoo Stake Manager

Google Apps Script and Google Form tooling for the Kindoo building access workflow.

## Purpose

This project supports a scheduler-led process for LDS stake building access:

- building schedulers verify calendar availability and submit the request
- the ledger records the request and notification state
- bishops receive an FYI email
- stake kindoo managers receive alerts for requests nearing their access window
- one manager claims the request through an email link
- the claiming manager receives daily reminders until the key is issued
- the member receives the final success email after issuance

## Scripts

- `kindoo-stake-manager/scripts/kindoo_form_blueprint.gs`
  - repairs the live Google Form in place with the canonical structure
  - can also create a separate new fallback form without mutating the live one
- `kindoo-stake-manager/scripts/notify_and_document.gs`
  - handles form-submit notifications
  - assigns request IDs
  - runs the upcoming-access scan
  - powers the claim and issued web-app actions
  - creates installable triggers

## Script Properties

The Apps Script project expects these required Script Properties:

- `WARD_1_EMAIL`
- `WARD_2_EMAIL`
- `WARD_4_EMAIL`
- `WARD_5_EMAIL`
- `WARD_7_EMAIL`
- `STAKE_TECHNOLOGY_SPECIALIST_EMAIL`
- `STAKE_MANAGER_EMAILS`
- `CLAIM_LINK_SECRET`
- `ISSUED_LINK_SECRET`

These Script Properties are optional but recommended:

- `LEDGER_SPREADSHEET_ID`
  - `1LvGWUqpqwAzyTkMfLphH_5mO2hzPgjzwt64C5X-bknE`
  - if omitted, the script uses the active spreadsheet when running from a spreadsheet-bound project
- `WEB_APP_URL`
  - if omitted, the script falls back to the deployed Apps Script service URL when available
- `LEDGER_SHEET_NAME`
  - if omitted, the script uses the first sheet in the ledger spreadsheet

`STAKE_TECHNOLOGY_SPECIALIST_EMAIL` is also used as the fallback recipient when the submitted ward does not match a configured ward email.

## Form Blueprint

The live form is repaired from the canonical blueprint in `kindoo-stake-manager/scripts/kindoo_form_blueprint.gs`.

Canonical form settings:

- collects email: `false`
- allows response edits: `true`
- limits one response per user: `false`
- shows progress bar: `true`
- shuffles questions: `false`
- shows link to respond again: `true`
- publishes response summary: `false`

Canonical form fields:

- `Scheduler Verification`
  - required checkbox field
  - requires exactly five acknowledgements
- `Requester Name`
  - required text field
  - rejects numbers
- `Requester Email`
  - required email field
  - must match Member Tools and the Kindoo app
- `Requester Phone Number`
  - required text field
  - must match `123-123-1234`
- `Building Location`
  - required multiple choice field
  - choices: `Stake Center`, `South Building`
- `Requester's Ward`
  - required dropdown field
  - choices: `1st Ward`, `2nd Ward`, `4th Ward`, `5th Ward`, `7th Ward`
- `Access Start (Date & Time)`
  - required date/time field
- `Access End (Date & Time)`
  - required date/time field

The form includes a schedule notice that the script does not validate whether dates are in the past or whether the end date is after the start date.

## Setup Order

For a fresh Apps Script owner or a moved Google account, use this order:

1. Paste `kindoo-stake-manager/scripts/notify_and_document.gs` and `kindoo-stake-manager/scripts/kindoo_form_blueprint.gs` into the live Apps Script project.
2. Create or update the required Script Properties manually in the live Apps Script project.
3. Run `createKindooTriggers()` once.
4. Deploy the script as a web app:
   - execute as `Me`
   - allow access for `Anyone with Google account`
5. If the deployment URL changes, update `WEB_APP_URL` in Script Properties.

## Triggered Workflow

- `onFormSubmitTrigger(e)`
  - assigns a stable `Request ID`
  - emails the bishop
  - emails the requester
  - records `Vetted and Scheduled` or `Updated and Scheduled`
- `runUpcomingAccessScan()`
  - alerts all stake kindoo managers for unclaimed requests within the next 7 days
  - once claimed, reminds only the claiming manager daily until issuance
  - skips requests already issued, requests in the past, requests more than 7 days away, and rows already alerted that day
  - stops reminders after the key is issued
- `doGet(e)`
  - handles signed `action=claim` links
  - handles signed `action=issued` links
  - updates the ledger and sends follow-up notifications

## Ledger Columns

The script creates or updates helper columns in the ledger as needed:

- `Request ID`
  - generated as `REQ-0001`, `REQ-0002`, and so on
- `Status`
  - set to `Vetted and Scheduled` for new submissions
  - set to `Updated and Scheduled` when an edited submission already has a status
- `Manager Claim Status`
- `Manager Alert Status`
- `Manager Alert Last Sent At`
- `Manager Alert Count`
- `Claimed By`
- `Claimed At`
- `Kindoo Key Status`
- `Issued By`
- `Issued At`

The claim and issued links are signed with HMAC tokens using the request ID, action, actor email, and the matching secret property. A request can only be claimed once, and an issued request is not processed twice.

## Workflow

The workflow diagram lives in a separate source-of-truth file so it can be updated in one place:

- [Building Access Workflow](./kindoo-stake-manager/building-access-workflow.md)

## Maintenance Helpers

- `updateExistingKindooForm()`
  - repairs the live form in place without recreating response-bound questions
- `createNewKindooForm()`
  - creates a separate new form for true replacement scenarios
- `cleanupDuplicateRequestIdColumns()`
  - merges duplicate `Request ID` columns into the leftmost canonical column
- `createKindooTriggers()`
  - recreates the spreadsheet submit and daily scan triggers

## Operational Notes

- Use `updateExistingKindooForm()` for normal repair of the live form.
- Use `createNewKindooForm()` only when you intentionally want a brand-new replacement form.
- `runUpcomingAccessScan()` is configured by `createKindooDailyScanTrigger()` as a daily time-driven trigger at 8 AM in the script timezone.
- `onFormSubmitTrigger(e)` should be configured as the form submit trigger for the ledger spreadsheet.
- The web app deployment must stay in sync with `WEB_APP_URL` after redeployments.
- Old claim or issued emails should be treated as stale after a redeployment or secret change.
