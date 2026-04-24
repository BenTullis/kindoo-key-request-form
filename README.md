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

The Apps Script project expects these Script Properties:

- `LEDGER_SPREADSHEET_ID`
  - `1LvGWUqpqwAzyTkMfLphH_5mO2hzPgjzwt64C5X-bknE`
- `WARD_1_EMAIL`
- `WARD_2_EMAIL`
- `WARD_4_EMAIL`
- `WARD_5_EMAIL`
- `WARD_7_EMAIL`
- `STAKE_TECHNOLOGY_SPECIALIST_EMAIL`
- `STAKE_MANAGER_EMAILS`
- `CLAIM_LINK_SECRET`
- `ISSUED_LINK_SECRET`
- `WEB_APP_URL`
- `LEDGER_SHEET_NAME`

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
  - stops reminders after the key is issued
- `doGet(e)`
  - handles the claim link
  - handles the issued link
  - updates the ledger and sends follow-up notifications

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
- `runUpcomingAccessScan()` should be configured as a daily time-driven trigger.
- `onFormSubmitTrigger(e)` should be configured as the form submit trigger for the ledger spreadsheet.
- The web app deployment must stay in sync with `WEB_APP_URL` after redeployments.
- Old claim or issued emails should be treated as stale after a redeployment or secret change.
