# Kindoo Stake Manager

Google Apps Script and Google Form tooling for the Kindoo building access workflow.

## Purpose

This project supports a scheduler-led process for LDS stake building access:

- building schedulers verify calendar availability and submit the request
- the ledger records the request and notification state
- bishops receive an FYI email
- stake managers receive alerts for requests nearing their access window
- managers claim requests and mark Kindoo key issuance through email-driven actions

## Scripts

- `scripts/kindoo_form_blueprint.gs`
  - repairs the live Google Form in place with the canonical structure
  - can also create a separate new fallback form without mutating the live one
- `scripts/notify_and_document.gs`
  - handles form-submit notifications
  - assigns request IDs
  - runs the upcoming-access scan
  - powers the claim and issued web-app actions

## Script Properties

The Apps Script project expects these Script Properties:

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

## Workflow

The workflow diagram lives in a separate source-of-truth file so it can be updated in one place:

- [Building Access Workflow](./building-access-workflow.md)

## Operational Notes

- Use `updateExistingKindooForm()` for normal repair of the live form.
- Use `createNewKindooForm()` only when you intentionally want a brand-new replacement form.
- `runUpcomingAccessScan()` should be configured as a daily time-driven trigger.
- `onFormSubmitTrigger(e)` should be configured as the form submit trigger for the ledger spreadsheet.
- The web app deployment must stay in sync with `WEB_APP_URL` after redeployments.
