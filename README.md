# Dripify Reply Tracker

Google Sheets + Apps Script webhook tracker for Dripify first replies.

## What It Does

When someone replies to a Dripify campaign, Dripify sends a webhook to a Google Apps Script Web App. The script adds the lead to a central Google Sheet so replies can be reviewed, filtered, and organized without Slack or Zapier.

The tracker captures:

- Received timestamp
- Team member
- Campaign
- Lead name
- Company
- Title
- LinkedIn URL
- Email
- Status
- Notes
- Raw Dripify payload

## Sheet Columns

Create a Google Sheet with these headers in row 1:

```text
Received At | Team Member | Campaign | Lead Name | Company | Title | Linkedin URL | Email | Status | Notes | Raw Payload
```

The script expects the sheet tab to be named `Sheet1`. If your tab has a different name, update this line in `Code.gs`:

```javascript
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
```

## Install

1. Open the Google Sheet.
2. Go to **Extensions** -> **Apps Script**.
3. Delete the starter `myFunction` code.
4. Paste the contents of `Code.gs`.
5. Save the project.

## Deploy

1. In Apps Script, click **Deploy**.
2. Click **New deployment**.
3. Choose **Web app**.
4. Set **Execute as** to `Me`.
5. Set **Who has access** to `Anyone`.
6. Click **Deploy**.
7. Authorize the script.
8. Copy the Web App URL.

## Dripify Webhook URL Format

Use the Web App URL in Dripify, then add team member and campaign labels:

```text
WEB_APP_URL?teamMember=TEAM_MEMBER&campaign=CAMPAIGN
```

Example:

```text
https://script.google.com/macros/s/DEPLOYMENT_ID/exec?teamMember=Sneha&campaign=Offer-Insurance
```

For names or campaigns with spaces, use `%20`:

```text
https://script.google.com/macros/s/DEPLOYMENT_ID/exec?teamMember=Alex%20Cardoso&campaign=Offer%20led%20campaign%20messaging
```

## Dripify Setup

For each campaign:

1. Open the campaign in Dripify.
2. Go to the webhook/integration settings.
3. Add a webhook for the **Response** or first-reply event.
4. Paste the campaign-specific URL.
5. Save and test.

## Notes

Dripify may not include the campaign name, team member name, or reply text in its webhook payload. This tracker solves team member and campaign tracking by putting those values directly in the webhook URL.

If Dripify sends `conversation: null`, the reply text is not available from the webhook. In that case, use this sheet as the first-reply alert log, then open Dripify or LinkedIn to read the actual message.
