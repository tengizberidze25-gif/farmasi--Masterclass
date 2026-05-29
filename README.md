# FARMASI Masterclass 2026

Files in this workspace:

- `index.html` - landing page for `https://masterclass.farmasi.ge`.
- `checkin.html` - QR/ticket check-in page for event staff.
- `Farmasi Masterclass 2026.xlsx` - source workbook structure for Google Sheets.
- `Code.gs` - Google Apps Script Web App backend for this masterclass site.

## Apps Script Setup

1. Upload/import `Farmasi Masterclass 2026.xlsx` into Google Sheets.
2. Open `Extensions -> Apps Script`.
3. Replace the script content with `Code.gs`.
4. Deploy as Web App:
   - Execute as: `Me`
   - Who has access: `Anyone`
5. Copy the Web App URL into `index.html` and `checkin.html` at:

```js
const API_URL = "YOUR_WEB_APP_URL";
```

The current files already point to the deployed Apps Script URL used by the site.

The frontend uses JSONP for Apps Script requests, so it works from the production custom domain without browser CORS failures.

## Script Properties

For SMS delivery, set these Apps Script project properties:

- `PUBLIC_KEY`
- `PRIVATE_KEY`
- `SMS_SENDER` optional, defaults to `FARMASI`

Without these properties, the site still loads, registers, cancels, and shows tickets, but `sendTicketLink` cannot send SMS.
