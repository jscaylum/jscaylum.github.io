# Free newsletter setup

This uses Google Sheets and Google Apps Script. It does not require Mailchimp or a paid server.

## 1. Create the newsletter spreadsheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a blank spreadsheet.
2. Name it `jscaylum newsletter`.
3. Open **Extensions > Apps Script**.
4. Delete the starter code.
5. Open `newsletter/Code.gs` from this project and paste its contents into the Apps Script editor.
6. Click **Save**.
7. Select `setupNewsletter` in the function dropdown and click **Run**.
8. Approve Google’s permission prompts. This gives the script permission to edit this Sheet and send email from your account.

The script creates three tabs:

- `Subscribers`: email addresses collected from the website.
- `Draft`: the message you are preparing to send.
- `Campaigns`: a record of messages that were sent.

## 2. Publish the signup endpoint

1. In Apps Script, click **Deploy > New deployment**.
2. Choose **Web app**.
3. Set **Execute as** to `Me`.
4. Set **Who has access** to `Anyone`.
5. Click **Deploy** and copy the Web app URL.

The URL should end in `/exec`. Do not post it publicly with extra permissions or share your Google account credentials.

## 3. Connect the website

In `index.html`, find:

```js
const SIGNUP_PROVIDER = 'formsubmit';
```

Change it to:

```js
const SIGNUP_PROVIDER = 'googleSheets';
```

Then add this entry inside `SIGNUP_CONFIG`:

```js
googleSheets: {
  endpoint: 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE',
  payload: (email) => ({ email })
},
```

For this provider, the browser cannot reliably read the cross-origin Apps Script response. The signup handler should use a simple form-style POST and show success after the request is sent. The current FormSubmit path can remain active until the Apps Script URL is ready.

## 4. Send a message with a photo

1. Put a subject in `Draft!A2`.
2. Put the message in `Draft!B2`. Use a blank line between paragraphs.
3. Put a publicly accessible image URL in `Draft!C2`. A Google Drive sharing link is not always a direct image URL; use an image hosted by your site or another service that allows embedding.
4. In Apps Script, select `sendDraft` and click **Run**.
5. The message is sent to every subscriber whose `Status` is `active`.
6. The campaign is saved in the `Campaigns` tab.

## Free limits

Personal Gmail accounts have daily sending limits that Google can change. Start with small batches. Never put your Google password, verification code, or private API key in the website files.
