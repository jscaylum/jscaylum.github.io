const SHEET_NAMES = {
  subscribers: 'Subscribers',
  campaigns: 'Campaigns',
  draft: 'Draft',
  scheduled: 'Scheduled'
};

function setupNewsletter() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheet.getId());

  ensureSheet_(spreadsheet, SHEET_NAMES.subscribers, ['Email', 'Signed up', 'Status', 'Welcome']);
  ensureSheet_(spreadsheet, SHEET_NAMES.campaigns, ['Sent', 'Subject', 'Message', 'Attachments', 'Recipients']);
  ensureSheet_(spreadsheet, SHEET_NAMES.scheduled, ['Subject', 'Message', 'Attachments', 'Send at', 'Status']);
  const draft = ensureSheet_(spreadsheet, SHEET_NAMES.draft, ['Subject', 'Message', 'Drive file ID(s)', 'Send at (optional)']);

  if (draft.getLastRow() < 2) {
    draft.getRange(2, 1, 1, 4).setValues([[
      'a little note from jscaylum',
      'write your message here',
      '',
      ''
    ]]);
  }
}

function doGet(event) {
  const unsubscribeEmail = event && event.parameter && event.parameter.unsubscribe;
  if (unsubscribeEmail) return handleUnsubscribe_(unsubscribeEmail);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'jscaylum newsletter' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleUnsubscribe_(email) {
  const normalized = String(email).trim().toLowerCase();
  const spreadsheet = getSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.subscribers);
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const emails = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    emails.forEach((row, index) => {
      if (String(row[0]).trim().toLowerCase() === normalized) {
        sheet.getRange(index + 2, 3).setValue('unsubscribed');
      }
    });
  }

  return HtmlService.createHtmlOutput(
    '<p style="font-family:Georgia,serif;padding:40px;">you\'ve been unsubscribed. take care.</p>'
  );
}

function doPost(event) {
  try {
    const data = parseRequest_(event);
    if (!data.email) return json_({ ok: false, error: 'email is required' });

    const email = String(data.email).trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      return json_({ ok: false, error: 'invalid email' });
    }

    const spreadsheet = getSpreadsheet_();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.subscribers);
    const emails = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1)
      .getValues()
      .flat()
      .map(value => String(value).trim().toLowerCase());

    const isNewSubscriber = !emails.includes(email);
    if (isNewSubscriber) {
      sheet.appendRow([email, new Date(), 'active', 'sending']);
      const rowNumber = sheet.getLastRow();
      try {
        sendWelcomeEmail_(email);
        sheet.getRange(rowNumber, 4).setValue('sent');
      } catch (welcomeError) {
        sheet.getRange(rowNumber, 4).setValue(`failed: ${welcomeError.message}`);
        console.warn(`Welcome email failed for ${email}: ${welcomeError.message}`);
      }
    }

    return json_({ ok: true });
  } catch (error) {
    return json_({ ok: false, error: error.message });
  }
}

function sendWelcomeEmail_(email) {
  const subject = 'welcome to the quiet side of things';
  const message = [
    'hey,',
    '',
    'you\'re in.',
    '',
    'this is where i\'ll send the things that aren\'t public yet: song previews, private photos, early news, and little notes i don\'t want to post anywhere else.',
    '',
    'no noise. no constant updates. just the first look when there\'s something worth sharing.',
    '',
    'glad you\'re here.',
    '',
    'jscaylum'
  ].join('\\n');

  GmailApp.sendEmail(email, subject, message, {
    htmlBody: buildWelcomeHtml_(),
    name: 'message from jscaylum'
  });
}

function buildWelcomeHtml_() {
  return `<!doctype html><html><body style="margin:0;background:#f4eee9;color:#211c22;font:16px/1.7 Georgia,serif;"><div style="max-width:620px;margin:0 auto;padding:48px 28px;"><p style="margin:0 0 28px;color:#a3576b;font:11px/1.2 Arial,sans-serif;letter-spacing:.2em;text-transform:uppercase;">jscaylum · welcome</p><h1 style="margin:0 0 28px;font:italic 42px/1 Georgia,serif;color:#211c22;">you're in.</h1><p>this is where i\'ll send the things that aren\'t public yet: song previews, private photos, early news, and little notes i don\'t want to post anywhere else.</p><p>no noise. no constant updates. just the first look when there\'s something worth sharing.</p><p style="margin-top:36px;">glad you\'re here.<br><br>jscaylum</p><div style="height:1px;margin-top:42px;background:#d8c9c4;"></div><p style="margin-top:16px;color:#8b7881;font:12px/1.5 Arial,sans-serif;">private previews, early news, and things before they go public.</p></div></body></html>`;
}

function sendTestWelcomeEmail() {
  const email = Session.getEffectiveUser().getEmail();
  if (!email) throw new Error('Google could not identify the account running this script.');
  sendWelcomeEmail_(email);
}

function retryFailedWelcomeEmails() {
  const spreadsheet = getSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.subscribers);
  const rows = sheet.getLastRow() < 2
    ? []
    : sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();

  rows.forEach((row, index) => {
    const email = String(row[0]).trim();
    const status = String(row[2]).trim().toLowerCase();
    const welcome = String(row[3]).trim().toLowerCase();
    if (!email || status !== 'active' || !welcome.startsWith('failed:')) return;

    const rowNumber = index + 2;
    try {
      sendWelcomeEmail_(email);
      sheet.getRange(rowNumber, 4).setValue('sent');
    } catch (error) {
      sheet.getRange(rowNumber, 4).setValue(`failed: ${error.message}`);
    }
  });
}

// permanently deletes rows for anyone who unsubscribed (unsubscribing alone just hides them from sends)
function deleteUnsubscribedRows() {
  const spreadsheet = getSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.subscribers);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const statuses = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  for (let i = statuses.length - 1; i >= 0; i--) {
    if (String(statuses[i][0]).trim().toLowerCase() === 'unsubscribed') {
      sheet.deleteRow(i + 2);
    }
  }
}

function sendDraft() {
  const spreadsheet = getSpreadsheet_();
  const draft = spreadsheet.getSheetByName(SHEET_NAMES.draft);
  const values = draft.getRange(2, 1, 1, 3).getValues()[0];
  const subject = String(values[0] || '').trim();
  const message = String(values[1] || '').trim();
  const attachmentIds = parseAttachmentIds_(values[2]);

  if (!subject || !message) throw new Error('Add a subject and message in the Draft sheet first.');
  sendCampaign_(subject, message, attachmentIds);
}

function scheduleDraft() {
  const spreadsheet = getSpreadsheet_();
  const draft = spreadsheet.getSheetByName(SHEET_NAMES.draft);
  const values = draft.getRange(2, 1, 1, 4).getValues()[0];
  const subject = String(values[0] || '').trim();
  const message = String(values[1] || '').trim();
  const attachmentIds = parseAttachmentIds_(values[2]);
  const sendAt = values[3];

  if (!subject || !message) throw new Error('Add a subject and message in the Draft sheet first.');
  if (!(sendAt instanceof Date) || isNaN(sendAt.getTime())) {
    throw new Error('Put a future date and time in the "Send at (optional)" cell, or use sendDraft to send immediately.');
  }
  if (sendAt.getTime() <= Date.now()) {
    throw new Error('The scheduled time must be in the future.');
  }

  spreadsheet.getSheetByName(SHEET_NAMES.scheduled)
    .appendRow([subject, message, attachmentIds.join(', '), sendAt, 'pending']);

  enableScheduledSending();
}

// run automatically by a time trigger; sends any due, still-pending campaigns
function runScheduledCampaigns() {
  const spreadsheet = getSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.scheduled);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const now = Date.now();

  rows.forEach((row, index) => {
    const [subject, message, attachmentsText, sendAt, status] = row;
    if (String(status).trim().toLowerCase() !== 'pending') return;
    if (!(sendAt instanceof Date) || sendAt.getTime() > now) return;

    const rowNumber = index + 2;
    const attachmentIds = String(attachmentsText || '').split(',').map(id => id.trim()).filter(Boolean);
    try {
      sendCampaign_(String(subject), String(message), attachmentIds);
      sheet.getRange(rowNumber, 5).setValue('sent');
    } catch (error) {
      sheet.getRange(rowNumber, 5).setValue(`failed: ${error.message}`);
    }
  });
}

function enableScheduledSending() {
  const alreadyRunning = ScriptApp.getProjectTriggers()
    .some(trigger => trigger.getHandlerFunction() === 'runScheduledCampaigns');
  if (alreadyRunning) return;

  ScriptApp.newTrigger('runScheduledCampaigns')
    .timeBased()
    .everyMinutes(15)
    .create();
}

function sendCampaign_(subject, message, attachmentIds) {
  const spreadsheet = getSpreadsheet_();
  const subscribers = spreadsheet.getSheetByName(SHEET_NAMES.subscribers);
  const rows = subscribers.getLastRow() < 2
    ? []
    : subscribers.getRange(2, 1, subscribers.getLastRow() - 1, 3).getValues();
  const recipients = rows
    .filter(row => String(row[2]).toLowerCase() === 'active')
    .map(row => String(row[0]).trim())
    .filter(Boolean);

  if (!recipients.length) throw new Error('There are no active subscribers yet.');

  const attachments = attachmentIds.map(id => DriveApp.getFileById(id).getBlob());

  // sent one at a time (not comma-joined) so subscribers never see each other's addresses
  recipients.forEach(email => {
    GmailApp.sendEmail(email, subject, message, {
      htmlBody: buildEmailHtml_(message, email),
      attachments,
      name: 'message from jscaylum'
    });
  });

  spreadsheet.getSheetByName(SHEET_NAMES.campaigns).appendRow([
    new Date(), subject, message, attachmentIds.join(', '), recipients.length
  ]);
}

function buildEmailHtml_(message, email) {
  const paragraphs = escapeHtml_(message)
    .split(/\n{2,}/)
    .map(paragraph => `<p>${paragraph.replace(/\n/g, '<br>')}</p>`)
    .join('');
  const unsubscribeUrl = `${ScriptApp.getService().getUrl()}?unsubscribe=${encodeURIComponent(email)}`;
  return `<!doctype html><html><body style="margin:0;background:#f4eee9;color:#211c22;font:16px/1.6 Georgia,serif;"><div style="max-width:620px;margin:0 auto;padding:36px 24px;">${paragraphs}<p style="margin-top:36px;color:#8b7881;font-size:13px;">jscaylum · quiet updates, songs, and photos</p><p style="margin-top:8px;"><a href="${unsubscribeUrl}" style="color:#8b7881;font-size:12px;">unsubscribe</a></p></div></body></html>`;
}

function parseAttachmentIds_(value) {
  return String(value || '')
    .split(/[\s,]+/)
    .map(value => value.trim())
    .filter(Boolean)
    .map(value => {
      const match = value.match(/[-\w]{20,}/);
      return match ? match[0] : value;
    });
}

function parseRequest_(event) {
  if (event && event.postData && event.postData.contents) {
    try {
      return JSON.parse(event.postData.contents);
    } catch (error) {
      return event.parameter || {};
    }
  }
  return (event && event.parameter) || {};
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Run setupNewsletter once from the Apps Script editor.');
  return SpreadsheetApp.openById(id);
}

function ensureSheet_(spreadsheet, name, headers) {
  const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function escapeHtml_(value) {
  return value.replace(/[&<>"']/g, character => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[character]));
}

function escapeAttribute_(value) {
  return escapeHtml_(value);
}
