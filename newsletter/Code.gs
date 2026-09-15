const SHEET_NAMES = {
  subscribers: 'Subscribers',
  campaigns: 'Campaigns',
  draft: 'Draft'
};

function setupNewsletter() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheet.getId());

  ensureSheet_(spreadsheet, SHEET_NAMES.subscribers, ['Email', 'Signed up', 'Status']);
  ensureSheet_(spreadsheet, SHEET_NAMES.campaigns, ['Sent', 'Subject', 'Message', 'Photo URL', 'Recipients']);
  const draft = ensureSheet_(spreadsheet, SHEET_NAMES.draft, ['Subject', 'Message', 'Photo URL']);

  if (draft.getLastRow() < 2) {
    draft.getRange(2, 1, 1, 3).setValues([[
      'a little note from jscaylum',
      'write your message here',
      ''
    ]]);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'jscaylum newsletter' }))
    .setMimeType(ContentService.MimeType.JSON);
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

    if (!emails.includes(email)) {
      sheet.appendRow([email, new Date(), 'active']);
    }

    return json_({ ok: true });
  } catch (error) {
    return json_({ ok: false, error: error.message });
  }
}

function sendDraft() {
  const spreadsheet = getSpreadsheet_();
  const draft = spreadsheet.getSheetByName(SHEET_NAMES.draft);
  const values = draft.getRange(2, 1, 1, 3).getValues()[0];
  const subject = String(values[0] || '').trim();
  const message = String(values[1] || '').trim();
  const photoUrl = String(values[2] || '').trim();

  if (!subject || !message) throw new Error('Add a subject and message in the Draft sheet first.');
  sendCampaign_(subject, message, photoUrl);
}

function sendCampaign_(subject, message, photoUrl) {
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

  const htmlBody = buildEmailHtml_(message, photoUrl);
  const plainBody = photoUrl ? `${message}\n\nPhoto: ${photoUrl}` : message;
  const batchSize = 80;

  for (let index = 0; index < recipients.length; index += batchSize) {
    const batch = recipients.slice(index, index + batchSize);
    GmailApp.sendEmail(batch.join(','), subject, plainBody, {
      htmlBody,
      name: 'jscaylum'
    });
  }

  spreadsheet.getSheetByName(SHEET_NAMES.campaigns).appendRow([
    new Date(), subject, message, photoUrl, recipients.length
  ]);
}

function buildEmailHtml_(message, photoUrl) {
  const paragraphs = escapeHtml_(message)
    .split(/\n{2,}/)
    .map(paragraph => `<p>${paragraph.replace(/\n/g, '<br>')}</p>`)
    .join('');
  const image = photoUrl
    ? `<p><img src="${escapeAttribute_(photoUrl)}" alt="" style="display:block;max-width:100%;height:auto;border-radius:8px;"></p>`
    : '';

  return `<!doctype html><html><body style="margin:0;background:#f4eee9;color:#211c22;font:16px/1.6 Georgia,serif;"><div style="max-width:620px;margin:0 auto;padding:36px 24px;">${image}${paragraphs}<p style="margin-top:36px;color:#8b7881;font-size:13px;">jscaylum · quiet updates, songs, and photos</p></div></body></html>`;
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
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
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
