/**
 * Lead Capture Automation - "Request a Quote"
 * Google Form -> Google Sheet -> Thank-you email + owner alert + Status tracking
 *
 * Setup:
 *  1. Open the response Sheet linked to your Form > Extensions > Apps Script.
 *  2. Paste this file into Code.gs and edit CONFIG below.
 *  3. Choose `setup` in the toolbar and click Run (approve the permissions once).
 *  4. Submit the form to test.
 */

// ====== CONFIG: edit these ======
const CONFIG = {
  OWNER_EMAIL: 'you@example.com',         // where new-lead alerts go
  BUSINESS_NAME: 'Your Business Name',
  SHEET_NAME: 'Form Responses 1',         // tab the form writes to
  // These must match your form question titles EXACTLY
  FIELDS: {
    name: 'Full name',
    email: 'Email address',
    phone: 'Phone number',
    company: 'Company / business name',
    service: 'What do you need help with?',
    budget: 'Estimated budget',
    timeline: 'When do you need this done?',
    details: 'Project details',
  },
  STATUS_HEADER: 'Status',
  STATUS_OPTIONS: ['New', 'Contacted', 'Quote sent', 'Won', 'Lost'],
  REPLY_WITHIN: '24 hours',
};

/** Run once: adds the Status column, dropdown, colours, and the form-submit trigger. */
function setup() {
  const sheet = getSheet_();
  const statusCol = ensureStatusColumn_(sheet);

  // Dropdown for the whole Status column (below the header)
  const range = sheet.getRange(2, statusCol, sheet.getMaxRows() - 1, 1);
  range.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.STATUS_OPTIONS, true)
      .setAllowInvalid(false)
      .build()
  );

  // Colour-code each status
  const colours = {
    'New': '#fff2cc',
    'Contacted': '#cfe2f3',
    'Quote sent': '#d9d2e9',
    'Won': '#d9ead3',
    'Lost': '#f4cccc',
  };
  const rules = sheet.getConditionalFormatRules().filter(
    r => !r.getRanges().some(rg => rg.getColumn() === statusCol)
  );
  Object.keys(colours).forEach(status => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(status)
        .setBackground(colours[status])
        .setRanges([range])
        .build()
    );
  });
  sheet.setConditionalFormatRules(rules);

  // Header styling + freeze header row
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // Install the form-submit trigger (remove old ones first to avoid duplicates)
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onFormSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();

  Logger.log('Setup complete. Status column is column ' + statusCol + '.');
}

/** Runs automatically every time the form is submitted. */
function onFormSubmit(e) {
  const v = e.namedValues;
  const get = key => (v[CONFIG.FIELDS[key]] || [''])[0].toString().trim();

  const lead = {
    name: get('name'),
    email: get('email'),
    phone: get('phone'),
    company: get('company'),
    service: get('service'),
    budget: get('budget'),
    timeline: get('timeline'),
    details: get('details'),
    row: e.range.getRow(),
  };

  // 1. Set Status = "New" on this row
  const sheet = e.range.getSheet();
  const statusCol = ensureStatusColumn_(sheet);
  sheet.getRange(lead.row, statusCol).setValue('New');

  // 2. Thank-you email to the lead
  if (lead.email) {
    try {
      sendThankYou_(lead);
    } catch (err) {
      console.error('Thank-you email failed: ' + err);
    }
  }

  // 3. Alert email to you
  sendOwnerAlert_(lead, sheet);
}

// ====== Emails ======

function sendThankYou_(lead) {
  const firstName = (lead.name.split(' ')[0]) || 'there';
  const subject = `Thanks for your quote request, ${firstName}!`;
  const html = `
    <div style="font-family:Arial,sans-serif;max-width:560px;color:#222">
      <h2 style="color:#1f3864">Thanks, ${esc_(firstName)}! We've got your request.</h2>
      <p>We received your quote request for <b>${esc_(lead.service || 'your project')}</b>
         and will get back to you within <b>${CONFIG.REPLY_WITHIN}</b>.</p>
      <p style="margin-top:20px"><b>Here's what you sent us:</b></p>
      <table style="border-collapse:collapse;font-size:14px">
        ${row_('Service', lead.service)}
        ${row_('Budget', lead.budget)}
        ${row_('Timeline', lead.timeline)}
        ${row_('Details', lead.details)}
      </table>
      <p style="margin-top:20px">Need to add something? Just reply to this email.</p>
      <p>Best regards,<br><b>${esc_(CONFIG.BUSINESS_NAME)}</b></p>
    </div>`;

  MailApp.sendEmail({
    to: lead.email,
    subject: subject,
    htmlBody: html,
    name: CONFIG.BUSINESS_NAME,
    replyTo: CONFIG.OWNER_EMAIL,
  });
}

function sendOwnerAlert_(lead, sheet) {
  const url = sheet.getParent().getUrl() + '#gid=' + sheet.getSheetId() + '&range=A' + lead.row;
  const subject = `🔔 New lead: ${lead.name || 'Unknown'} – ${lead.service || 'Quote request'}`;
  const html = `
    <div style="font-family:Arial,sans-serif;max-width:560px;color:#222">
      <h2 style="color:#1f3864">New quote request</h2>
      <table style="border-collapse:collapse;font-size:14px">
        ${row_('Name', lead.name)}
        ${row_('Email', lead.email)}
        ${row_('Phone', lead.phone)}
        ${row_('Company', lead.company)}
        ${row_('Service', lead.service)}
        ${row_('Budget', lead.budget)}
        ${row_('Timeline', lead.timeline)}
        ${row_('Details', lead.details)}
      </table>
      <p style="margin-top:20px">
        <a href="${url}" style="background:#1f3864;color:#fff;padding:10px 16px;
           text-decoration:none;border-radius:4px">Open lead in Sheet (row ${lead.row})</a>
      </p>
    </div>`;

  MailApp.sendEmail({
    to: CONFIG.OWNER_EMAIL,
    subject: subject,
    htmlBody: html,
    replyTo: lead.email || CONFIG.OWNER_EMAIL,
  });
}

// ====== Helpers ======

function getSheet_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${CONFIG.SHEET_NAME}" not found. Check CONFIG.SHEET_NAME.`);
  return sheet;
}

/** Finds the Status column by header, creating it at the end if missing. */
function ensureStatusColumn_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf(CONFIG.STATUS_HEADER);
  if (idx !== -1) return idx + 1;
  const col = lastCol + 1;
  sheet.getRange(1, col).setValue(CONFIG.STATUS_HEADER);
  return col;
}

function row_(label, value) {
  if (!value) return '';
  return `<tr>
    <td style="padding:6px 12px 6px 0;color:#666;vertical-align:top"><b>${label}</b></td>
    <td style="padding:6px 0">${esc_(value).replace(/\n/g, '<br>')}</td>
  </tr>`;
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/** Optional: check your remaining daily email quota. */
function checkQuota() {
  Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota());
}
