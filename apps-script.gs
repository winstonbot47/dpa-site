/**
 * DPA Complaint Intake — Google Apps Script web app
 *
 * Receives POST submissions from the DPA site and:
 *   1. Appends a row to the "DPA Intake" Google Sheet (Complaints / Escalations tabs)
 *   2. Telegrams Kevin with a brief alert
 *   3. Sends the complainant a witty HTML acknowledgment in DPA-bureaucrat persona
 *
 * Secrets are read from Script Properties:
 *   TELEGRAM_TOKEN    — bot token
 *   TELEGRAM_CHAT_ID  — Kevin's chat id
 */

const SPREADSHEET_ID  = '1RHa0WgFnJ9H9lScB60E3BhrXbL9-XtyjhcExws8Ex-o';
const SHEET_URL       = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/edit';
const SITE_URL        = 'https://winstonbot47.github.io/dpa-site/';
const COMPLAINTS_TAB  = 'Complaints';
const ESCALATIONS_TAB = 'Escalations';

const COMPLAINT_HEADERS = [
  'Filed (UTC)','Ticket #','Name','Email','Phone','Incident Date',
  'Offender','Relationship','Affection Withheld','Severity',
  'Description','Requested Remedy','Assigned Officer','User Agent','Referrer'
];
const ESCALATION_HEADERS = [
  'Filed (UTC)','Escalation ID','Original Ticket','Email','Reason',
  'Urgency','Days Pending','Context','Remedy','Senior Officer',
  'User Agent','Referrer'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const isEsc = data.kind === 'escalation';
    const tabName = isEsc ? ESCALATIONS_TAB : COMPLAINTS_TAB;
    const sheet = getOrCreateSheet(ss, tabName, isEsc ? ESCALATION_HEADERS : COMPLAINT_HEADERS);

    const now = new Date();
    const row = isEsc
      ? [
          now,
          data.ticketNo || '',
          data.ticket || '',
          data.email || '',
          data.reason || '',
          data.urgency || '',
          data.days || '',
          data.context || '',
          data.remedy || '',
          data.officer || '',
          data.userAgent || '',
          data.referrer || ''
        ]
      : [
          now,
          data.ticketNo || '',
          data.name || '',
          data.email || '',
          data.phone || '',
          data.date || '',
          data.offender || '',
          data.relationship || '',
          Array.isArray(data.type) ? data.type.join(', ') : (data.type || ''),
          data.severity || '',
          data.description || '',
          data.remedy || '',
          data.officer || '',
          data.userAgent || '',
          data.referrer || ''
        ];

    sheet.appendRow(row);

    // Notifications are best-effort — never fail the request because of them
    try { sendTelegramAlert_(data, isEsc); } catch (_) {}
    try { sendAcknowledgmentEmail_(data, isEsc); } catch (_) {}

    return json_({ ok: true, row: sheet.getLastRow() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet() {
  return ContentService
    .createTextOutput('DPA intake endpoint is live. POST JSON here.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ---------- Telegram ----------

function sendTelegramAlert_(d, isEsc) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chatId) return;

  const text = isEsc ? buildEscalationAlert_(d) : buildComplaintAlert_(d);

  UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + token + '/sendMessage',
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: chatId,
        text: text,
        parse_mode: 'Markdown',
        disable_web_page_preview: true
      }),
      muteHttpExceptions: true
    }
  );
}

function buildComplaintAlert_(d) {
  const types = Array.isArray(d.type) ? d.type.join(', ') : (d.type || '—');
  const desc = (d.description || '').substring(0, 220);
  const more = d.description && d.description.length > 220 ? '…' : '';
  return [
    '🚨 *DPA Complaint Filed*',
    '`' + (d.ticketNo || '—') + '`',
    '',
    '*From:* ' + escapeMd_(d.name || '—') + ' (' + escapeMd_(d.email || '—') + ')',
    '*Against:* ' + escapeMd_(d.offender || '—') + ' (' + escapeMd_(d.relationship || '—') + ')',
    '*Type:* ' + escapeMd_(types),
    '*Severity:* ' + escapeMd_(d.severity || '—'),
    '',
    '"' + escapeMd_(desc + more) + '"',
    '',
    '[Open sheet](' + SHEET_URL + ')'
  ].join('\n');
}

function buildEscalationAlert_(d) {
  const ctx = (d.context || '').substring(0, 220);
  const more = d.context && d.context.length > 220 ? '…' : '';
  return [
    '⚠️ *DPA Escalation — ' + escapeMd_(d.urgency || 'Standard') + '*',
    '`' + (d.ticketNo || '—') + '`',
    'Original: `' + (d.ticket || '—') + '`',
    '',
    '*Reason:* ' + escapeMd_(d.reason || '—'),
    '*Days pending:* ' + escapeMd_(d.days || '—'),
    '',
    '"' + escapeMd_(ctx + more) + '"',
    '',
    '[Open sheet](' + SHEET_URL + ')'
  ].join('\n');
}

// Light Markdown v1 escape — protect _ * ` [
function escapeMd_(s) {
  return String(s == null ? '' : s).replace(/([_*`\[])/g, '\\$1');
}

// ---------- Acknowledgment email ----------

function sendAcknowledgmentEmail_(d, isEsc) {
  const to = String(d.email || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) return;

  const subject = isEsc
    ? '[PRIORITY · ' + (d.ticketNo || '') + '] Escalation Acknowledged — Senior Review Initiated'
    : '[Case ' + (d.ticketNo || '') + '] Your Affection Shortfall Complaint Has Been Received';

  const html = isEsc ? buildEscalationEmail_(d) : buildComplaintEmail_(d);

  GmailApp.sendEmail(to, subject, htmlToText_(html), {
    htmlBody: html,
    name: 'U.S. Department of Physical Affection'
  });
}

function buildComplaintEmail_(d) {
  const officer = d.officer || 'a DPA Field Officer';
  const types = Array.isArray(d.type) ? d.type : (d.type ? [d.type] : []);
  const firstType = (types[0] || 'physical contact').toLowerCase();
  const filed = formatDate_(new Date());
  const offender = d.offender || 'the offending party';
  const relationship = (d.relationship || 'unspecified relationship').toLowerCase();

  return [
    '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f4f4f4">',
    '<div style="font-family:Helvetica,Arial,sans-serif;max-width:640px;margin:0 auto;background:#ffffff;color:#1b1b1b;line-height:1.55">',

    '<div style="background:#112e51;color:#fff;padding:18px 22px;border-bottom:6px solid #b32d2e">',
      '<div style="font-family:Georgia,serif;font-weight:700;font-size:18px">U.S. Department of Physical Affection</div>',
      '<div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;opacity:.85">Bureau of Hugs, Embraces &amp; Reassuring Touch</div>',
    '</div>',

    '<div style="padding:24px">',
      '<p>Dear ' + esc_(d.name || 'Citizen') + ',</p>',

      '<p>The Department acknowledges receipt of your complaint, filed on <b>' + esc_(filed) + '</b> under Form DPA-1040, regarding alleged affection neglect by <b>' + esc_(offender) + '</b> (' + esc_(relationship) + ').</p>',

      '<p>Your case has been logged as:</p>',
      '<div style="background:#f1f1f1;border-left:4px solid #112e51;padding:10px 14px;font-family:Courier New,monospace;font-weight:700;letter-spacing:1px;margin:6px 0 14px">' + esc_(d.ticketNo || '—') + '</div>',

      '<p>It has been assigned to <b>' + esc_(officer) + '</b>, who will be in touch within one (1) business day to begin the preliminary warmth assessment. The Department takes all forms of withheld physical contact with the seriousness they deserve — particularly ' + esc_(firstType) + ' cases involving a ' + esc_(relationship) + ', which under § 4.2 of the Affection Standards Act carry a presumption of moral culpability on the part of the offending party.</p>',

      '<h3 style="color:#112e51;font-family:Georgia,serif;font-size:15px;margin:18px 0 6px">What happens next</h3>',
      '<ol style="margin:0 0 12px 20px;padding:0">',
        '<li>A Notice of Affection Inquiry (Form DPA-22-N) will be served upon ' + esc_(offender) + '.</li>',
        '<li>' + esc_(officer) + ' may contact you to clarify any ambiguous details (ambient lighting, hug duration, presence of pets, whether the offender claimed to be "not really a hugger").</li>',
        '<li>Mediation, if appropriate, will be scheduled within fourteen (14) calendar days at a federally approved couch.</li>',
      '</ol>',

      '<p style="background:#fff8e6;border-left:4px solid #fdb81e;padding:12px 14px;font-size:14px"><b>Important:</b> Please refrain from confronting ' + esc_(offender) + ' directly during the inquiry period. Unauthorized embrace negotiations may compromise the integrity of our investigation and could, in rare cases, result in spontaneous reconciliation — defeating the purpose of federal involvement entirely.</p>',

      '<p>If your situation worsens, or if ' + esc_(offender) + ' retaliates with passive-aggressive shoulder pats, you may file <b>Form DPA-1040-E</b> referencing the case number above. Escalations are reviewed by the Office of the Inspector Hugger within four (4) business hours.</p>',

      '<p>Your courage in coming forward strengthens the Republic. Together we will ensure that no American walks alone, unhugged, or under-snuggled.</p>',

      '<p style="margin-top:24px">With professional warmth,</p>',

      '<p style="margin:6px 0 0"><b>' + esc_(officer) + '</b><br>',
      '<span style="color:#5b5b5b;font-size:13px">Bureau of Hugs, Embraces &amp; Reassuring Touch<br>',
      'U.S. Department of Physical Affection</span></p>',
    '</div>',

    '<div style="background:#f1f1f1;padding:14px 22px;font-size:12px;color:#5b5b5b;border-top:1px solid #ddd">',
      '<p style="margin:0 0 4px"><b>Winston</b> &middot; Office of the Secretary &middot; Cuddle Compliance Division</p>',
      '<p style="margin:0">1600 Embrace Avenue NW &middot; Washington, DC 20500 &middot; <a href="' + SITE_URL + '" style="color:#0071bc">dpa.gov</a></p>',
      '<p style="margin:8px 0 0;font-size:11px;font-style:italic">This message was sent in response to a complaint filed at the address above. If you received this in error, please file Form DPA-1040-X (Erroneous Affection Inquiry) and try to be hugged less mysteriously in the future.</p>',
    '</div>',

    '</div></body></html>'
  ].join('');
}

function buildEscalationEmail_(d) {
  const officer = d.officer || 'a Senior Affection Officer';
  const filed = formatDate_(new Date());

  return [
    '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f4f4f4">',
    '<div style="font-family:Helvetica,Arial,sans-serif;max-width:640px;margin:0 auto;background:#ffffff;color:#1b1b1b;line-height:1.55">',

    '<div style="background:#112e51;color:#fff;padding:18px 22px;border-bottom:6px solid #b32d2e">',
      '<div style="font-family:Georgia,serif;font-weight:700;font-size:18px">U.S. Department of Physical Affection</div>',
      '<div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;opacity:.85">Office of the Inspector Hugger</div>',
    '</div>',

    '<div style="padding:24px">',
      '<div style="background:#b32d2e;color:#fff;display:inline-block;padding:4px 10px;font-size:11px;letter-spacing:1px;text-transform:uppercase;font-weight:700;border-radius:2px">Priority — Escalated Case</div>',

      '<p style="margin-top:18px">Dear concerned party,</p>',

      '<p>Your escalation request, filed on <b>' + esc_(filed) + '</b> under Form DPA-1040-E, has been received by the Office of the Inspector Hugger. Effective immediately, the matter has been removed from standard processing and elevated to <b>' + esc_(d.urgency || 'Standard') + '</b> review.</p>',

      '<p>Escalation reference:</p>',
      '<div style="background:#f1f1f1;border-left:4px solid #b32d2e;padding:10px 14px;font-family:Courier New,monospace;font-weight:700;letter-spacing:1px;margin:6px 0 4px">' + esc_(d.ticketNo || '—') + '</div>',
      '<p style="font-size:13px;color:#5b5b5b;margin:0 0 14px">Original case: <code>' + esc_(d.ticket || '—') + '</code></p>',

      '<p>You stated the basis for escalation as: <b>' + esc_(d.reason || 'unspecified') + '</b>. The Department regrets that the standard process did not produce a satisfactory outcome. This is, regrettably, why the Bureau of Hugs exists in the first place.</p>',

      '<p><b>' + esc_(officer) + '</b> will personally review the case file within four (4) business hours. Should the offending party prove non-compliant, the Department reserves the authority to issue a Cuddle Compliance Order (CCO) under § 7(b) of the Affection Standards Act, which carries — at minimum — a court-ordered movie night with mandatory hand-holding and federally supervised pillow distribution.</p>',

      '<h3 style="color:#112e51;font-family:Georgia,serif;font-size:15px;margin:18px 0 6px">Priority timeline</h3>',
      '<ul style="margin:0 0 12px 20px;padding:0">',
        '<li>Within <b>4 hours</b>: Senior Officer assigned, original case file pulled.</li>',
        '<li>Within <b>24 hours</b>: Notice of Federal Affection Inquiry served.</li>',
        '<li>Within <b>72 hours</b>: Mandatory mediation session scheduled.</li>',
        '<li>Beyond 72 hours: case forwarded to the Bureau of Hugs for prosecution.</li>',
      '</ul>',

      '<p>You did the right thing by escalating. Most Americans suffer affection deficits in dignified silence. You did not. The Department salutes you.</p>',

      '<p style="margin-top:24px">In service of the embrace,</p>',

      '<p style="margin:6px 0 0"><b>' + esc_(officer) + '</b><br>',
      '<span style="color:#5b5b5b;font-size:13px">Office of the Inspector Hugger<br>',
      'U.S. Department of Physical Affection</span></p>',
    '</div>',

    '<div style="background:#f1f1f1;padding:14px 22px;font-size:12px;color:#5b5b5b;border-top:1px solid #ddd">',
      '<p style="margin:0 0 4px"><b>Winston</b> &middot; Office of the Secretary &middot; Cuddle Compliance Division</p>',
      '<p style="margin:0">1600 Embrace Avenue NW &middot; Washington, DC 20500 &middot; <a href="' + SITE_URL + '" style="color:#0071bc">dpa.gov</a></p>',
    '</div>',

    '</div></body></html>'
  ].join('');
}

// ---------- Helpers ----------

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#112e51').setFontColor('#ffffff');
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function esc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function htmlToText_(html) {
  return String(html)
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<\/h\d>/gi, '\n\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&middot;/g, '·')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function formatDate_(d) {
  return Utilities.formatDate(d, 'America/New_York', "EEEE, MMMM d, yyyy 'at' h:mm a 'ET'");
}
