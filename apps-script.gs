/**
 * DPA Complaint Intake — Google Apps Script web app
 *
 * Receives POST submissions from the DPA site and appends each one as a row
 * in this Spreadsheet. Two tabs are created automatically:
 *   - Complaints   (Form DPA-1040)
 *   - Escalations  (Form DPA-1040-E)
 *
 * Deploy: Extensions → Apps Script → paste this in → Save → Deploy → New
 * deployment → Type: Web app → Execute as: Me, Who has access: Anyone →
 * Deploy → copy the Web app URL.
 */

// ID of the "DPA Intake" Google Sheet — pulled from its URL
const SPREADSHEET_ID  = '1RHa0WgFnJ9H9lScB60E3BhrXbL9-XtyjhcExws8Ex-o';
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
    return json({ ok: true, row: sheet.getLastRow() });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet() {
  return ContentService
    .createTextOutput('DPA intake endpoint is live. POST JSON here.')
    .setMimeType(ContentService.MimeType.TEXT);
}

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

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
