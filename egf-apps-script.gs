// ============================================================
// EGF Submission Handler — Google Apps Script
// Paste this entire file into your Apps Script editor
// See setup instructions in egf-setup-guide.html
// ============================================================

const SHEET_NAME = 'Submissions';
const NOTIFY_EMAIL = 'josie.attard@xero.com'; // Change if needed

// Column headers — DO NOT reorder without updating the code below
const HEADERS = [
  'Ref', 'Submitted At', 'Status', 'App Name', 'App ID', 'Review Type',
  'Segment', 'Markets', 'JTBD', 'Connections', 'Active ABs', 'Partner Since',
  'Sentinel Tier', 'Problem Summary', 'Key Risks', 'Mitigants',
  'Recommendation', 'Options Considered', 'ARR at Risk', 'Partner Billed %',
  'Orgs 12m', 'Orgs 3m', 'Commercial Terms', 'Urgency', 'Target Date',
  'Supporting Links', 'Submitter Name', 'Submitter Email', 'Notes', 'Last Updated'
];

// ── MAIN ENTRY POINT ──────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'submit') return handleSubmit(data);
    if (action === 'getAll') return handleGetAll();
    if (action === 'updateStatus') return handleUpdateStatus(data);
    if (action === 'updateNotes') return handleUpdateNotes(data);
    if (action === 'delete') return handleDelete(data);

    return respond({ success: false, error: 'Unknown action' });
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  // Allow GET for fetching all submissions (easier for tracker to call)
  const action = e.parameter.action;
  if (action === 'getAll') return handleGetAll();
  return respond({ success: false, error: 'GET only supports getAll' });
}

// ── HANDLERS ─────────────────────────────────────────────

function handleSubmit(data) {
  const sheet = getOrCreateSheet();
  const ref = generateRef();
  const now = new Date().toISOString();

  const row = [
    ref,                          // Ref
    now,                          // Submitted At
    'New',                        // Status
    data.appName || '',
    data.appId || '',
    data.type || '',
    data.segment || '',
    data.markets || '',
    data.jtbd || '',
    data.connections || '',
    data.abCount || '',
    data.partnerSince || '',
    data.sentinelTier || '',
    data.problem || '',
    data.risks || '',
    data.mitigants || '',
    data.recommendation || '',
    data.optionsConsidered || '',
    data.revRisk || '',
    data.partnerBilled || '',
    data.orgs12m || '',
    data.orgs3m || '',
    data.commercialTerms || '',
    data.urgency || '',
    data.targetDate || '',
    data.links || '',
    data.submitterName || '',
    data.submitterEmail || '',
    '',                           // Notes (empty on create)
    now                           // Last Updated
  ];

  sheet.appendRow(row);
  formatLastRow(sheet);
  sendNotificationEmail(ref, data);

  return respond({ success: true, ref });
}

function handleGetAll() {
  const sheet = getOrCreateSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return respond({ success: true, submissions: [] });

  const headers = rows[0];
  const submissions = rows.slice(1).map((row, i) => {
    const obj = { _row: i + 2 }; // 1-indexed, +1 for header
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  });

  return respond({ success: true, submissions });
}

function handleUpdateStatus(data) {
  const sheet = getOrCreateSheet();
  const row = findRowByRef(sheet, data.ref);
  if (!row) return respond({ success: false, error: 'Ref not found' });

  const statusCol = HEADERS.indexOf('Status') + 1;
  const updatedCol = HEADERS.indexOf('Last Updated') + 1;
  sheet.getRange(row, statusCol).setValue(data.status);
  sheet.getRange(row, updatedCol).setValue(new Date().toISOString());

  return respond({ success: true });
}

function handleUpdateNotes(data) {
  const sheet = getOrCreateSheet();
  const row = findRowByRef(sheet, data.ref);
  if (!row) return respond({ success: false, error: 'Ref not found' });

  const notesCol = HEADERS.indexOf('Notes') + 1;
  const updatedCol = HEADERS.indexOf('Last Updated') + 1;
  sheet.getRange(row, notesCol).setValue(data.notes);
  sheet.getRange(row, updatedCol).setValue(new Date().toISOString());

  return respond({ success: true });
}

function handleDelete(data) {
  const sheet = getOrCreateSheet();
  const row = findRowByRef(sheet, data.ref);
  if (!row) return respond({ success: false, error: 'Ref not found' });
  sheet.deleteRow(row);
  return respond({ success: true });
}

// ── HELPERS ──────────────────────────────────────────────

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a1a1a');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontSize(11);
    sheet.setFrozenRows(1);

    // Set column widths
    const widths = [100,140,110,160,220,120,100,120,180,90,80,90,120,400,300,300,400,300,140,100,80,80,160,120,100,300,140,180,400,140];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }

  return sheet;
}

function formatLastRow(sheet) {
  const lastRow = sheet.getLastRow();
  const statusCell = sheet.getRange(lastRow, HEADERS.indexOf('Status') + 1);

  // Alternate row shading
  if (lastRow % 2 === 0) {
    sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground('#f8f8f6');
  }

  // Wrap text on long columns
  ['Problem Summary','Key Risks','Mitigants','Recommendation','Options Considered','Notes'].forEach(col => {
    const idx = HEADERS.indexOf(col) + 1;
    if (idx > 0) sheet.getRange(lastRow, idx).setWrap(true);
  });
}

function findRowByRef(sheet, ref) {
  const refCol = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (let i = 1; i < refCol.length; i++) {
    if (refCol[i][0] === ref) return i + 1; // 1-indexed
  }
  return null;
}

function generateRef() {
  const year = new Date().getFullYear();
  const rand = Math.floor(Math.random() * 9000) + 1000;
  return `EGF-${year}-${rand}`;
}

function sendNotificationEmail(ref, data) {
  try {
    const typeLabels = {
      api_access: 'API Access Review',
      sentinel: 'Sentinel Pricing Exemption',
      managed_risk: 'Managed Risk',
      other: 'Other Decision'
    };
    const subject = `[EGF ${ref}] New submission — ${typeLabels[data.type] || data.type} — ${data.appName}`;
    const body = `A new EGF submission has been received.\n\n` +
      `Reference: ${ref}\n` +
      `App: ${data.appName} (${data.appId || 'ID not provided'})\n` +
      `Review type: ${typeLabels[data.type] || data.type}\n` +
      `Segment: ${data.segment || '—'}\n` +
      `Markets: ${data.markets || '—'}\n` +
      `Connections: ${data.connections || '—'}\n` +
      `JTBD: ${data.jtbd || '—'}\n` +
      `Urgency: ${data.urgency || '—'}\n` +
      `Target decision: ${data.targetDate || '—'}\n` +
      `Submitted by: ${data.submitterName} (${data.submitterEmail})\n\n` +
      `PROBLEM SUMMARY\n${data.problem || '—'}\n\n` +
      `RECOMMENDATION\n${data.recommendation || '—'}\n\n` +
      `KEY RISKS\n${data.risks || '—'}\n\n` +
      `Open the EGF Tracker to review and triage this submission.`;

    GmailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch (err) {
    // Email failure shouldn't break the submission
    console.error('Email notification failed:', err);
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
