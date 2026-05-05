// ============================================================
// EGF Submission Handler — Google Apps Script
// Version 2.0
// ============================================================

const SHEET_NAME = 'Submissions';
const NOTIFY_EMAIL = 'josie.attard@xero.com';

const HEADERS = [
  'Ref','Submitted At','Status','App Name','App ID','Review Type',
  'Segment','Markets','JTBD','Connections','Active ABs','Partner Since',
  'Sentinel Tier','Problem Summary','Key Risks','Mitigants',
  'Recommendation','Options Considered','ARR at Risk','Partner Billed %',
  'Orgs 12m','Orgs 3m','Commercial Terms','Urgency','Target Date',
  'Supporting Links','Submitter Name','Submitter Email','Notes','Last Updated'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'submit') return handleSubmit(data);
    if (data.action === 'updateStatus') return handleUpdate(data, 'Status', data.status);
    if (data.action === 'updateNotes') return handleUpdate(data, 'Notes', data.notes);
    if (data.action === 'delete') return handleDelete(data);
    return respond({ success: false, error: 'Unknown action' });
  } catch(err) {
    return respond({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  if (e.parameter.action === 'getAll') return handleGetAll();
  return respond({ success: false, error: 'Use action=getAll' });
}

function handleSubmit(data) {
  const sheet = getOrCreateSheet();
  const ref = 'EGF-' + new Date().getFullYear() + '-' + (Math.floor(Math.random()*9000)+1000);
  const now = new Date().toISOString();
  sheet.appendRow([
    ref, now, 'New',
    data.appName||'', data.appId||'', data.type||'',
    data.segment||'', data.markets||'', data.jtbd||'',
    data.connections||'', data.abCount||'', data.partnerSince||'',
    data.sentinelTier||'', data.problem||'', data.risks||'',
    data.mitigants||'', data.recommendation||'', data.optionsConsidered||'',
    data.revRisk||'', data.partnerBilled||'', data.orgs12m||'', data.orgs3m||'',
    data.commercialTerms||'', data.urgency||'', data.targetDate||'',
    data.links||'', data.submitterName||'', data.submitterEmail||'',
    '', now
  ]);
  formatLastRow(sheet);
  sendEmail(ref, data);
  return respond({ success: true, ref });
}

function handleGetAll() {
  const sheet = getOrCreateSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return respond({ success: true, submissions: [] });
  const headers = rows[0];
  const submissions = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
  return respond({ success: true, submissions });
}

function handleUpdate(data, field, value) {
  const sheet = getOrCreateSheet();
  const row = findRow(sheet, data.ref);
  if (!row) return respond({ success: false, error: 'Ref not found' });
  sheet.getRange(row, HEADERS.indexOf(field)+1).setValue(value);
  sheet.getRange(row, HEADERS.indexOf('Last Updated')+1).setValue(new Date().toISOString());
  return respond({ success: true });
}

function handleDelete(data) {
  const sheet = getOrCreateSheet();
  const row = findRow(sheet, data.ref);
  if (!row) return respond({ success: false, error: 'Ref not found' });
  sheet.deleteRow(row);
  return respond({ success: true });
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const hRange = sheet.getRange(1, 1, 1, HEADERS.length);
    hRange.setValues([HEADERS]);
    hRange.setFontWeight('bold');
    hRange.setBackground('#1a1a1a');
    hRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    [120,140,100,160,200,110,90,120,200,90,80,100,110,400,300,300,400,300,150,100,80,80,160,120,100,300,150,180,400,140]
      .forEach((w,i) => sheet.setColumnWidth(i+1, w));
  }
  return sheet;
}

function formatLastRow(sheet) {
  const row = sheet.getLastRow();
  if (row % 2 === 0) sheet.getRange(row, 1, 1, HEADERS.length).setBackground('#f8f8f6');
  ['Problem Summary','Key Risks','Mitigants','Recommendation','Notes'].forEach(col => {
    const i = HEADERS.indexOf(col)+1;
    if (i > 0) sheet.getRange(row, i).setWrap(true);
  });
}

function findRow(sheet, ref) {
  const vals = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === ref) return i+1;
  }
  return null;
}

function sendEmail(ref, data) {
  try {
    const labels = { api_access:'API Access Review', sentinel:'Sentinel Pricing Exemption', managed_risk:'Managed Risk', other:'Other Decision' };
    const label = labels[data.type] || data.type;
    GmailApp.sendEmail(
      NOTIFY_EMAIL,
      `[EGF ${ref}] ${label} — ${data.appName}`,
      `New EGF submission received.\n\nRef: ${ref}\nApp: ${data.appName}\nType: ${label}\nSegment: ${data.segment||'—'}\nMarkets: ${data.markets||'—'}\nConnections: ${data.connections||'—'}\nUrgency: ${data.urgency||'—'}\nSubmitted by: ${data.submitterName} (${data.submitterEmail})\n\nPROBLEM\n${data.problem||'—'}\n\nRECOMMENDATION\n${data.recommendation||'—'}\n\nOpen the EGF Tracker to review: https://josieattard-bot.github.io/egf-governance/egf-tracker.html`
    );
  } catch(e) { console.error('Email failed:', e); }
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
