/**
 * Lead capture + basic CRM in Google Sheets.
 * Replaces apps-script.gs. Works with landing-page.html and project-landing-sample.html.
 *
 * Setup (about 10 minutes):
 * 1. Create a Google Sheet. Extensions > Apps Script. Delete the default code and paste this file.
 * 2. Set NOTIFY_EMAIL (and OWNERS if you have a team) below.
 * 3. Run setup() once from the editor and accept the permissions. It builds the Leads sheet,
 *    status dropdown, colours, filters and a 9 am daily digest.
 * 4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 * 5. Paste the web app URL into CONFIG.formEndpoint in the HTML file.
 * After any later code change: Deploy > Manage deployments > Edit > New version.
 */

const NOTIFY_EMAIL = 'you@example.com';          // lead alerts and daily digest, comma separate for several
const OWNERS = ['Me'];                            // names for the Owner dropdown
const TZ = 'Asia/Kolkata';
const SHEET_NAME = 'Leads';
const DUPLICATE_DAYS = 30;                        // same phone within this many days is flagged
const DIGEST_HOUR = 9;                            // daily digest hour in the script time zone
const CRM_WEBHOOK_URL = '';                       // optional: each new lead is also POSTed here as JSON

const STATUSES = ['New', 'Contacted', 'Not reachable', 'Interested', 'Site visit scheduled',
                  'Site visit done', 'Negotiation', 'Booked', 'Lost'];
const CLOSED = ['Booked', 'Lost'];

// Column order. Keys with a "field" are filled from the form post.
const COLS = [
  { h: 'Timestamp' }, { h: 'Status' }, { h: 'Name', f: 'name' }, { h: 'Phone' },
  { h: 'Call' }, { h: 'WhatsApp' }, { h: 'Project', f: 'project' }, { h: 'Home type', f: 'unit' },
  { h: 'Area', f: 'area' }, { h: 'Budget', f: 'budget' }, { h: 'Buying for', f: 'purpose' },
  { h: 'Timeline', f: 'timeline' }, { h: 'Goal', f: 'goal' }, { h: 'Visit preference', f: 'visit' }, { h: 'Interest', f: 'interest' },
  { h: 'Source', f: 'utm_source' }, { h: 'Campaign', f: 'utm_campaign' }, { h: 'Keyword', f: 'utm_term' },
  { h: 'Ad group', f: 'utm_content' }, { h: 'Follow-up date' }, { h: 'Owner' }, { h: 'Notes' },
  { h: 'Site visit date' }, { h: 'Duplicate' },
  { h: 'gclid', f: 'gclid' }, { h: 'gbraid', f: 'gbraid' }, { h: 'fbclid', f: 'fbclid' },
  { h: 'Form', f: 'form' }, { h: 'Landing URL', f: 'landing_url' }, { h: 'Referrer', f: 'referrer' },
  { h: 'Event ID', f: 'event_id' }
];
const C = {}; COLS.forEach(function (c, i) { C[c.h] = i + 1; });   // header -> 1-based column

/* ---------------- web endpoint ---------------- */

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = (e && e.parameter) || {};
    const phone = String(p.phone || '').replace(/\D/g, '').slice(-10);
    if (!phone) return ContentService.createTextOutput('error');

    const sheet = getSheet_();
    const dup = isDuplicate_(sheet, phone);
    const row = sheet.getLastRow() + 1;
    const values = COLS.map(function (c) { return c.f ? clean_(p[c.f]) : ''; });
    values[C['Timestamp'] - 1] = new Date();
    values[C['Status'] - 1] = 'New';
    values[C['Phone'] - 1] = "'" + phone;
    values[C['Duplicate'] - 1] = dup ? 'Repeat enquiry' : '';
    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    sheet.getRange(row, C['Call']).setFormula('=HYPERLINK("tel:+91"&' + a1_(row, 'Phone') + ',"Call")');
    sheet.getRange(row, C['WhatsApp']).setFormula('=HYPERLINK("https://wa.me/91"&' + a1_(row, 'Phone') + ',"WhatsApp")');
    sheet.getRange(row, C['Timestamp']).setNumberFormat('dd mmm yyyy, hh:mm');

    notify_(p, phone, dup);
    forward_(p, phone, dup);
    return ContentService.createTextOutput('ok');
  } catch (err) {
    console.error(err);
    return ContentService.createTextOutput('error');
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return ContentService.createTextOutput('Lead endpoint is running.'); }

/* ---------------- setup and menu ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Leads')
    .addItem('Set up sheet and daily digest', 'setup')
    .addItem('Send digest now', 'dailyDigest')
    .addToUi();
}

function setup() {
  const sheet = getSheet_();
  const n = COLS.length;

  sheet.getRange(1, 1, 1, n).setValues([COLS.map(function (c) { return c.h; })])
    .setFontWeight('bold').setBackground('#0B2230').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  if (!sheet.getFilter()) sheet.getRange(1, 1, sheet.getMaxRows(), n).createFilter();

  const rows = sheet.getMaxRows() - 1;
  sheet.getRange(2, C['Status'], rows).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
  sheet.getRange(2, C['Owner'], rows).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(OWNERS, true).setAllowInvalid(true).build());
  sheet.getRange(2, C['Follow-up date'], rows).setNumberFormat('dd mmm yyyy');
  sheet.getRange(2, C['Site visit date'], rows).setNumberFormat('dd mmm yyyy');
  sheet.getRange(2, C['Phone'], rows).setNumberFormat('@');

  // colours by status, overdue follow-ups in red
  const stRange = sheet.getRange(2, C['Status'], rows);
  const colours = { 'New': '#FFF3CD', 'Contacted': '#E3F2FD', 'Not reachable': '#ECEFF1', 'Interested': '#D7F5E3',
    'Site visit scheduled': '#C8E6C9', 'Site visit done': '#A5D6A7', 'Negotiation': '#FFE0B2', 'Booked': '#66BB6A', 'Lost': '#EF9A9A' };
  const rules = Object.keys(colours).map(function (s) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colours[s]).setRanges([stRange]).build();
  });
  const fu = sheet.getRange(2, C['Follow-up date'], rows);
  const fuCol = colLetter_(C['Follow-up date']), stCol = colLetter_(C['Status']);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + fuCol + '2<>"",$' + fuCol + '2<TODAY(),$' + stCol + '2<>"Booked",$' + stCol + '2<>"Lost")')
    .setBackground('#F8B4B4').setRanges([fu]).build());
  sheet.setConditionalFormatRules(rules);

  const widths = { 'Timestamp': 140, 'Status': 140, 'Name': 150, 'Phone': 110, 'Call': 60, 'WhatsApp': 80, 'Notes': 280, 'Interest': 200 };
  Object.keys(widths).forEach(function (h) { sheet.setColumnWidth(C[h], widths[h]); });
  ['gclid', 'gbraid', 'fbclid', 'Form', 'Landing URL', 'Referrer', 'Event ID'].forEach(function (h) { sheet.hideColumns(C[h]); });

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyDigest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyDigest').timeBased().everyDays(1).atHour(DIGEST_HOUR).create();
  SpreadsheetApp.getActive().toast('Sheet ready. Deploy the web app next.', 'Leads', 6);
}

/* ---------------- daily digest ---------------- */

function dailyDigest() {
  if (!NOTIFY_EMAIL || NOTIFY_EMAIL === 'you@example.com') return;
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return;
  const data = sheet.getRange(2, 1, last - 1, COLS.length).getValues();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const due = [], fresh = [];

  data.forEach(function (r) {
    const status = r[C['Status'] - 1];
    if (!status || CLOSED.indexOf(status) > -1) return;
    const line = r[C['Name'] - 1] + ', ' + String(r[C['Phone'] - 1]).replace(/\D/g, '').slice(-10) +
      ' (' + status + (r[C['Budget'] - 1] ? ', ' + r[C['Budget'] - 1] : '') + ')';
    const f = r[C['Follow-up date'] - 1];
    if (f instanceof Date && f <= new Date(today.getTime() + 86399000)) due.push(line);
    if (status === 'New') fresh.push(line);
  });
  if (!due.length && !fresh.length) return;

  const body = [
    'Follow-ups due or overdue (' + due.length + '):', due.join('\n') || 'None', '',
    'Leads still marked New (' + fresh.length + '):', fresh.join('\n') || 'None', '',
    'Sheet: ' + SpreadsheetApp.getActive().getUrl()
  ].join('\n');
  MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: 'Lead digest: ' + due.length + ' follow-ups, ' + fresh.length + ' new', body: body });
}

/* ---------------- helpers ---------------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS.map(function (c) { return c.h; })]);
  }
  return sh;
}

function isDuplicate_(sheet, phone) {
  const last = sheet.getLastRow();
  if (last < 2) return false;
  const cutoff = Date.now() - DUPLICATE_DAYS * 86400000;
  const ts = sheet.getRange(2, C['Timestamp'], last - 1).getValues();
  const ph = sheet.getRange(2, C['Phone'], last - 1).getValues();
  for (let i = ph.length - 1; i >= 0; i--) {
    const t = ts[i][0] instanceof Date ? ts[i][0].getTime() : 0;
    if (t && t < cutoff) break;
    if (String(ph[i][0]).replace(/\D/g, '').slice(-10) === phone) return true;
  }
  return false;
}

function a1_(row, header) { return colLetter_(C[header]) + row; }

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Cap length and stop values that Sheets would read as formulas.
function clean_(v) {
  let s = String(v === undefined ? '' : v).slice(0, 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function notify_(p, phone, dup) {
  if (!NOTIFY_EMAIL || NOTIFY_EMAIL === 'you@example.com') return;
  const lines = [
    'Name: ' + clean_(p.name), 'Phone: +91 ' + phone,
    'Project: ' + clean_(p.project), 'Home type: ' + clean_(p.unit),
    'Budget: ' + clean_(p.budget), 'Area: ' + clean_(p.area),
    'Buying for: ' + clean_(p.purpose), 'Timeline: ' + clean_(p.timeline), 'Goal: ' + clean_(p.goal), 'Visit preference: ' + clean_(p.visit),
    'Interested in: ' + clean_(p.interest),
    'Campaign: ' + clean_(p.utm_campaign) + ' / ' + clean_(p.utm_term),
    dup ? 'Note: repeat enquiry within ' + DUPLICATE_DAYS + ' days' : '',
    'Call now: tel:+91' + phone, 'WhatsApp: https://wa.me/91' + phone
  ].filter(String);
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: (dup ? 'Repeat lead: ' : 'New lead: ') + clean_(p.name) + (p.budget ? ' (' + clean_(p.budget) + ')' : ''),
    body: lines.join('\n')
  });
}

function forward_(p, phone, dup) {
  if (!CRM_WEBHOOK_URL) return;
  try {
    UrlFetchApp.fetch(CRM_WEBHOOK_URL, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ name: p.name, phone: '+91' + phone, project: p.project, unit: p.unit, area: p.area,
        budget: p.budget, purpose: p.purpose, timeline: p.timeline, interest: p.interest,
        utm_source: p.utm_source, utm_campaign: p.utm_campaign, utm_term: p.utm_term, gclid: p.gclid, repeat: dup })
    });
  } catch (err) { console.error(err); }
}
