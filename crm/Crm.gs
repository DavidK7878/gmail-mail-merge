/**
 * Email CRM for Google Sheets — entry points, menu, tab setup, shared helpers.
 *
 * Tabs: Master | Interactions | Import | Rules | Settings | Dashboard | Log
 *
 * Every function and file here is prefixed with Crm/crm so these files can live in the
 * same Apps Script project as the mail merge (Code.gs, Send.gs, …) without collisions.
 * See README.md → "Installing alongside the mail merge".
 */

var CRM_MASTER_COLUMNS = [
  // identity
  'Email', 'Name', 'FirstName', 'LastName', 'Company', 'Domain', 'Title',
  // classification
  'Category', 'Stage', 'Confidence', 'Sentiment', 'AISummary', 'NextAction', 'ClassifiedBy', 'ClassifiedAt',
  // yours to edit
  'Lock', 'Tags', 'Notes', 'DoNotContact',
  // facts from Gmail / imports
  'Source', 'FirstSentAt', 'LastSentAt', 'LastReplyAt', 'LastTouchAt',
  'SentCount', 'ReceivedCount', 'ThreadCount', 'NeedsReply', 'Bounced',
  'Subjects', 'LastOutboundSnippet', 'LastInboundSnippet', 'ThreadIds', 'UpdatedAt'
];

var CRM_INTERACTION_COLUMNS = [
  'ThreadId', 'Email', 'Subject', 'FirstMessageAt', 'LastMessageAt', 'Messages',
  'Outbound', 'Inbound', 'LastDirection', 'LastFrom', 'LastSnippet', 'Link', 'Participants'
];

var CRM_IMPORT_USER_COLUMNS = ['Email', 'Name', 'FirstName', 'LastName', 'Company', 'Title', 'Notes'];
var CRM_IMPORT_MACHINE_COLUMNS = ['DedupeStatus', 'MatchedName', 'MatchedCategory', 'MatchedStage', 'LastTouch', 'Recommendation', 'PushedAt'];

var CRM_RULE_COLUMNS = ['Type', 'Pattern', 'Category', 'Company', 'Notes'];

var CRM_DEFAULT_CATEGORIES = ['Investor', 'Engineer', 'Customer', 'Partner', 'Advisor', 'Press', 'Recruiter', 'Vendor', 'Personal', 'Other'];
var CRM_STAGES = ['New', 'Contacted', 'Replied', 'Engaged', 'Meeting', 'Won', 'Dormant', 'Not Interested', 'Bounced'];

var CRM_TIME_BUDGET_MS = 4.5 * 60 * 1000; // stay under the 6-minute execution cap
var CRM_API_KEY_PROP = 'CLAUDE_API_KEY';

// ---------------------------------------------------------------- Menu / UI

function onOpen() {
  buildCrmMenu_(SpreadsheetApp.getUi());
}

/** Adds the CRM menu. Call this from the mail merge's onOpen when both live in one project. */
function buildCrmMenu_(ui) {
  ui.createMenu('CRM')
    .addItem('Open sidebar', 'showCrmSidebar')
    .addSeparator()
    .addItem('Scan sent mail (new since last scan)', 'crmScanIncremental')
    .addItem('Full rescan from ScanSince', 'crmScanFull')
    .addItem('Stop scan', 'crmStopScan')
    .addSeparator()
    .addItem('Classify unclassified contacts', 'crmClassifyUnclassified')
    .addItem('Re-classify all (unlocked) contacts', 'crmClassifyAll')
    .addItem('Recompute pipeline stages', 'crmRecomputePipeline')
    .addItem('Sync from mail merge Contacts', 'crmSyncFromMailMerge')
    .addSeparator()
    .addItem('Import from Drive file…', 'crmImportFromDrivePrompt')
    .addItem('Check Import against Master', 'crmCheckImport')
    .addItem('Push NEW rows to Mail Merge', 'crmPushNewToMailMerge')
    .addItem('Clear Import tab', 'crmClearImport')
    .addSeparator()
    .addItem('Set Claude API key…', 'crmSetApiKeyPrompt')
    .addItem('Set up / repair CRM tabs', 'crmSetupSheet')
    .addItem('Stop all CRM triggers', 'crmStopEverything')
    .addToUi();
}

function showCrmSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('CrmSidebar').setTitle('CRM');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ---------------------------------------------------------------- Setup

/** Creates/repairs all CRM tabs. Safe to re-run; never deletes data. */
function crmSetupSheet() {
  var ss = SpreadsheetApp.getActive();

  crmEnsureHeaders_(ss, 'Master', CRM_MASTER_COLUMNS, CRM_MASTER_COLUMNS);
  crmEnsureHeaders_(ss, 'Interactions', CRM_INTERACTION_COLUMNS, CRM_INTERACTION_COLUMNS);
  crmEnsureHeaders_(ss, 'Import', CRM_IMPORT_USER_COLUMNS.concat(CRM_IMPORT_MACHINE_COLUMNS), CRM_IMPORT_MACHINE_COLUMNS);

  // Rules — seeded with examples the first time.
  var rules = ss.getSheetByName('Rules') || ss.insertSheet('Rules');
  if (rules.getLastRow() === 0) {
    var seed = [
      CRM_RULE_COLUMNS,
      ['domain', '.vc', 'Investor', '', 'Any address ending in .vc'],
      ['domain', 'a16z.com', 'Investor', 'Andreessen Horowitz', ''],
      ['domain', 'sequoiacap.com', 'Investor', 'Sequoia Capital', ''],
      ['domain', 'ycombinator.com', 'Investor', 'Y Combinator', ''],
      ['keyword', 'term sheet', 'Investor', '', 'Matched against subjects and excerpts'],
      ['keyword', 'pull request', 'Engineer', '', ''],
      ['domain', 'calendly.com', 'IGNORE', '', 'IGNORE drops the contact entirely'],
      ['email', 'notifications@github.com', 'IGNORE', '', '']
    ];
    rules.getRange(1, 1, seed.length, CRM_RULE_COLUMNS.length).setValues(seed);
    rules.getRange(1, 1, 1, CRM_RULE_COLUMNS.length).setFontWeight('bold');
    rules.setFrozenRows(1);
  }

  // Settings
  var settings = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  var since = new Date(); since.setMonth(since.getMonth() - 12);
  var defaults = [
    ['CompanyDescription', '', 'One paragraph about your startup: what you sell, to whom, and what you are raising. The classifier uses this to tell customers from partners.'],
    ['MyDomains', '', 'Comma-separated domains or addresses that count as "us" (co-founders, team aliases). Your own address is always included.'],
    ['ScanSince', Utilities.formatDate(since, Session.getScriptTimeZone(), 'yyyy-MM-dd'), 'Earliest date a full rescan reaches back to (yyyy-MM-dd).'],
    ['MaxThreadsPerRun', 250, 'Threads processed per trigger step. Lower if you see timeouts.'],
    ['IgnoreDomains', '', 'Comma-separated domains never added to Master (e.g. your bank, SaaS notification senders).'],
    ['Model', 'claude-opus-5', 'Claude model id used for classification.'],
    ['ClassifyBatchSize', 8, 'Contacts per API request.'],
    ['AutoClassifyAfterScan', true, 'Start classification automatically when a scan finishes.'],
    ['DormantDays', 45, 'Days without any touch before a Replied/Engaged/Meeting contact becomes Dormant.'],
    ['Categories', CRM_DEFAULT_CATEGORIES.join(', '), 'Allowed categories. Edit freely; the classifier and dashboard follow this list.'],
    ['SnippetChars', 400, 'Excerpt length stored per contact and sent to the classifier. 0 = metadata only.'],
    ['MailMergeSheetUrl', '', 'Blank = the "Contacts" tab in this spreadsheet. Otherwise the URL of the mail merge spreadsheet.'],
    ['ImportRecentDays', 30, 'Dedupe: a duplicate emailed within this many days is recommended "Skip".']
  ];
  if (settings.getLastRow() === 0) {
    settings.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'Notes']]);
    settings.getRange(1, 1, 1, 3).setFontWeight('bold');
    settings.setFrozenRows(1);
    settings.setColumnWidth(3, 520);
  }
  // Append any missing keys (lets the CRM share a Settings tab with the mail merge).
  var existingKeys = settings.getRange(1, 1, Math.max(settings.getLastRow(), 1), 1).getValues()
    .map(function (r) { return String(r[0]).trim(); });
  defaults.forEach(function (d) {
    if (existingKeys.indexOf(d[0]) === -1) settings.appendRow(d);
  });

  // Dashboard
  crmBuildDashboard_(ss);

  // Log
  var log = ss.getSheetByName('Log') || ss.insertSheet('Log');
  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Event', 'Target', 'Details']]);
    log.getRange(1, 1, 1, 4).setFontWeight('bold');
    log.setFrozenRows(1);
  }

  crmApplyValidation_(ss);
  crmLog_('CRM_SETUP', '', 'CRM tabs created/verified');
  try { SpreadsheetApp.getUi().alert('CRM tabs are ready.\n\n1. Fill Settings → CompanyDescription.\n2. CRM → Set Claude API key…\n3. CRM → Scan sent mail.'); } catch (e) { /* headless */ }
}

/** Create the sheet if missing; write headers if empty; append any missing `required` headers. */
function crmEnsureHeaders_(ss, name, allHeaders, requiredHeaders) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim(); }).filter(String);
  if (existing.length === 0) {
    sheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
    sheet.getRange(1, 1, 1, allHeaders.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }
  var lower = existing.map(function (h) { return h.toLowerCase(); });
  var missing = requiredHeaders.filter(function (h) { return lower.indexOf(h.toLowerCase()) === -1; });
  if (missing.length) {
    sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, existing.length + 1, 1, missing.length).setFontWeight('bold');
  }
  return sheet;
}

/** Dropdowns and checkboxes on Master so hand edits stay consistent. */
function crmApplyValidation_(ss) {
  try {
    var master = ss.getSheetByName('Master');
    var map = crmColMap_(master, CRM_MASTER_COLUMNS);
    var rows = Math.max(master.getMaxRows() - 1, 1);
    var cats = crmSettings_().Categories.concat(['Ignore']);
    master.getRange(2, map['category'], rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(cats, true).setAllowInvalid(true).build());
    master.getRange(2, map['stage'], rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(CRM_STAGES, true).setAllowInvalid(true).build());
    ['lock', 'donotcontact', 'needsreply', 'bounced'].forEach(function (h) {
      master.getRange(2, map[h], rows, 1).insertCheckboxes();
    });
    master.getRange(2, map['aisummary'], rows, 1).setWrap(false);
  } catch (e) { /* cosmetic */ }
}

function crmBuildDashboard_(ss) {
  var dash = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');
  if (dash.getLastRow() > 0) return; // don't clobber a customised dashboard; delete the tab to rebuild
  var cats = crmSettings_().Categories;

  // COUNTIF over a Master column found by header name (same trick as the mail merge Dashboard).
  var col = function (header) { return 'INDEX(Master!A:AZ,0,MATCH("' + header + '",Master!1:1,0))'; };
  var C = function (header, crit) { return '=IFERROR(COUNTIF(' + col(header) + ',' + crit + '),0)'; };

  var rows = [
    ['Metric', 'Value'],
    ['Contacts', '=IFERROR(COUNTA(' + col('Email') + ')-1,0)'],
    ['Unclassified', '=IFERROR(COUNTIFS(' + col('Email') + ',"<>",' + col('Category') + ',""),0)'],
    ['Needs reply', C('NeedsReply', 'TRUE')],
    ['Replied (any)', C('ReceivedCount', '">0"')],
    ['Reply rate', '=IFERROR(B5/COUNTIF(' + col('SentCount') + ',">0"),"")'],
    ['Dormant', C('Stage', '"Dormant"')],
    ['Meetings', C('Stage', '"Meeting"')],
    ['Won', C('Stage', '"Won"')],
    ['Bounced / Do not contact', '=IFERROR(COUNTIF(' + col('Bounced') + ',TRUE)+COUNTIF(' + col('DoNotContact') + ',TRUE),0)'],
    ['Last scan', '=IFERROR(MAX(' + col('UpdatedAt') + '),"")']
  ];
  dash.getRange(1, 1, rows.length, 2).setValues(rows);
  dash.getRange(1, 1, 1, 2).setFontWeight('bold');
  dash.getRange('B6').setNumberFormat('0.0%');
  dash.getRange('B11').setNumberFormat('yyyy-mm-dd hh:mm');

  // Category × Stage matrix
  var top = rows.length + 2;
  var header = ['Category \\ Stage'].concat(CRM_STAGES).concat(['Total', 'Reply rate']);
  var matrix = [header];
  cats.forEach(function (cat) {
    var line = [cat];
    CRM_STAGES.forEach(function (st) {
      line.push('=IFERROR(COUNTIFS(' + col('Category') + ',"' + cat + '",' + col('Stage') + ',"' + st + '"),0)');
    });
    line.push('=IFERROR(COUNTIF(' + col('Category') + ',"' + cat + '"),0)');
    line.push('=IFERROR(COUNTIFS(' + col('Category') + ',"' + cat + '",' + col('ReceivedCount') + ',">0")/COUNTIFS(' + col('Category') + ',"' + cat + '",' + col('SentCount') + ',">0"),"")');
    matrix.push(line);
  });
  dash.getRange(top, 1, matrix.length, header.length).setValues(matrix);
  dash.getRange(top, 1, 1, header.length).setFontWeight('bold');
  dash.getRange(top + 1, header.length, cats.length, 1).setNumberFormat('0.0%');
  dash.setColumnWidth(1, 200);
}

// ---------------------------------------------------------------- Settings

function crmSettings_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  if (!sheet) throw new Error('Settings tab missing — run CRM → Set up / repair CRM tabs.');
  var values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).getValues();
  var s = {};
  values.forEach(function (r) { var k = String(r[0]).trim(); if (k) s[k] = r[1]; });

  var out = {};
  out.CompanyDescription = String(s.CompanyDescription || '').trim();
  out.MyDomains = crmSplitList_(s.MyDomains).map(function (d) { return d.toLowerCase().replace(/^@/, ''); });
  out.ScanSince = crmParseDate_(s.ScanSince) || (function () { var d = new Date(); d.setMonth(d.getMonth() - 12); return d; })();
  out.MaxThreadsPerRun = Number(s.MaxThreadsPerRun) || 250;
  out.IgnoreDomains = crmSplitList_(s.IgnoreDomains).map(function (d) { return d.toLowerCase(); });
  out.Model = String(s.Model || 'claude-opus-5').trim();
  out.ClassifyBatchSize = Math.max(1, Math.min(Number(s.ClassifyBatchSize) || 8, 25));
  out.AutoClassifyAfterScan = s.AutoClassifyAfterScan === '' || s.AutoClassifyAfterScan === undefined ? true : crmBool_(s.AutoClassifyAfterScan);
  out.DormantDays = Number(s.DormantDays) || 45;
  out.Categories = crmSplitList_(s.Categories);
  if (!out.Categories.length) out.Categories = CRM_DEFAULT_CATEGORIES.slice();
  out.SnippetChars = s.SnippetChars === '' || s.SnippetChars === undefined ? 400 : Math.max(0, Number(s.SnippetChars) || 0);
  out.MailMergeSheetUrl = String(s.MailMergeSheetUrl || '').trim();
  out.ImportRecentDays = Number(s.ImportRecentDays) || 30;
  return out;
}

function crmSplitList_(v) {
  return String(v || '').split(',').map(function (x) { return x.trim(); }).filter(String);
}

function crmParseDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  var s = String(v || '').trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function crmBool_(v) {
  if (typeof v === 'boolean') return v;
  return /^(true|yes|1)$/i.test(String(v).trim());
}

// ---------------------------------------------------------------- API key

function crmSetApiKeyPrompt() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Claude API key', 'Paste your Anthropic API key (sk-ant-…). It is stored in this script\'s properties, not in any cell.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  crmSetApiKey(res.getResponseText());
  ui.alert('API key saved.');
}

function crmSetApiKey(key) {
  key = String(key || '').trim();
  if (!key) throw new Error('Empty API key.');
  PropertiesService.getScriptProperties().setProperty(CRM_API_KEY_PROP, key);
  crmLog_('API_KEY', '', 'Claude API key updated');
}

function crmGetApiKey_() {
  return PropertiesService.getScriptProperties().getProperty(CRM_API_KEY_PROP) || '';
}

// ---------------------------------------------------------------- Sheet access helpers

function crmSheet_(name) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error('"' + name + '" tab missing — run CRM → Set up / repair CRM tabs.');
  return sheet;
}

/** Map of lowercased header → 1-based column index; throws if any `required` header is missing. */
function crmColMap_(sheet, required) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function (h, i) {
    var key = String(h).trim().toLowerCase();
    if (key && !(key in map)) map[key] = i + 1;
  });
  (required || []).forEach(function (c) {
    if (!(c.toLowerCase() in map)) throw new Error('"' + sheet.getName() + '" tab is missing the "' + c + '" column — run CRM → Set up / repair CRM tabs.');
  });
  return map;
}

/**
 * Load a whole tab into memory: { sheet, headers, map, rows, index }.
 * `rows` are arrays aligned to `headers`; `index` maps a key (from keyFn) to row position.
 * Mutate rows via crmSet_/crmGet_, append via crmAppendRow_, then crmSaveTable_.
 */
function crmLoadTable_(name, required, keyFn) {
  var sheet = crmSheet_(name);
  var map = crmColMap_(sheet, required);
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  var lastRow = sheet.getLastRow();
  var rows = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  var t = { sheet: sheet, headers: headers, map: map, rows: rows, index: {}, dirty: false, appended: 0 };
  if (keyFn) {
    rows.forEach(function (r, i) {
      var k = keyFn(r, t);
      if (k && !(k in t.index)) t.index[k] = i;
    });
  }
  return t;
}

function crmGet_(t, row, header) {
  var c = t.map[header.toLowerCase()];
  return c ? row[c - 1] : '';
}

function crmSet_(t, row, header, value) {
  var c = t.map[header.toLowerCase()];
  if (!c) return;
  if (row[c - 1] !== value) { row[c - 1] = value; t.dirty = true; }
}

function crmAppendRow_(t, key) {
  var row = [];
  for (var i = 0; i < t.headers.length; i++) row.push('');
  t.rows.push(row);
  if (key) t.index[key] = t.rows.length - 1;
  t.dirty = true;
  t.appended++;
  return row;
}

function crmSaveTable_(t) {
  if (!t.dirty || !t.rows.length) return;
  t.sheet.getRange(2, 1, t.rows.length, t.headers.length).setValues(t.rows);
  t.dirty = false;
}

// ---------------------------------------------------------------- Email helpers (pure)

/** Lowercase, trim, canonical Gmail (dots and +tags removed). Empty string if not an email. */
function crmNormalizeEmail_(raw) {
  var e = String(raw || '').trim().toLowerCase();
  var m = e.match(/[a-z0-9._%+\-']+@[a-z0-9.\-]+\.[a-z]{2,}/);
  if (!m) return '';
  e = m[0];
  var parts = e.split('@');
  var local = parts[0], domain = parts[1];
  if (domain === 'googlemail.com') domain = 'gmail.com';
  if (domain === 'gmail.com') local = local.split('+')[0].replace(/\./g, '');
  return local + '@' + domain;
}

function crmDomainOf_(email) {
  var at = String(email || '').lastIndexOf('@');
  return at === -1 ? '' : String(email).slice(at + 1).toLowerCase();
}

var CRM_FREEMAIL = { 'gmail.com': 1, 'yahoo.com': 1, 'outlook.com': 1, 'hotmail.com': 1, 'live.com': 1, 'icloud.com': 1, 'me.com': 1, 'aol.com': 1, 'protonmail.com': 1, 'proton.me': 1, 'hey.com': 1, 'msn.com': 1, 'ymail.com': 1, 'gmx.com': 1, 'fastmail.com': 1, 'yahoo.co.uk': 1, 'googlemail.com': 1 };
function crmIsFreemail_(domain) { return !!CRM_FREEMAIL[String(domain || '').toLowerCase()]; }

/**
 * Parse an RFC-style address list ("Jane Doe <jane@x.com>, bob@y.com") into
 * [{ name, email }] with normalized emails. Handles quoted names with commas.
 */
function crmParseAddresses_(headerValue) {
  var out = [];
  var s = String(headerValue || '');
  var parts = [], cur = '', inQ = false, depth = 0;
  for (var i = 0; i < s.length; i++) {
    var ch = s[i];
    if (ch === '"') inQ = !inQ;
    else if (!inQ && ch === '<') depth++;
    else if (!inQ && ch === '>') depth = Math.max(0, depth - 1);
    if (ch === ',' && !inQ && depth === 0) { parts.push(cur); cur = ''; } else cur += ch;
  }
  if (cur.trim()) parts.push(cur);
  parts.forEach(function (p) {
    p = p.trim(); if (!p) return;
    var m = p.match(/^(.*?)\s*<([^>]+)>\s*$/);
    var name = '', email = '';
    if (m) { name = m[1].trim().replace(/^"|"$/g, '').trim(); email = m[2]; } else { email = p; }
    var norm = crmNormalizeEmail_(email);
    if (!norm) return;
    if (name && name.toLowerCase() === email.trim().toLowerCase()) name = '';
    out.push({ name: name, email: norm, raw: email.trim().toLowerCase() });
  });
  return out;
}

/** "Jane Q. Doe" → { first: "Jane", last: "Doe" }. Handles "Doe, Jane". */
function crmSplitName_(name) {
  var n = String(name || '').replace(/["']/g, '').replace(/\s+/g, ' ').trim();
  if (!n) return { first: '', last: '' };
  if (n.indexOf(',') !== -1) { var p = n.split(','); return { first: p[1].trim().split(' ')[0] || '', last: p[0].trim() }; }
  var parts = n.split(' ');
  if (parts.length === 1) return { first: parts[0], last: '' };
  return { first: parts[0], last: parts[parts.length - 1] };
}

/** Addresses that are never real contacts. */
var CRM_MACHINE_LOCAL_RE = /^(no-?reply|noreply|do-?not-?reply|donotreply|mailer-daemon|postmaster|bounce[s]?|notifications?|notify|alerts?|calendar-notification|drive-shares(-dm)?-noreply|comments-noreply|forms-receipts-noreply|docs-noreply|meet-recordings-noreply|unsubscribe|newsletter|digest|automated|robot|bot|system|daemon)(\+.*)?$/i;
var CRM_MACHINE_DOMAIN_RE = /(^|\.)(calendar\.google\.com|docs\.google\.com|google\.com|googlegroups\.com|linkedin\.com|facebookmail\.com|twitter\.com|x\.com|intercom-mail\.com|hubspot\.com|zendesk\.com|calendly\.com|zoom\.us|slack\.com|notion\.so|stripe\.com|amazonses\.com|sendgrid\.net|mailchimp\.com|mailgun\.org|substack\.com|medium\.com|github\.com|atlassian\.net|figma\.com|docusign\.net|dropbox\.com)$/i;

function crmIsMachineAddress_(email, ignoreDomains) {
  var e = String(email || '').toLowerCase();
  if (!e) return true;
  var local = e.split('@')[0], domain = crmDomainOf_(e);
  if (CRM_MACHINE_LOCAL_RE.test(local)) return true;
  if (CRM_MACHINE_DOMAIN_RE.test(domain)) return true;
  if (/^(reply|r|msg|m|bounce|b)[-+._][a-z0-9]{8,}/.test(local)) return true; // reply-tokens from SaaS tools
  for (var i = 0; i < (ignoreDomains || []).length; i++) {
    var d = ignoreDomains[i];
    if (domain === d || domain.slice(-(d.length + 1)) === '.' + d) return true;
  }
  return false;
}

/**
 * Plain-text excerpt: strip quoted history, signatures, and whitespace; clip to maxChars.
 */
function crmExcerpt_(plainBody, maxChars) {
  var text = String(plainBody || '').replace(/\r/g, '');
  // Cut at the start of quoted history.
  var cutters = [
    /\n\s*On [\s\S]{5,200}?wrote:\s*\n/,     // "On Mon, Jan 1 … <x@y> wrote:" (may wrap)
    /\n\s*-{2,}\s*Original Message\s*-{2,}/i,
    /\n\s*From:\s.+\n\s*(Sent|Date):\s.+/i,  // Outlook-style header block
    /\n\s*_{5,}\s*\n/,                       // Outlook separator
    /\n\s*Le .{5,200}? a écrit\s*:/,         // French
    /\n\s*Am .{5,200}? schrieb .{0,80}:/     // German
  ];
  cutters.forEach(function (re) { var m = text.match(re); if (m && m.index > 0) text = text.slice(0, m.index); });
  // Drop quoted lines and the signature block.
  var lines = text.split('\n').filter(function (l) { return !/^\s*>/.test(l); });
  var sig = -1;
  for (var i = 0; i < lines.length; i++) {
    if (/^\s*(--|—|__)\s*$/.test(lines[i]) || /^\s*(Sent from my|Get Outlook for)/i.test(lines[i])) { sig = i; break; }
  }
  if (sig > 0) lines = lines.slice(0, sig);
  text = lines.join('\n').replace(/\[image:[^\]]*\]/gi, '').replace(/[ \t]+/g, ' ').replace(/\n{2,}/g, '\n').trim();
  if (maxChars === 0) return '';
  var max = maxChars || 400;
  if (text.length > max) text = text.slice(0, max).replace(/\s+\S*$/, '') + '…';
  return text;
}

// ---------------------------------------------------------------- Logging / triggers

function crmLog_(event, target, details) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('Log');
    if (sheet) sheet.appendRow([new Date(), event, target || '', String(details || '').slice(0, 2000)]);
  } catch (e) { /* never let logging break the pipeline */ }
}

/** Non-blocking feedback for menu-invoked actions (harmless from triggers/sidebar). */
function crmToast_(msg) {
  try { SpreadsheetApp.getActive().toast(String(msg), 'CRM', 8); } catch (e) { /* headless */ }
}

function crmMyEmail_() {
  return Session.getEffectiveUser().getEmail().toLowerCase();
}

function crmStopEverything() {
  ['crmScanStep', 'crmClassifyStep'].forEach(crmDeleteTriggersFor_);
  PropertiesService.getScriptProperties().deleteProperty('CRM_SCAN_STATE');
  PropertiesService.getScriptProperties().deleteProperty('CRM_CLASSIFY_STATE');
  crmLog_('CRM_TRIGGERS', '', 'All CRM triggers removed');
  try { SpreadsheetApp.getUi().alert('CRM triggers removed. Scan and classification jobs are stopped.'); } catch (e) { }
}

function crmDeleteTriggersFor_(fnName) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === fnName) ScriptApp.deleteTrigger(t);
  });
}

function crmHasTriggerFor_(fnName) {
  return ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === fnName; });
}

function crmEnsureMinuteTrigger_(fnName) {
  if (!crmHasTriggerFor_(fnName)) ScriptApp.newTrigger(fnName).timeBased().everyMinutes(1).create();
}

function crmMaxDate_(a, b) {
  var da = a instanceof Date ? a : null, db = b instanceof Date ? b : null;
  if (!da) return db; if (!db) return da;
  return da.getTime() >= db.getTime() ? da : db;
}

function crmMinDate_(a, b) {
  var da = a instanceof Date ? a : null, db = b instanceof Date ? b : null;
  if (!da) return db; if (!db) return da;
  return da.getTime() <= db.getTime() ? da : db;
}

function crmDaysAgo_(d) {
  return d instanceof Date ? (Date.now() - d.getTime()) / 86400000 : Infinity;
}

/** Comma-joined set add, capped so cells stay readable. */
function crmAddToList_(existing, value, cap) {
  var list = crmSplitList_(existing);
  var v = String(value || '').trim();
  if (v && list.indexOf(v) === -1) { list.push(v); if (cap && list.length > cap) list = list.slice(-cap); }
  return list.join(', ');
}
