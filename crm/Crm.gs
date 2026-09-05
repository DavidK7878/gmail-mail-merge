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
  'Category', 'Stage', 'Confidence', 'ReviewNeeded', 'Evidence', 'Sentiment', 'AISummary', 'NextAction', 'ClassifiedBy', 'ClassifiedAt',
  // yours to edit
  'Lock', 'Tags', 'Notes', 'DoNotContact',
  // facts from Gmail / imports
  'Source', 'FirstSentAt', 'LastSentAt', 'LastReplyAt', 'LastTouchAt',
  'SentCount', 'ReceivedCount', 'AutoReplies', 'ThreadCount', 'NeedsReply', 'Bounced',
  'Subjects', 'LastOutboundSnippet', 'LastInboundSnippet', 'Signature', 'ThreadIds', 'UpdatedAt'
];

var CRM_INTERACTION_COLUMNS = [
  'ThreadId', 'Email', 'Subject', 'FirstMessageAt', 'LastMessageAt', 'Messages',
  'Outbound', 'Inbound', 'AutoReplies', 'LastDirection', 'LastFrom', 'LastSnippet', 'LastInboundSnippet', 'Link', 'Participants'
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
    .addItem('Preview classification (5 contacts, nothing saved)', 'crmPreviewClassificationPrompt')
    .addItem('Recompute pipeline stages', 'crmRecomputePipeline')
    .addItem('Audit Master integrity', 'crmAuditPrompt')
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

function crmPreviewClassificationPrompt() {
  var ui = SpreadsheetApp.getUi();
  try {
    var rows = crmPreviewClassification(5);
    ui.alert('Preview (not saved)', rows.map(function (r) {
      return r.email + '\n  ' + r.category + ' (' + Math.round(r.confidence * 100) + '%) · ' + r.relationship + (r.review ? ' · REVIEW' : '') + '\n  evidence: ' + r.evidence + '\n  ' + r.summary;
    }).join('\n\n'), ui.ButtonSet.OK);
  } catch (e) { ui.alert(String(e.message || e)); }
}

function crmAuditPrompt() {
  var ui = SpreadsheetApp.getUi();
  var rep = crmAudit();
  var lines = rep.findings.map(function (f) { return '• ' + f.check + ': ' + f.count + (f.note ? ' — ' + f.note : '') + '\n    e.g. ' + f.examples.join(', '); });
  ui.alert('Master audit', rep.summary + (lines.length ? '\n\n' + lines.join('\n') : '\n\nNo issues found.'), ui.ButtonSet.OK);
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
    ['ImportRecentDays', 30, 'Dedupe: a duplicate emailed within this many days is recommended "Skip".'],
    ['MinConfidence', 0.7, 'AI results below this confidence (0-1) get ReviewNeeded = TRUE and, if RetryLowConfidence is on, a second pass at higher effort.'],
    ['RetryLowConfidence', true, 'Re-run low-confidence contacts one at a time at effort "high" with the full evidence pack.'],
    ['EvidenceThreads', 3, 'How many recent threads per contact are sent to the classifier as evidence.'],
    ['TreatMyDomainAsTeam', true, 'Treat everyone at your own (non-freemail) email domain as a teammate, not a contact. Turn off if you email customers at your own domain.'],
    ['ScanWindowDays', 7, 'The scan walks the date range in windows of this many days so paging stays stable while new mail arrives. Lower to 1 if you send hundreds of emails a day.']
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
    ['lock', 'donotcontact', 'needsreply', 'bounced', 'reviewneeded'].forEach(function (h) {
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
    ['Needs review (low confidence)', C('ReviewNeeded', 'TRUE')],
    ['Needs reply', C('NeedsReply', 'TRUE')],
    ['Replied (any)', C('ReceivedCount', '">0"')],
    ['Reply rate', '=IFERROR(B6/COUNTIF(' + col('SentCount') + ',">0"),"")'],
    ['Dormant', C('Stage', '"Dormant"')],
    ['Meetings', C('Stage', '"Meeting"')],
    ['Won', C('Stage', '"Won"')],
    ['Bounced / Do not contact', '=IFERROR(COUNTIF(' + col('Bounced') + ',TRUE)+COUNTIF(' + col('DoNotContact') + ',TRUE),0)'],
    ['Last scan', '=IFERROR(MAX(' + col('UpdatedAt') + '),"")']
  ];
  dash.getRange(1, 1, rows.length, 2).setValues(rows);
  dash.getRange(1, 1, 1, 2).setFontWeight('bold');
  dash.getRange('B7').setNumberFormat('0.0%');
  dash.getRange('B12').setNumberFormat('yyyy-mm-dd hh:mm');

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
  out.ClassifyBatchSize = Math.max(1, Math.min(Number(s.ClassifyBatchSize) || 8, 15));
  out.AutoClassifyAfterScan = s.AutoClassifyAfterScan === '' || s.AutoClassifyAfterScan === undefined ? true : crmBool_(s.AutoClassifyAfterScan);
  out.DormantDays = Number(s.DormantDays) || 45;
  out.Categories = crmSplitList_(s.Categories);
  if (!out.Categories.length) out.Categories = CRM_DEFAULT_CATEGORIES.slice();
  out.SnippetChars = s.SnippetChars === '' || s.SnippetChars === undefined ? 400 : Math.max(0, Number(s.SnippetChars) || 0);
  out.MailMergeSheetUrl = String(s.MailMergeSheetUrl || '').trim();
  out.ImportRecentDays = Number(s.ImportRecentDays) || 30;
  out.MinConfidence = s.MinConfidence === '' || s.MinConfidence === undefined ? 0.7 : Math.max(0, Math.min(1, Number(s.MinConfidence) || 0));
  out.RetryLowConfidence = s.RetryLowConfidence === '' || s.RetryLowConfidence === undefined ? true : crmBool_(s.RetryLowConfidence);
  out.EvidenceThreads = Math.max(0, Math.min(Number(s.EvidenceThreads === '' || s.EvidenceThreads === undefined ? 3 : s.EvidenceThreads) || 0, 6));
  out.TreatMyDomainAsTeam = s.TreatMyDomainAsTeam === '' || s.TreatMyDomainAsTeam === undefined ? true : crmBool_(s.TreatMyDomainAsTeam);
  out.ScanWindowDays = Math.max(1, Math.min(Number(s.ScanWindowDays) || 7, 60));
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
 * Load a whole tab into memory: { sheet, headers, map, rows, index, keyFn }.
 * `rows` are arrays aligned to `headers`; `index` maps keyFn(row) to row position.
 * Mutate rows via crmSet_ (tracks which cells changed), append via crmAppendRow_,
 * then crmSaveTable_ merges the changes back into whatever the sheet holds *now*.
 */
function crmLoadTable_(name, required, keyFn) {
  var sheet = crmSheet_(name);
  var map = crmColMap_(sheet, required);
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  var lastRow = sheet.getLastRow();
  var rows = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  var t = { sheet: sheet, headers: headers, map: map, rows: rows, index: {}, keyFn: keyFn || null, dirty: false, appended: 0 };
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

/** Set a cell in memory and remember that this job changed it (row._d = {colIndex: true}). */
function crmSet_(t, row, header, value) {
  var c = t.map[header.toLowerCase()];
  if (!c) return;
  if (row[c - 1] !== value) {
    row[c - 1] = value;
    if (!row._d) row._d = {};
    row._d[c] = true;
    t.dirty = true;
  }
}

function crmAppendRow_(t, key) {
  var row = [];
  for (var i = 0; i < t.headers.length; i++) row.push('');
  row._new = true;
  t.rows.push(row);
  if (key) t.index[key] = t.rows.length - 1;
  t.dirty = true;
  t.appended++;
  return row;
}

/** Columns a human edit + Lock protects. If Lock was ticked while a job ran, the job's values for these are dropped. */
var CRM_LOCK_PROTECTED = ['Category', 'Stage', 'Company', 'Title', 'Name', 'FirstName', 'LastName', 'Confidence', 'ReviewNeeded', 'Evidence', 'Sentiment', 'AISummary', 'NextAction', 'ClassifiedBy', 'ClassifiedAt', 'Tags'];

/**
 * Write changes back. Jobs run for minutes, so the sheet may have been edited, sorted, or
 * had rows deleted meanwhile. With a keyFn, the current sheet contents are re-read and only
 * the cells this job changed are applied to the matching (by key) current rows; new rows are
 * appended. Rows the user deleted stay deleted unless this job changed them. A Lock ticked
 * during the job wins over the job's classification columns. Without a keyFn the table is
 * written positionally (only used for the short-lived Import tab).
 */
function crmSaveTable_(t) {
  if (!t.dirty || !t.rows.length) return;
  var width = t.headers.length;
  var changed = t.rows.filter(function (r) { return r._d || r._new; });

  if (!t.keyFn) {
    t.sheet.getRange(2, 1, t.rows.length, width).setValues(t.rows.map(function (r) { return r.slice(0, width); }));
  } else {
    var lastRow = t.sheet.getLastRow();
    var fresh = lastRow >= 2 ? t.sheet.getRange(2, 1, lastRow - 1, width).getValues() : [];
    var freshIndex = {};
    fresh.forEach(function (r, i) { var k = t.keyFn(r, t); if (k && !(k in freshIndex)) freshIndex[k] = i; });
    var lockCol = t.map['lock'];
    var protectedCols = {};
    if (lockCol) CRM_LOCK_PROTECTED.forEach(function (h) { var c = t.map[h.toLowerCase()]; if (c) protectedCols[c] = true; });

    var appended = [];
    changed.forEach(function (row) {
      var key = t.keyFn(row, t);
      var j = key ? freshIndex[key] : undefined;
      if (j === undefined) { appended.push(row.slice(0, width)); return; }
      var target = fresh[j];
      var lockedMeanwhile = lockCol && crmBool_(target[lockCol - 1]) && !crmBool_(row[lockCol - 1]);
      var cols = row._new ? null : Object.keys(row._d).map(Number);
      for (var c = 1; c <= width; c++) {
        if (cols && cols.indexOf(c) === -1) continue;
        if (lockedMeanwhile && protectedCols[c]) continue;
        target[c - 1] = row[c - 1];
      }
    });
    var out = fresh.concat(appended);
    if (out.length) t.sheet.getRange(2, 1, out.length, width).setValues(out);
  }
  changed.forEach(function (r) { delete r._d; delete r._new; });
  t.dirty = false;
}

/**
 * Persist a job's state only if the job is still the current one. Returns false if the job was
 * stopped (property deleted) or replaced (a newer job started) while this tick was running, so
 * the caller must not re-create the trigger or overwrite the newer state.
 */
function crmCommitState_(prop, state) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(prop);
  if (!raw) return false;
  try { if (JSON.parse(raw).startedAt !== state.startedAt) return false; } catch (e) { return false; }
  props.setProperty(prop, JSON.stringify(state));
  return true;
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

var CRM_FREEMAIL = { 'gmail.com': 1, 'googlemail.com': 1, 'yahoo.com': 1, 'ymail.com': 1, 'rocketmail.com': 1, 'outlook.com': 1, 'hotmail.com': 1, 'live.com': 1, 'msn.com': 1, 'icloud.com': 1, 'me.com': 1, 'mac.com': 1, 'aol.com': 1, 'protonmail.com': 1, 'proton.me': 1, 'pm.me': 1, 'hey.com': 1, 'gmx.com': 1, 'gmx.net': 1, 'gmx.de': 1, 'web.de': 1, 'fastmail.com': 1, 'fastmail.fm': 1, 'zoho.com': 1, 'mail.com': 1, 'email.com': 1, 'yandex.com': 1, 'yandex.ru': 1, 'mail.ru': 1, 'qq.com': 1, '163.com': 1, '126.com': 1, 'naver.com': 1, 'daum.net': 1, 'tutanota.com': 1, 'tuta.io': 1, 'duck.com': 1, 'comcast.net': 1, 'verizon.net': 1, 'att.net': 1, 'sbcglobal.net': 1, 'bellsouth.net': 1, 'cox.net': 1, 'btinternet.com': 1, 'sky.com': 1, 'orange.fr': 1, 'wanadoo.fr': 1, 'free.fr': 1, 'laposte.net': 1, 'sfr.fr': 1, 't-online.de': 1, 'libero.it': 1, 'virgilio.it': 1, 'seznam.cz': 1, 'wp.pl': 1, 'o2.pl': 1, 'rediffmail.com': 1 };
var CRM_FREEMAIL_RE = /^(yahoo|ymail|hotmail|outlook|live|msn|googlemail|gmx|icloud|me|aol|protonmail|proton|yandex|mail)\.[a-z.]{2,8}$/;
function crmIsFreemail_(domain) {
  var d = String(domain || '').toLowerCase();
  return !!CRM_FREEMAIL[d] || CRM_FREEMAIL_RE.test(d);
}

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

/** Addresses that are never real contacts. Local-part rules do the work; the domain list is only true bot/list senders. */
var CRM_MACHINE_LOCAL_RE = /(^(mailer-daemon|postmaster|bounces?|notifications?|notify|alerts?|calendar-notification|unsubscribe|newsletter|digest|automated|robot|bot|system|daemon|invitations?|jobs-listings|jobalerts|receipts?|billing|invoices?)(\+.*)?$)|(no-?reply|do-?not-?reply|donotreply|noreply)/i;
var CRM_MACHINE_DOMAIN_RE = /(^|\.)(calendar\.google\.com|docs\.google\.com|drive-shares\.google\.com|googlegroups\.com|facebookmail\.com|intercom-mail\.com|amazonses\.com|sendgrid\.net|mailgun\.org|mandrillapp\.com|sparkpostmail\.com|docusign\.net|calendly\.com|hellosign\.com)$/i;

function crmIsMachineAddress_(email, ignoreDomains) {
  var e = String(email || '').toLowerCase();
  if (!e) return true;
  var local = e.split('@')[0], domain = crmDomainOf_(e);
  if (CRM_MACHINE_LOCAL_RE.test(local)) return true;
  if (CRM_MACHINE_DOMAIN_RE.test(domain)) return true;
  // Reply-token addresses from SaaS tools: reply+1a2b3c4d5e6f@…, msg-9f8e7d6c5b4a@…
  // Requires a long token WITH digits so "r.stevenson@…" style human addresses are untouched.
  var tok = local.match(/^(reply|msg|bounce|b|r|m)[-+._]([a-z0-9]{10,})$/);
  if (tok && /\d/.test(tok[2])) return true;
  for (var i = 0; i < (ignoreDomains || []).length; i++) {
    var d = ignoreDomains[i];
    if (domain === d || domain.slice(-(d.length + 1)) === '.' + d) return true;
  }
  return false;
}

/**
 * Split a plain-text body into { body, signature } with quoted history removed.
 * The signature is the block after a "--" / "—" delimiter, or the trailing run of short
 * lines that look like a sign-off (name, title, company, phone, URL).
 */
function crmSplitBody_(plainBody) {
  var text = String(plainBody || '').replace(/\r/g, '');
  // Cut at the start of quoted history.
  var cutters = [
    /\n\s*On [\s\S]{5,200}?wrote:\s*\n/,     // "On Mon, Jan 1 … <x@y> wrote:" (may wrap)
    /\n\s*-{2,}\s*Original Message\s*-{2,}/i,
    /\n\s*-{2,}\s*Forwarded message\s*-{2,}/i,
    /\n\s*From:\s.+\n\s*(Sent|Date|To):\s.+/i, // Outlook-style header block
    /\n\s*_{5,}\s*\n/,                       // Outlook separator
    /\n\s*Le .{5,200}? a écrit\s*:/,         // French
    /\n\s*Am .{5,200}? schrieb .{0,80}:/,    // German
    /\n\s*El .{5,200}? escribió:/            // Spanish
  ];
  cutters.forEach(function (re) { var m = text.match(re); if (m && m.index > 0) text = text.slice(0, m.index); });
  var lines = text.split('\n').filter(function (l) { return !/^\s*>/.test(l); });
  // Trim trailing blank lines.
  while (lines.length && !lines[lines.length - 1].trim()) lines.pop();

  var sigStart = -1;
  for (var i = 0; i < lines.length; i++) {
    if (/^\s*(--|—|__)\s*$/.test(lines[i]) || /^\s*(Sent from my|Get Outlook for|Sent via)/i.test(lines[i])) { sigStart = i; break; }
  }
  if (sigStart === -1) {
    // Heuristic sign-off: the last paragraph, if it is ≤ 6 short lines and contains a sign-off cue.
    var lastBlank = -1;
    for (var j = lines.length - 1; j >= 0; j--) if (!lines[j].trim()) { lastBlank = j; break; }
    var tail = lines.slice(lastBlank + 1);
    var cue = /(^\s*(best|thanks|thank you|regards|cheers|sincerely|warmly|talk soon|kind regards|br|rgds)[,!.]?\s*$)|(\+?\d[\d\s().-]{7,}\d)|(https?:\/\/|www\.)|(\b(ceo|cto|coo|cfo|founder|co-founder|partner|principal|associate|director|manager|engineer|vp|head of)\b)/i;
    if (tail.length >= 2 && tail.length <= 7 && tail.every(function (l) { return l.length <= 80; }) && tail.some(function (l) { return cue.test(l); })) {
      sigStart = lastBlank + 1;
    }
  }
  var bodyLines = sigStart >= 0 ? lines.slice(0, sigStart) : lines;
  var sigLines = sigStart >= 0 ? lines.slice(sigStart).filter(function (l) { return !/^\s*(--|—|__)\s*$/.test(l); }) : [];
  var clean = function (arr) {
    return arr.join('\n').replace(/\[image:[^\]]*\]/gi, '').replace(/[ \t]+/g, ' ').replace(/\n{2,}/g, '\n').trim();
  };
  return { body: clean(bodyLines), signature: clean(sigLines).slice(0, 300) };
}

/** Clip to maxChars on a word boundary with an ellipsis. maxChars 0 means "no text at all". */
function crmClip_(text, maxChars) {
  if (maxChars === 0) return '';
  var t = String(text || '');
  var max = maxChars || 400;
  return t.length > max ? t.slice(0, max).replace(/\s+\S*$/, '') + '…' : t;
}

/** Plain-text excerpt of the message body (quoted history and signature removed), clipped to maxChars. */
function crmExcerpt_(plainBody, maxChars) {
  if (maxChars === 0) return '';
  return crmClip_(crmSplitBody_(plainBody).body, maxChars);
}

/** Auto-replies (out of office, vacation responders, "Automatic reply") must not count as replies. */
var CRM_AUTOREPLY_SUBJECT_RE = /^\s*((re|fwd?):\s*)*(automatic reply|auto-?reply|autoreply|auto response|automated response|out of (the )?office|ooo\b|away from (the )?office|on vacation|on leave|abwesenheit|réponse automatique|respuesta automática|delivery status notification|undeliverable|undelivered mail|mail delivery (failed|subsystem)|failure notice)/i;
function crmIsAutoReply_(subject, headers) {
  if (CRM_AUTOREPLY_SUBJECT_RE.test(String(subject || ''))) return true;
  var h = headers || {};
  var auto = String(h['Auto-Submitted'] || h['auto-submitted'] || '').toLowerCase();
  if (auto && auto !== 'no') return true;
  if (String(h['X-Autoreply'] || h['X-Autorespond'] || h['X-Auto-Response-Suppress'] || '').trim()) return true;
  var precedence = String(h['Precedence'] || '').toLowerCase();
  if (precedence === 'bulk' || precedence === 'junk' || precedence === 'auto_reply') return true;
  return false;
}

/** Deterministic opt-out detection on inbound text. Conservative on purpose: false positives suppress a real contact. */
var CRM_OPTOUT_RE = /\b(unsubscribe( me)?|remove me from (your|this|the) (list|mailing)|take me off (your|this|the) list|stop (emailing|contacting|messaging) (me|us)|do not (email|contact) (me|us)( again)?|don'?t (email|contact) (me|us)( again)?|no longer interested|not interested,? thanks|please stop)\b/i;
function crmIsOptOut_(text) {
  return CRM_OPTOUT_RE.test(String(text || ''));
}

/** "Dr. Jane Q. van der Doe (she/her)" → "jane doe"; used for possible-same-person matching. */
function crmNormalizeName_(name) {
  var n = String(name || '').toLowerCase();
  try { n = n.normalize('NFD').replace(/[̀-ͯ]/g, ''); } catch (e) { /* older runtime */ }
  n = n.replace(/\([^)]*\)/g, ' ')                                  // (she/her), (Acme)
    .replace(/\b(dr|mr|mrs|ms|prof|sir|jr|sr|ii|iii|phd|mba|md)\b\.?/g, ' ')
    .replace(/[^a-z\s'-]/g, ' ').replace(/\s+/g, ' ').trim();
  var parts = n.split(' ').filter(String);
  if (parts.length >= 2) return parts[0] + ' ' + parts[parts.length - 1];
  return parts.join(' ');
}

/** "Acme, Inc." → "acme"; "The Acme Company LLC" → "acme". */
function crmNormalizeCompany_(company) {
  return String(company || '').toLowerCase()
    .replace(/[^a-z0-9\s&]/g, ' ')
    .replace(/\b(the|inc|incorporated|llc|ltd|limited|co|corp|corporation|company|gmbh|ag|sa|plc|group|holdings|ventures|capital|partners|labs)\b/g, ' ')
    .replace(/\s+/g, ' ').trim();
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
