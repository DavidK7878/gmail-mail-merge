/**
 * Import a new sourcing list, dedupe it against Master, and push the genuinely new
 * rows into the mail merge Contacts tab.
 */

var CRM_IMPORT_HEADER_ALIASES = {
  email: ['email', 'e-mail', 'email address', 'e-mail address', 'work email', 'mail'],
  name: ['name', 'full name', 'contact', 'contact name', 'person'],
  firstname: ['firstname', 'first name', 'first', 'given name'],
  lastname: ['lastname', 'last name', 'last', 'surname', 'family name'],
  company: ['company', 'organization', 'organisation', 'org', 'firm', 'fund', 'employer', 'account'],
  title: ['title', 'role', 'position', 'job title'],
  notes: ['notes', 'note', 'comment', 'comments', 'context']
};

// ---------------------------------------------------------------- Import from Drive

function crmImportFromDrivePrompt() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Import from Drive', 'Paste a Google Sheets URL, a Drive file ID, or the exact name of a CSV/Sheet in your Drive.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var msg = crmImportFromDrive(res.getResponseText());
  ui.alert(msg);
}

/** Append rows from a Drive spreadsheet (first tab) or CSV into the Import tab. */
function crmImportFromDrive(ref) {
  ref = String(ref || '').trim();
  if (!ref) throw new Error('No file reference given.');
  var table = crmReadDriveTable_(ref);
  if (!table.length) throw new Error('The file is empty.');
  var added = crmAppendToImport_(table[0], table.slice(1));
  crmLog_('IMPORT', ref, added + ' rows appended to Import');
  return added + ' row(s) added to the Import tab. Now run CRM → Check Import against Master.';
}

function crmReadDriveTable_(ref) {
  var file = null;
  var idMatch = ref.match(/[-\w]{25,}/);
  if (idMatch) {
    try { file = DriveApp.getFileById(idMatch[0]); } catch (e) { file = null; }
  }
  if (!file) {
    var files = DriveApp.getFilesByName(ref);
    if (files.hasNext()) file = files.next();
  }
  if (!file) throw new Error('File not found in Drive: "' + ref + '"');
  var mime = file.getMimeType();
  if (mime === MimeType.GOOGLE_SHEETS) {
    var sheet = SpreadsheetApp.openById(file.getId()).getSheets()[0];
    var lr = sheet.getLastRow(), lc = sheet.getLastColumn();
    return lr && lc ? sheet.getRange(1, 1, lr, lc).getValues() : [];
  }
  if (/csv|text\/plain/i.test(mime) || /\.csv$/i.test(file.getName())) {
    return Utilities.parseCsv(file.getBlob().getDataAsString());
  }
  if (/spreadsheetml|ms-excel/i.test(mime)) {
    throw new Error('Excel files must be converted to Google Sheets first (right-click in Drive → Open with → Google Sheets).');
  }
  throw new Error('Unsupported file type: ' + mime);
}

/** Map source headers → Import headers (aliases), adding unknown headers as new Import columns. */
function crmAppendToImport_(srcHeaders, srcRows) {
  var imp = crmSheet_('Import');
  var map = crmColMap_(imp, ['Email']);
  var headers = imp.getRange(1, 1, 1, imp.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });

  var plan = srcHeaders.map(function (h) {
    var key = String(h).trim().toLowerCase();
    if (!key) return null;
    var canonical = crmCanonicalImportHeader_(key);
    var target = canonical || String(h).trim();
    var idx = headers.map(function (x) { return x.toLowerCase(); }).indexOf(target.toLowerCase());
    if (idx === -1) { headers.push(target); idx = headers.length - 1; }
    return idx;
  });
  if (headers.length > imp.getLastColumn()) {
    imp.getRange(1, 1, 1, headers.length).setValues([headers]);
    imp.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  var out = [];
  srcRows.forEach(function (r) {
    if (!r.some(function (v) { return String(v).trim(); })) return;
    var row = headers.map(function () { return ''; });
    plan.forEach(function (idx, i) { if (idx !== null && idx !== undefined && i < r.length) row[idx] = r[i]; });
    if (!crmNormalizeEmail_(row[map['email'] - 1])) {
      // Try to find an email anywhere in the row if the header mapping missed it.
      var found = r.map(crmNormalizeEmail_).filter(String)[0];
      if (!found) return;
      row[map['email'] - 1] = found;
    }
    out.push(row);
  });
  if (out.length) imp.getRange(imp.getLastRow() + 1, 1, out.length, headers.length).setValues(out);
  return out.length;
}

function crmCanonicalImportHeader_(key) {
  var pretty = { email: 'Email', name: 'Name', firstname: 'FirstName', lastname: 'LastName', company: 'Company', title: 'Title', notes: 'Notes' };
  for (var k in CRM_IMPORT_HEADER_ALIASES) {
    if (CRM_IMPORT_HEADER_ALIASES[k].indexOf(key) !== -1) return pretty[k];
  }
  return null;
}

// ---------------------------------------------------------------- Dedupe

/** Fill DedupeStatus / Matched* / Recommendation for every Import row. */
function crmCheckImport() {
  var settings = crmSettings_();
  var imp = crmLoadTable_('Import', ['Email'].concat(CRM_IMPORT_MACHINE_COLUMNS));
  if (!imp.rows.length) return 'Import tab is empty.';
  var master = crmLoadTable_('Master', CRM_MASTER_COLUMNS, function (r, t) { return crmNormalizeEmail_(crmGet_(t, r, 'Email')); });

  // Domain index: best (most advanced) contact per non-freemail domain.
  var byDomain = {};
  master.rows.forEach(function (row) {
    var email = crmNormalizeEmail_(crmGet_(master, row, 'Email'));
    if (!email) return;
    var d = crmDomainOf_(email);
    if (!d || crmIsFreemail_(d)) return;
    var cur = byDomain[d];
    if (!cur || crmStageRank_(crmGet_(master, row, 'Stage')) > crmStageRank_(crmGet_(master, cur, 'Stage'))) byDomain[d] = row;
  });

  var seen = {};
  var counts = {};
  imp.rows.forEach(function (row) {
    var email = crmNormalizeEmail_(crmGet_(imp, row, 'Email'));
    var res = crmDedupeRow_(email, seen, master, byDomain, settings);
    seen[email] = true;
    crmSet_(imp, row, 'DedupeStatus', res.status);
    crmSet_(imp, row, 'MatchedName', res.matchedName);
    crmSet_(imp, row, 'MatchedCategory', res.matchedCategory);
    crmSet_(imp, row, 'MatchedStage', res.matchedStage);
    crmSet_(imp, row, 'LastTouch', res.lastTouch);
    crmSet_(imp, row, 'Recommendation', res.recommendation);
    counts[res.status] = (counts[res.status] || 0) + 1;
  });
  crmSaveTable_(imp);
  crmColorImport_(imp);
  var summary = Object.keys(counts).map(function (k) { return k + ': ' + counts[k]; }).join(', ');
  crmLog_('IMPORT_CHECK', '', summary);
  var msg = 'Checked ' + imp.rows.length + ' row(s) — ' + summary + '.';
  crmToast_(msg);
  return msg;
}

/**
 * Pure-ish dedupe decision for one import row (unit-tested).
 * master: loaded Master table; byDomain: domain → best master row.
 */
function crmDedupeRow_(email, seenInImport, master, byDomain, settings) {
  var out = { status: 'NEW', matchedName: '', matchedCategory: '', matchedStage: '', lastTouch: '', recommendation: 'Send' };
  if (!email) { out.status = 'INVALID'; out.recommendation = 'Fix email'; return out; }
  if (seenInImport[email]) { out.status = 'DUPLICATE_IN_IMPORT'; out.recommendation = 'Skip — repeated in this list'; return out; }

  var fill = function (row) {
    out.matchedName = String(crmGet_(master, row, 'Name') || crmGet_(master, row, 'Email'));
    out.matchedCategory = String(crmGet_(master, row, 'Category') || '');
    out.matchedStage = String(crmGet_(master, row, 'Stage') || '');
    var lt = crmGet_(master, row, 'LastTouchAt');
    out.lastTouch = lt instanceof Date ? lt : '';
  };

  if (master.index[email] !== undefined) {
    var row = master.rows[master.index[email]];
    out.status = 'DUPLICATE';
    fill(row);
    var stage = out.matchedStage;
    var days = crmDaysAgo_(crmGet_(master, row, 'LastTouchAt'));
    if (crmBool_(crmGet_(master, row, 'DoNotContact')) || crmBool_(crmGet_(master, row, 'Bounced')) || stage === 'Not Interested' || stage === 'Bounced') out.recommendation = 'Skip — do not contact';
    else if (stage === 'Replied' || stage === 'Engaged' || stage === 'Meeting' || stage === 'Won') out.recommendation = 'Skip — already in conversation';
    else if (crmBool_(crmGet_(master, row, 'NeedsReply'))) out.recommendation = 'Skip — they are waiting on your reply';
    else if (days <= (settings.ImportRecentDays || 30)) out.recommendation = 'Skip — emailed ' + Math.round(days) + ' day(s) ago';
    else if (stage === 'New') out.recommendation = 'Send — already queued, never sent';
    else out.recommendation = 'Re-engage — reference prior thread';
    return out;
  }

  var domain = crmDomainOf_(email);
  if (domain && !crmIsFreemail_(domain) && byDomain[domain]) {
    var best = byDomain[domain];
    out.status = 'SAME_COMPANY';
    fill(best);
    var bs = out.matchedStage;
    if (bs === 'Won') out.recommendation = 'Review — company is already a customer (' + out.matchedName + ')';
    else if (bs === 'Not Interested' || crmBool_(crmGet_(master, best, 'DoNotContact'))) out.recommendation = 'Review — ' + out.matchedName + ' at this company declined';
    else if (bs === 'Replied' || bs === 'Engaged' || bs === 'Meeting') out.recommendation = 'Review — in conversation with ' + out.matchedName;
    else out.recommendation = 'Send — colleague ' + out.matchedName + ' was contacted, no reply';
    return out;
  }
  return out;
}

function crmStageRank_(stage) {
  var order = ['', 'Bounced', 'Not Interested', 'New', 'Contacted', 'Dormant', 'Replied', 'Engaged', 'Meeting', 'Won'];
  var i = order.indexOf(String(stage || '').trim());
  return i === -1 ? 0 : i;
}

function crmColorImport_(imp) {
  try {
    var col = imp.map['dedupestatus'];
    var n = imp.rows.length;
    var range = imp.sheet.getRange(2, col, n, 1);
    var colors = imp.rows.map(function (row) {
      var s = String(crmGet_(imp, row, 'DedupeStatus'));
      return [s === 'NEW' ? '#e6f4ea' : s === 'SAME_COMPANY' ? '#fef7e0' : s === 'INVALID' ? '#f1f3f4' : '#fce8e6'];
    });
    range.setBackgrounds(colors);
  } catch (e) { /* cosmetic */ }
}

// ---------------------------------------------------------------- Push to mail merge

/** Menu entry. Pushes NEW rows. */
function crmPushNewToMailMerge() { var m = crmPushToMailMerge(['NEW']); crmToast_(m); return m; }

/**
 * Append Import rows with the given DedupeStatus values to the mail merge Contacts tab and
 * register them in Master (Stage New, Source Import). Rows already pushed are skipped.
 */
function crmPushToMailMerge(statuses) {
  statuses = statuses && statuses.length ? statuses : ['NEW'];
  var imp = crmLoadTable_('Import', ['Email'].concat(CRM_IMPORT_MACHINE_COLUMNS));
  var todo = imp.rows.filter(function (row) {
    return statuses.indexOf(String(crmGet_(imp, row, 'DedupeStatus'))) !== -1 && !crmGet_(imp, row, 'PushedAt');
  });
  if (!todo.length) return 'Nothing to push. Run "Check Import against Master" first, or all matching rows were already pushed.';

  var contacts = crmMailMergeContactsSheet_(true);
  var cHeaders = contacts.getRange(1, 1, 1, contacts.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  var cLower = cHeaders.map(function (h) { return h.toLowerCase(); });
  if (cLower.indexOf('email') === -1) throw new Error('The Contacts tab has no "Email" column.');

  // Existing Contacts emails → don't double-queue.
  var existing = {};
  if (contacts.getLastRow() >= 2) {
    contacts.getRange(2, cLower.indexOf('email') + 1, contacts.getLastRow() - 1, 1).getValues().forEach(function (r) {
      var e = crmNormalizeEmail_(r[0]); if (e) existing[e] = true;
    });
  }

  var master = crmLoadTable_('Master', CRM_MASTER_COLUMNS, function (r, t) { return crmNormalizeEmail_(crmGet_(t, r, 'Email')); });
  var outRows = [];
  var now = new Date();
  todo.forEach(function (row) {
    var email = crmNormalizeEmail_(crmGet_(imp, row, 'Email'));
    if (!email) return;
    var name = String(crmGet_(imp, row, 'Name')).trim();
    var first = String(crmGet_(imp, row, 'FirstName')).trim() || crmSplitName_(name).first;
    var last = String(crmGet_(imp, row, 'LastName')).trim() || crmSplitName_(name).last;
    var company = String(crmGet_(imp, row, 'Company')).trim();

    if (!existing[email]) {
      var line = cHeaders.map(function (h) {
        var k = h.toLowerCase();
        if (k === 'email') return String(crmGet_(imp, row, 'Email')).trim().toLowerCase();
        if (k === 'firstname') return first;
        if (k === 'lastname') return last;
        if (k === 'name') return name || (first + ' ' + last).trim();
        if (k === 'company') return company;
        // Any other Import column whose header matches a Contacts column passes through (custom merge fields).
        return imp.map[k] ? crmGet_(imp, row, h) : '';
      });
      outRows.push(line);
      existing[email] = true;
    }

    // Register in Master so the very next import dedupes against it.
    var isNew = master.index[email] === undefined;
    var mrow = isNew ? crmAppendRow_(master, email) : master.rows[master.index[email]];
    if (isNew) {
      crmSet_(master, mrow, 'Email', email);
      crmSet_(master, mrow, 'Domain', crmDomainOf_(email));
      crmSet_(master, mrow, 'Source', 'Import');
      crmSet_(master, mrow, 'Stage', 'New');
      crmSet_(master, mrow, 'SentCount', 0); crmSet_(master, mrow, 'ReceivedCount', 0); crmSet_(master, mrow, 'ThreadCount', 0);
      if (name || first) crmSet_(master, mrow, 'Name', name || (first + ' ' + last).trim());
      if (first) crmSet_(master, mrow, 'FirstName', first);
      if (last) crmSet_(master, mrow, 'LastName', last);
      if (company) crmSet_(master, mrow, 'Company', company);
      var title = String(crmGet_(imp, row, 'Title')).trim(); if (title) crmSet_(master, mrow, 'Title', title);
      var notes = String(crmGet_(imp, row, 'Notes')).trim(); if (notes) crmSet_(master, mrow, 'Notes', notes);
      crmSet_(master, mrow, 'UpdatedAt', now);
    }
    crmSet_(imp, row, 'PushedAt', now);
  });

  if (outRows.length) contacts.getRange(contacts.getLastRow() + 1, 1, outRows.length, cHeaders.length).setValues(outRows);
  crmSaveTable_(master);
  crmSaveTable_(imp);
  crmLog_('IMPORT_PUSH', statuses.join('/'), outRows.length + ' rows → Contacts, ' + todo.length + ' marked pushed');
  return outRows.length + ' contact(s) added to the mail merge Contacts tab' + (todo.length > outRows.length ? ' (' + (todo.length - outRows.length) + ' were already queued there)' : '') + '.';
}

function crmClearImport() {
  var imp = crmSheet_('Import');
  if (imp.getLastRow() >= 2) imp.getRange(2, 1, imp.getLastRow() - 1, imp.getLastColumn()).clear();
  crmToast_('Import tab cleared.');
  return 'Import tab cleared.';
}
