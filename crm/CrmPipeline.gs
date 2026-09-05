/**
 * Pipeline stages, dormant overlay, and sync with the mail merge Contacts tab.
 */

/** Menu entry: recompute Stage / LastTouchAt for every unlocked Master row. */
function crmRecomputePipeline() {
  var master = crmLoadTable_('Master', CRM_MASTER_COLUMNS);
  var changed = crmRecomputePipeline_(master);
  crmSaveTable_(master);
  crmLog_('PIPELINE', '', changed + ' stage changes');
  var msg = 'Pipeline recomputed: ' + changed + ' stage change(s).';
  crmToast_(msg);
  return msg;
}

/** In-memory recompute over a loaded Master table. Returns number of stage changes. */
function crmRecomputePipeline_(master) {
  var settings = crmSettings_();
  var changed = 0;
  master.rows.forEach(function (row) {
    if (!String(crmGet_(master, row, 'Email')).trim()) return;
    var lastTouch = crmMaxDate_(crmGet_(master, row, 'LastSentAt'), crmGet_(master, row, 'LastReplyAt'));
    if (lastTouch) crmSet_(master, row, 'LastTouchAt', lastTouch);
    if (crmBool_(crmGet_(master, row, 'Lock'))) return;
    var before = String(crmGet_(master, row, 'Stage')).trim();
    var after = crmComputeStage_({
      bounced: crmBool_(crmGet_(master, row, 'Bounced')),
      doNotContact: crmBool_(crmGet_(master, row, 'DoNotContact')),
      relationship: crmGetTag_(crmGet_(master, row, 'Tags'), 'rel:'),
      sent: Number(crmGet_(master, row, 'SentCount')) || 0,
      received: Number(crmGet_(master, row, 'ReceivedCount')) || 0,
      daysSinceTouch: crmDaysAgo_(lastTouch),
      currentStage: before
    }, settings.DormantDays);
    if (after !== before) { crmSet_(master, row, 'Stage', after); changed++; }
  });
  return changed;
}

/**
 * Pure stage function (unit-tested in tests/run.js).
 * f = { bounced, doNotContact, relationship, sent, received, daysSinceTouch, currentStage }
 */
function crmComputeStage_(f, dormantDays) {
  if (f.bounced) return 'Bounced';
  if (f.doNotContact || f.relationship === 'not_interested') return 'Not Interested';
  // Hand-set terminal stages survive recompute even without Lock.
  if (f.currentStage === 'Won' || f.currentStage === 'Not Interested') return f.currentStage;
  var stage;
  if (f.relationship === 'customer') stage = 'Won';
  else if (f.relationship === 'meeting_scheduled' || f.currentStage === 'Meeting') stage = 'Meeting';
  else if (f.received >= 2 || f.relationship === 'engaged') stage = 'Engaged';
  else if (f.received >= 1) stage = 'Replied';
  else if (f.sent >= 1) stage = 'Contacted';
  else stage = 'New';
  if ((stage === 'Replied' || stage === 'Engaged' || stage === 'Meeting') && f.daysSinceTouch > (dormantDays || 45)) return 'Dormant';
  return stage;
}

// ---------------------------------------------------------------- Mail merge sync

/** The mail merge "Contacts" sheet: this spreadsheet's tab, or the one at Settings → MailMergeSheetUrl. */
function crmMailMergeContactsSheet_(createIfMissing) {
  var settings = crmSettings_();
  var ss = SpreadsheetApp.getActive();
  if (settings.MailMergeSheetUrl) {
    var idMatch = settings.MailMergeSheetUrl.match(/[-\w]{25,}/);
    ss = idMatch ? SpreadsheetApp.openById(idMatch[0]) : SpreadsheetApp.openByUrl(settings.MailMergeSheetUrl);
  }
  var sheet = ss.getSheetByName('Contacts');
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet('Contacts');
    var headers = ['Email', 'FirstName', 'Company', 'Attachment', 'Status', 'SentAt', 'Opens', 'LastOpenAt', 'Clicks', 'Replied', 'Bounced', 'FollowUpStage', 'ThreadId', 'TrackingId'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  if (!sheet) throw new Error('No "Contacts" tab found. Set up the mail merge in this spreadsheet or point Settings → MailMergeSheetUrl at it.');
  return sheet;
}

/**
 * Pull mail merge outcomes into Master: creates missing contacts (Source = MailMerge),
 * copies Bounced, and treats Replied as at least one inbound message.
 */
function crmSyncFromMailMerge() {
  var contacts = crmMailMergeContactsSheet_(false);
  var cmap = crmColMap_(contacts, ['Email']);
  var lastRow = contacts.getLastRow();
  if (lastRow < 2) return 'Contacts tab is empty — nothing to sync.';
  var data = contacts.getRange(2, 1, lastRow - 1, contacts.getLastColumn()).getValues();
  var headers = contacts.getRange(1, 1, 1, contacts.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim().toLowerCase(); });

  var master = crmLoadTable_('Master', CRM_MASTER_COLUMNS, function (r, t) { return crmNormalizeEmail_(crmGet_(t, r, 'Email')); });
  var added = 0, updated = 0;
  data.forEach(function (r) {
    var email = crmNormalizeEmail_(r[cmap['email'] - 1]);
    if (!email) return;
    var val = function (h) { var i = headers.indexOf(h); return i === -1 ? '' : r[i]; };
    var isNew = master.index[email] === undefined;
    var row = isNew ? crmAppendRow_(master, email) : master.rows[master.index[email]];
    if (isNew) {
      crmSet_(master, row, 'Email', email);
      crmSet_(master, row, 'Domain', crmDomainOf_(email));
      crmSet_(master, row, 'Source', 'MailMerge');
      crmSet_(master, row, 'SentCount', 0);
      crmSet_(master, row, 'ReceivedCount', 0);
      crmSet_(master, row, 'ThreadCount', 0);
      var first = String(val('firstname')).trim(), name = String(val('name')).trim();
      if (name) crmSet_(master, row, 'Name', name);
      if (first) crmSet_(master, row, 'FirstName', first);
      if (!first && name) crmSet_(master, row, 'FirstName', crmSplitName_(name).first);
      var company = String(val('company')).trim();
      if (company) crmSet_(master, row, 'Company', company);
      added++;
    }
    var status = String(val('status')).trim();
    var sentAt = val('sentat');
    if (status === 'SENT' && sentAt instanceof Date) {
      if (!(crmGet_(master, row, 'FirstSentAt') instanceof Date)) crmSet_(master, row, 'FirstSentAt', sentAt);
      if (!(crmGet_(master, row, 'LastSentAt') instanceof Date) || sentAt > crmGet_(master, row, 'LastSentAt')) crmSet_(master, row, 'LastSentAt', sentAt);
      if ((Number(crmGet_(master, row, 'SentCount')) || 0) === 0) crmSet_(master, row, 'SentCount', 1);
    }
    if (crmBool_(val('bounced'))) crmSet_(master, row, 'Bounced', true);
    if (crmBool_(val('replied')) && (Number(crmGet_(master, row, 'ReceivedCount')) || 0) === 0) crmSet_(master, row, 'ReceivedCount', 1);
    var threadId = String(val('threadid')).trim();
    if (threadId) crmSet_(master, row, 'ThreadIds', crmAddToList_(crmGet_(master, row, 'ThreadIds'), threadId, 300));
    crmSet_(master, row, 'UpdatedAt', new Date());
    if (!isNew) updated++;
  });
  crmRecomputePipeline_(master);
  crmSaveTable_(master);
  crmLog_('SYNC_MAILMERGE', '', added + ' added, ' + updated + ' updated');
  var msg = 'Synced from mail merge: ' + added + ' new contact(s), ' + updated + ' updated.';
  crmToast_(msg);
  return msg;
}

// ---------------------------------------------------------------- Audit

/** Menu/sidebar entry: integrity report for Master. Returns the report text; also logged. */
function crmAudit() {
  var master = crmLoadTable_('Master', CRM_MASTER_COLUMNS);
  var settings = crmSettings_();
  var report = crmAuditMaster_(master, settings);
  crmLog_('AUDIT', '', report.summary);
  crmToast_(report.summary);
  return report;
}

/**
 * Pure audit over a loaded Master table. Returns { summary, findings: [{ check, count, examples }] }.
 * Each check is an invariant the pipeline is supposed to maintain.
 */
function crmAuditMaster_(master, settings) {
  var findings = [];
  var add = function (check, rows, note) {
    if (!rows.length) return;
    findings.push({ check: check, count: rows.length, note: note || '', examples: rows.slice(0, 5).map(function (i) { return String(crmGet_(master, master.rows[i], 'Email')); }) });
  };
  var idx = function (pred) { var out = []; master.rows.forEach(function (r, i) { if (String(crmGet_(master, r, 'Email')).trim() && pred(r)) out.push(i); }); return out; };
  var g = function (r, h) { return crmGet_(master, r, h); };
  var n = function (r, h) { return Number(g(r, h)) || 0; };

  // Duplicate normalized emails (should be impossible after a scan; happens after hand edits).
  var seen = {}, dups = [];
  master.rows.forEach(function (r, i) {
    var e = crmNormalizeEmail_(g(r, 'Email'));
    if (!e) return;
    if (seen[e] !== undefined) dups.push(i); else seen[e] = i;
  });
  add('Duplicate emails after normalization', dups, 'Merge the rows by hand; the scan updates the first one only.');

  add('Email cell not a valid address', idx(function (r) { return !crmNormalizeEmail_(g(r, 'Email')); }));
  add('Unclassified (no Category)', idx(function (r) { return !String(g(r, 'Category')).trim(); }), 'Run Classify.');
  add('Flagged for review (ReviewNeeded)', idx(function (r) { return crmBool_(g(r, 'ReviewNeeded')); }), 'Low confidence or unclear — check Evidence, fix by hand, tick Lock.');
  add('Confidence below MinConfidence but not flagged', idx(function (r) {
    var c = g(r, 'Confidence'); return String(g(r, 'Category')).trim() && c !== '' && Number(c) < settings.MinConfidence && !crmBool_(g(r, 'ReviewNeeded')) && String(g(r, 'ClassifiedBy')) === 'AI';
  }));
  add('AI-classified without Evidence', idx(function (r) { return String(g(r, 'ClassifiedBy')) === 'AI' && !String(g(r, 'Evidence')).trim(); }));
  add('Stage disagrees with facts', idx(function (r) {
    if (crmBool_(g(r, 'Lock'))) return false;
    var expected = crmComputeStage_({
      bounced: crmBool_(g(r, 'Bounced')), doNotContact: crmBool_(g(r, 'DoNotContact')), relationship: crmGetTag_(g(r, 'Tags'), 'rel:'),
      sent: n(r, 'SentCount'), received: n(r, 'ReceivedCount'), daysSinceTouch: crmDaysAgo_(g(r, 'LastTouchAt')), currentStage: String(g(r, 'Stage')).trim()
    }, settings.DormantDays);
    return expected !== String(g(r, 'Stage')).trim();
  }), 'Run Recompute pipeline stages.');
  add('Replied but no inbound excerpt', idx(function (r) { return n(r, 'ReceivedCount') > 0 && !String(g(r, 'LastInboundSnippet')).trim() && String(g(r, 'Source')) === 'Scan' && settings.SnippetChars > 0; }), 'Usually an inbound message with an empty plain-text body.');
  add('Counts without thread ids', idx(function (r) { return (n(r, 'SentCount') + n(r, 'ReceivedCount')) > 0 && !String(g(r, 'ThreadIds')).trim() && String(g(r, 'Source')) === 'Scan'; }));
  add('NeedsReply on a suppressed contact', idx(function (r) { return crmBool_(g(r, 'NeedsReply')) && (crmBool_(g(r, 'DoNotContact')) || crmBool_(g(r, 'Bounced'))); }));
  add('Opted out but still contactable', idx(function (r) { return crmSplitList_(g(r, 'Tags')).indexOf('optout') !== -1 && !crmBool_(g(r, 'DoNotContact')); }), 'Someone cleared DoNotContact after an opt-out was detected.');
  add('Category "Ignore" still in Master', idx(function (r) { return String(g(r, 'Category')) === 'Ignore'; }), 'Add a Rules row (IGNORE) for the domain and delete these rows.');
  add('Locked rows', idx(function (r) { return crmBool_(g(r, 'Lock')); }), 'Informational.');

  var total = idx(function () { return true; }).length;
  var problems = findings.filter(function (f) { return f.check !== 'Locked rows' && f.check !== 'Flagged for review (ReviewNeeded)'; })
    .reduce(function (a, f) { return a + f.count; }, 0);
  var summary = 'Audit: ' + total + ' contacts, ' + problems + ' integrity issue(s)' + (findings.length ? ' — ' + findings.map(function (f) { return f.check + ': ' + f.count; }).join('; ') : '') + '.';
  return { summary: summary, findings: findings, total: total, problems: problems };
}
