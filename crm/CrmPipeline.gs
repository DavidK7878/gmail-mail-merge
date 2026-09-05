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
