/**
 * Gmail sent-mail scan → Master + Interactions.
 *
 * Resumable: state lives in Script Properties (CRM_SCAN_STATE) and a 1-minute trigger
 * calls crmScanStep until the search is exhausted. Each step stops at ~4.5 minutes.
 *
 * Incremental: the query is bounded (after:<since> before:<scan start day + 1>) so paging
 * with an offset is stable, and the next incremental scan starts 2 days before the last
 * completed scan's start.
 */

var CRM_SCAN_STATE_PROP = 'CRM_SCAN_STATE';
var CRM_LAST_SCAN_PROP = 'CRM_LAST_SCAN_AT';
var CRM_SEARCH_PAGE = 50; // threads per GmailApp.search page

function crmScanIncremental() { var m = crmStartScan(false); crmToast_(m); return m; }
function crmScanFull() { var m = crmStartScan(true); crmToast_(m); return m; }

/** Menu/sidebar entry point. Returns a status string. */
function crmStartScan(fullRescan) {
  var settings = crmSettings_();
  var props = PropertiesService.getScriptProperties();
  var tz = Session.getScriptTimeZone();

  var since = settings.ScanSince;
  if (!fullRescan) {
    var last = props.getProperty(CRM_LAST_SCAN_PROP);
    if (last) { since = new Date(Number(last)); since.setDate(since.getDate() - 2); }
  }
  var startedAt = new Date();
  var before = new Date(startedAt.getTime() + 86400000);
  var query = 'in:sent after:' + Utilities.formatDate(since, tz, 'yyyy/MM/dd') + ' before:' + Utilities.formatDate(before, tz, 'yyyy/MM/dd');

  var state = { query: query, offset: 0, startedAt: startedAt.getTime(), threads: 0, contacts: 0, full: !!fullRescan, done: false };
  props.setProperty(CRM_SCAN_STATE_PROP, JSON.stringify(state));
  crmEnsureMinuteTrigger_('crmScanStep');
  crmLog_('SCAN_START', query, fullRescan ? 'Full rescan' : 'Incremental scan');
  crmScanStep();
  return 'Scan started (' + query + '). Progress shows in the sidebar; you can close the sheet.';
}

function crmStopScan() {
  crmDeleteTriggersFor_('crmScanStep');
  PropertiesService.getScriptProperties().deleteProperty(CRM_SCAN_STATE_PROP);
  crmLog_('SCAN_STOP', '', 'Scan stopped by user');
  crmToast_('Scan stopped.');
  return 'Scan stopped.';
}

function crmScanStatus_() {
  var raw = PropertiesService.getScriptProperties().getProperty(CRM_SCAN_STATE_PROP);
  var last = PropertiesService.getScriptProperties().getProperty(CRM_LAST_SCAN_PROP);
  return { running: !!raw, state: raw ? JSON.parse(raw) : null, lastScanAt: last ? Number(last) : null };
}

/** Trigger handler. Processes up to MaxThreadsPerRun threads within the time budget. */
function crmScanStep() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(CRM_SCAN_STATE_PROP);
    if (!raw) { crmDeleteTriggersFor_('crmScanStep'); return; }
    var state = JSON.parse(raw);
    var settings = crmSettings_();
    var started = Date.now();

    var ctx = crmScanContext_(settings);
    var master = crmLoadTable_('Master', CRM_MASTER_COLUMNS, function (r, t) { return crmNormalizeEmail_(crmGet_(t, r, 'Email')); });
    var inter = crmLoadTable_('Interactions', CRM_INTERACTION_COLUMNS, function (r, t) { return String(crmGet_(t, r, 'ThreadId')).trim(); });

    var processed = 0;
    while (processed < settings.MaxThreadsPerRun && Date.now() - started < CRM_TIME_BUDGET_MS) {
      var threads;
      try {
        threads = GmailApp.search(state.query, state.offset, CRM_SEARCH_PAGE);
      } catch (e) {
        crmLog_('SCAN_ERROR', state.query, 'search failed at offset ' + state.offset + ': ' + e);
        break;
      }
      if (!threads.length) { state.done = true; break; }
      for (var i = 0; i < threads.length; i++) {
        try {
          state.contacts += crmProcessThread_(threads[i], master, inter, ctx);
        } catch (e) {
          crmLog_('SCAN_ERROR', threads[i].getId(), String(e));
        }
        state.offset++; state.threads++; processed++;
        if (Date.now() - started > CRM_TIME_BUDGET_MS) break;
      }
      if (threads.length < CRM_SEARCH_PAGE) { state.done = true; break; }
    }

    crmSaveTable_(inter);
    crmSaveTable_(master);

    if (state.done) {
      props.deleteProperty(CRM_SCAN_STATE_PROP);
      props.setProperty(CRM_LAST_SCAN_PROP, String(state.startedAt));
      crmDeleteTriggersFor_('crmScanStep');
      crmLog_('SCAN_DONE', '', state.threads + ' threads, ' + state.contacts + ' contact updates');
      crmRecomputePipeline_(master);
      crmSaveTable_(master);
      crmApplyValidation_(SpreadsheetApp.getActive());
      // Queue classification for the next trigger tick — never run it inside this execution's time window.
      if (settings.AutoClassifyAfterScan) crmStartClassify_('unclassified', true);
    } else {
      props.setProperty(CRM_SCAN_STATE_PROP, JSON.stringify(state));
      crmEnsureMinuteTrigger_('crmScanStep');
    }
  } finally {
    lock.releaseLock();
  }
}

/** Precompute "who is us" and ignore rules for the scan. */
function crmScanContext_(settings) {
  var me = crmMyEmail_();
  var mine = {};
  mine[crmNormalizeEmail_(me)] = true;
  var myDomains = [];
  settings.MyDomains.forEach(function (d) {
    if (d.indexOf('@') !== -1) mine[crmNormalizeEmail_(d)] = true; else myDomains.push(d);
  });
  // Treat your own (non-freemail) domain as "us" so teammates on Cc aren't contacts.
  var myDomain = crmDomainOf_(me);
  if (myDomain && !crmIsFreemail_(myDomain) && myDomains.indexOf(myDomain) === -1) myDomains.push(myDomain);

  var ignoreRules = crmLoadRules_().filter(function (r) { return r.category.toUpperCase() === 'IGNORE'; });

  return {
    isMine: function (email) {
      if (mine[email]) return true;
      var d = crmDomainOf_(email);
      return myDomains.some(function (md) { return d === md || d.slice(-(md.length + 1)) === '.' + md; });
    },
    isIgnored: function (email) {
      if (crmIsMachineAddress_(email, settings.IgnoreDomains)) return true;
      for (var i = 0; i < ignoreRules.length; i++) if (crmRuleMatches_(ignoreRules[i], { email: email, domain: crmDomainOf_(email), text: '' })) return true;
      return false;
    },
    snippetChars: settings.SnippetChars
  };
}

/**
 * Fold one thread into Master + Interactions. Returns number of contact rows touched.
 */
function crmProcessThread_(thread, master, inter, ctx) {
  var messages = thread.getMessages();
  if (!messages.length) return 0;
  var threadId = thread.getId();
  var subject = thread.getFirstMessageSubject() || '(no subject)';

  var touched = {};   // email → per-thread aggregate
  var outbound = 0, inbound = 0, hasOutbound = false;
  var lastMsg = messages[messages.length - 1];
  var lastFrom = crmParseAddresses_(lastMsg.getFrom())[0];
  var lastIsMine = lastFrom ? ctx.isMine(lastFrom.email) : false;

  for (var m = 0; m < messages.length; m++) {
    var msg = messages[m];
    var date = msg.getDate();
    var from = crmParseAddresses_(msg.getFrom())[0];
    if (!from) continue;
    var fromMine = ctx.isMine(from.email);

    if (fromMine) {
      outbound++; hasOutbound = true;
      var rcpts = crmParseAddresses_(msg.getTo()).concat(crmParseAddresses_(msg.getCc()));
      var excerpt = null;
      rcpts.forEach(function (r) {
        if (ctx.isMine(r.email) || ctx.isIgnored(r.email)) return;
        var a = touched[r.email] || (touched[r.email] = crmEmptyAgg_(r.email));
        if (r.name && !a.name) a.name = r.name;
        a.sent++;
        a.firstSent = crmMinDate_(a.firstSent, date);
        if (!a.lastSent || date > a.lastSent) {
          a.lastSent = date;
          if (excerpt === null) excerpt = crmSafeExcerpt_(msg, ctx.snippetChars);
          a.lastOut = excerpt;
        }
      });
    } else {
      inbound++;
      if (ctx.isIgnored(from.email)) {
        if (/mailer-daemon|postmaster/.test(from.email)) crmMarkBounceFromDaemon_(msg, touched);
        continue;
      }
      var a2 = touched[from.email] || (touched[from.email] = crmEmptyAgg_(from.email));
      if (from.name && !a2.name) a2.name = from.name;
      a2.received++;
      if (!a2.lastReply || date > a2.lastReply) {
        a2.lastReply = date;
        a2.lastIn = crmSafeExcerpt_(msg, ctx.snippetChars);
      }
    }
  }

  // Sent-only scan: a thread with no outbound message from us is noise (shouldn't happen for in:sent).
  if (!hasOutbound) return 0;

  // Interactions row (one per thread; the Email column holds the primary counterpart).
  var emails = Object.keys(touched);
  if (!emails.length) return 0;
  var primary = emails.sort(function (x, y) { return (touched[y].sent + touched[y].received) - (touched[x].sent + touched[x].received); })[0];
  var irow = inter.index[threadId] !== undefined ? inter.rows[inter.index[threadId]] : crmAppendRow_(inter, threadId);
  crmSet_(inter, irow, 'ThreadId', threadId);
  crmSet_(inter, irow, 'Email', primary);
  crmSet_(inter, irow, 'Subject', subject);
  crmSet_(inter, irow, 'FirstMessageAt', messages[0].getDate());
  crmSet_(inter, irow, 'LastMessageAt', lastMsg.getDate());
  crmSet_(inter, irow, 'Messages', messages.length);
  crmSet_(inter, irow, 'Outbound', outbound);
  crmSet_(inter, irow, 'Inbound', inbound);
  crmSet_(inter, irow, 'LastDirection', lastIsMine ? 'OUT' : 'IN');
  crmSet_(inter, irow, 'LastFrom', lastFrom ? lastFrom.email : '');
  crmSet_(inter, irow, 'LastSnippet', crmSafeExcerpt_(lastMsg, ctx.snippetChars));
  crmSet_(inter, irow, 'Link', 'https://mail.google.com/mail/u/0/#all/' + threadId);
  var participants = {};
  emails.forEach(function (e) { participants[e] = [touched[e].sent, touched[e].received]; });
  crmSet_(inter, irow, 'Participants', JSON.stringify(participants));

  // Master rows
  var lastAt = lastMsg.getDate();
  var count = 0;
  emails.forEach(function (email) {
    var a = touched[email];
    var isNew = master.index[email] === undefined;
    var row = isNew ? crmAppendRow_(master, email) : master.rows[master.index[email]];
    if (isNew) {
      crmSet_(master, row, 'Email', email);
      crmSet_(master, row, 'Domain', crmDomainOf_(email));
      crmSet_(master, row, 'Source', 'Scan');
      crmSet_(master, row, 'SentCount', 0);
      crmSet_(master, row, 'ReceivedCount', 0);
      crmSet_(master, row, 'ThreadCount', 0);
    }
    if (a.name && !String(crmGet_(master, row, 'Name')).trim()) {
      crmSet_(master, row, 'Name', a.name);
      var sp = crmSplitName_(a.name);
      if (!String(crmGet_(master, row, 'FirstName')).trim()) crmSet_(master, row, 'FirstName', sp.first);
      if (!String(crmGet_(master, row, 'LastName')).trim()) crmSet_(master, row, 'LastName', sp.last);
    }

    // Thread membership: counts are rebuilt from Interactions.Participants below, so a
    // re-scan of the same thread (the 2-day overlap window) never double counts.
    var threadIds = crmSplitList_(crmGet_(master, row, 'ThreadIds'));
    if (threadIds.indexOf(threadId) === -1) threadIds.push(threadId);
    crmSet_(master, row, 'ThreadIds', threadIds.slice(-300).join(', '));
    if (a.firstSent) crmSet_(master, row, 'FirstSentAt', crmMinDate_(crmGet_(master, row, 'FirstSentAt'), a.firstSent));
    if (a.lastSent) {
      var prevLastSent = crmGet_(master, row, 'LastSentAt');
      if (!(prevLastSent instanceof Date) || a.lastSent > prevLastSent) {
        crmSet_(master, row, 'LastSentAt', a.lastSent);
        if (a.lastOut) crmSet_(master, row, 'LastOutboundSnippet', a.lastOut);
      }
    }
    if (a.lastReply) {
      var prevLastReply = crmGet_(master, row, 'LastReplyAt');
      if (!(prevLastReply instanceof Date) || a.lastReply > prevLastReply) {
        crmSet_(master, row, 'LastReplyAt', a.lastReply);
        if (a.lastIn) crmSet_(master, row, 'LastInboundSnippet', a.lastIn);
      }
    }
    if (a.bounced) crmSet_(master, row, 'Bounced', true);
    crmSet_(master, row, 'Subjects', crmAddToList_(crmGet_(master, row, 'Subjects'), subject.replace(/,/g, ' ').slice(0, 80), 8));

    // NeedsReply reflects the most recent thread this contact is in.
    var prevTouch = crmGet_(master, row, 'LastTouchAt');
    if (!(prevTouch instanceof Date) || lastAt >= prevTouch) {
      crmSet_(master, row, 'NeedsReply', !lastIsMine && !!lastFrom && lastFrom.email === email);
    }
    crmSet_(master, row, 'LastTouchAt', crmMaxDate_(crmGet_(master, row, 'LastSentAt'), crmGet_(master, row, 'LastReplyAt')));
    crmSet_(master, row, 'UpdatedAt', new Date());
    count++;
  });

  // Totals come from the per-thread Participants map so overlap re-scans stay exact.
  emails.forEach(function (email) {
    var row = master.rows[master.index[email]];
    crmRebuildCountsFromInteractions_(master, row, inter, email);
  });
  return count;
}

/** Exact SentCount/ReceivedCount/ThreadCount for a contact from Interactions.Participants. */
function crmRebuildCountsFromInteractions_(master, row, inter, email) {
  var threadIds = crmSplitList_(crmGet_(master, row, 'ThreadIds'));
  var sent = 0, recv = 0, known = 0;
  threadIds.forEach(function (tid) {
    var idx = inter.index[tid];
    if (idx === undefined) return;
    var raw = crmGet_(inter, inter.rows[idx], 'Participants');
    if (!raw) return;
    try {
      var p = JSON.parse(raw)[email];
      if (p) { sent += Number(p[0]) || 0; recv += Number(p[1]) || 0; known++; }
    } catch (e) { /* malformed cell */ }
  });
  if (known) {
    crmSet_(master, row, 'SentCount', sent);
    crmSet_(master, row, 'ReceivedCount', recv);
  }
  crmSet_(master, row, 'ThreadCount', threadIds.length);
}

function crmEmptyAgg_(email) {
  return { email: email, name: '', sent: 0, received: 0, firstSent: null, lastSent: null, lastReply: null, lastOut: '', lastIn: '', bounced: false };
}

function crmSafeExcerpt_(msg, chars) {
  if (chars === 0) return '';
  try { return crmExcerpt_(msg.getPlainBody(), chars); } catch (e) { return ''; }
}

/** A mailer-daemon message in a sent thread: flag the recipient addresses mentioned in it. */
function crmMarkBounceFromDaemon_(msg, touched) {
  try {
    var body = msg.getPlainBody().toLowerCase();
    Object.keys(touched).forEach(function (email) {
      if (body.indexOf(email) !== -1) touched[email].bounced = true;
    });
  } catch (e) { /* ignore */ }
}
