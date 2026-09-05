/**
 * Gmail scan → Master + Interactions.
 *
 * Driver design (why it looks the way it does):
 *  - Gmail search orders threads by most recent activity, so offset paging over one big
 *    query is NOT stable while mail arrives. The scan therefore walks the date range in
 *    windows of Settings → ScanWindowDays: `after:<epoch> before:<epoch>` matches on the
 *    *sent message's* date, which never changes, so a window's result set is stable.
 *  - Phase 1 query is `in:sent` (threads where we wrote). Incremental scans add phase 2,
 *    `-in:sent …` over the same window, to pick up replies that arrived on older threads.
 *    Threads without any outbound message are skipped, so phase 2 costs one getMessages
 *    per inbound thread and adds nothing else.
 *  - A thread whose message count and last-message date match the Interactions row is
 *    skipped without reading its messages (re-scans are cheap and idempotent).
 *  - State lives in Script Properties; a 1-minute trigger keeps calling crmScanStep until
 *    done. Each tick stops at ~4.5 minutes. State is only written back if the job is still
 *    the current one (Stop / a newer Start during a tick win). Three consecutive tick
 *    failures abort the job instead of running forever.
 */

var CRM_SCAN_STATE_PROP = 'CRM_SCAN_STATE';
var CRM_LAST_SCAN_PROP = 'CRM_LAST_SCAN_AT';
var CRM_SEARCH_PAGE = 100;   // threads per GmailApp.search page (max 500)
var CRM_SCAN_OVERLAP_MS = 2 * 86400000;
var CRM_PHASE_QUERY = { 1: 'in:sent', 2: '-in:sent -in:chats -in:drafts -in:spam -in:trash' };

function crmScanIncremental() { var m = crmStartScan(false); crmToast_(m); return m; }
function crmScanFull() { var m = crmStartScan(true); crmToast_(m); return m; }

/** Menu/sidebar entry point. Returns a status string. */
function crmStartScan(fullRescan) {
  var settings = crmSettings_();
  var props = PropertiesService.getScriptProperties();
  var startedAt = Date.now();

  var sinceMs = settings.ScanSince.getTime();
  if (!fullRescan) {
    var last = props.getProperty(CRM_LAST_SCAN_PROP);
    if (last) sinceMs = Math.max(sinceMs, Number(last) - CRM_SCAN_OVERLAP_MS);
  }
  var state = {
    full: !!fullRescan, startedAt: startedAt, sinceMs: sinceMs, untilMs: startedAt,
    windowDays: settings.ScanWindowDays, phase: 1, cursorMs: sinceMs, offset: 0,
    threads: 0, contacts: 0, skipped: 0, errors: 0, done: false
  };
  props.setProperty(CRM_SCAN_STATE_PROP, JSON.stringify(state));
  crmEnsureMinuteTrigger_('crmScanStep');
  var tz = Session.getScriptTimeZone();
  crmLog_('SCAN_START', Utilities.formatDate(new Date(sinceMs), tz, 'yyyy-MM-dd') + ' → now', fullRescan ? 'Full rescan' : 'Incremental scan (sent + inbound on existing threads)');
  crmScanStep();
  return 'Scan started from ' + Utilities.formatDate(new Date(sinceMs), tz, 'yyyy-MM-dd') + '. Progress shows in the sidebar; you can close the sheet.';
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
  var state = raw ? JSON.parse(raw) : null;
  if (state) state.progress = crmScanProgress_(state);
  return { running: !!raw, state: state, lastScanAt: last ? Number(last) : null };
}

/** 0..1 across both phases (phase 2 only exists for incremental scans). */
function crmScanProgress_(state) {
  var span = Math.max(1, state.untilMs - state.sinceMs);
  var phases = state.full ? 1 : 2;
  var within = Math.min(1, Math.max(0, (state.cursorMs - state.sinceMs) / span));
  return Math.min(1, ((state.phase - 1) + within) / phases);
}

/** Pure: the next window [start, end) and query for the current state, or null when the phase is finished. */
function crmScanWindow_(state) {
  if (state.cursorMs >= state.untilMs) return null;
  var end = Math.min(state.untilMs, state.cursorMs + state.windowDays * 86400000);
  return {
    start: state.cursorMs, end: end,
    query: CRM_PHASE_QUERY[state.phase] + ' after:' + Math.floor(state.cursorMs / 1000) + ' before:' + Math.ceil(end / 1000)
  };
}

/** Trigger handler. */
function crmScanStep() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(CRM_SCAN_STATE_PROP);
    if (!raw) { crmDeleteTriggersFor_('crmScanStep'); return; }
    var state = JSON.parse(raw);
    try {
      crmScanTick_(state);
      state.errors = 0;
    } catch (e) {
      state.errors = (state.errors || 0) + 1;
      crmLog_('SCAN_ERROR', 'tick', String(e && e.stack || e));
      if (state.errors >= 3) {
        props.deleteProperty(CRM_SCAN_STATE_PROP);
        crmDeleteTriggersFor_('crmScanStep');
        crmLog_('SCAN_ABORT', '', 'Three consecutive tick failures — scan stopped. Fix the error above and start again.');
        return;
      }
    }
    if (state.done) {
      if (crmCommitState_(CRM_SCAN_STATE_PROP, state)) {
        props.deleteProperty(CRM_SCAN_STATE_PROP);
        props.setProperty(CRM_LAST_SCAN_PROP, String(state.untilMs));
        crmDeleteTriggersFor_('crmScanStep');
        crmLog_('SCAN_DONE', '', state.threads + ' threads read, ' + state.skipped + ' unchanged threads skipped, ' + state.contacts + ' contact updates');
        crmApplyValidation_(SpreadsheetApp.getActive());
        // Queue classification for the next trigger tick — never run it inside this execution's time window.
        if (crmSettings_().AutoClassifyAfterScan) crmStartClassify_('unclassified', true);
      } else {
        crmLog_('SCAN_SUPERSEDED', '', 'Scan finished but had been stopped or replaced meanwhile; results kept, state not written');
      }
    } else if (crmCommitState_(CRM_SCAN_STATE_PROP, state)) {
      crmEnsureMinuteTrigger_('crmScanStep');
    } else {
      crmLog_('SCAN_SUPERSEDED', '', 'Tick finished after Stop / a newer Start; its state was discarded');
    }
  } finally {
    lock.releaseLock();
  }
}

/** One tick: process windows until the time budget is spent or the job is done. Mutates state. */
function crmScanTick_(state) {
  var settings = crmSettings_();
  var started = Date.now();
  var ctx = crmScanContext_(settings);
  var master = crmLoadTable_('Master', CRM_MASTER_COLUMNS, function (r, t) { return crmNormalizeEmail_(crmGet_(t, r, 'Email')); });
  var inter = crmLoadTable_('Interactions', CRM_INTERACTION_COLUMNS, function (r, t) { return String(crmGet_(t, r, 'ThreadId')).trim(); });
  var processed = 0;

  try {
    while (processed < settings.MaxThreadsPerRun && Date.now() - started < CRM_TIME_BUDGET_MS) {
      var win = crmScanWindow_(state);
      if (!win) {
        if (state.phase === 1 && !state.full) { state.phase = 2; state.cursorMs = state.sinceMs; state.offset = 0; continue; }
        state.done = true;
        break;
      }
      var threads = GmailApp.search(win.query, state.offset, CRM_SEARCH_PAGE);
      var i = 0;
      for (; i < threads.length; i++) {
        if (Date.now() - started > CRM_TIME_BUDGET_MS || processed >= settings.MaxThreadsPerRun) break;
        try {
          var r = crmProcessThread_(threads[i], master, inter, ctx);
          if (r === null) state.skipped++; else { state.contacts += r; state.threads++; }
        } catch (e) {
          crmLog_('SCAN_ERROR', threads[i].getId(), String(e));
        }
        state.offset++; processed++;
      }
      var pageComplete = i === threads.length;
      if (pageComplete && threads.length < CRM_SEARCH_PAGE) {
        // Window exhausted → advance. Only when every thread of the final page was handled.
        state.cursorMs = win.end; state.offset = 0;
      }
      // Otherwise: resume this window at state.offset on the next loop/tick.
    }
  } finally {
    // Always persist what was processed, even if a later thread threw.
    if (master.dirty || inter.dirty) {
      crmSaveTable_(inter);
      if (state.done || master.dirty) { crmRecomputePipeline_(master); crmSaveTable_(master); }
    }
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
  // Your own company domain counts as "us" so teammates on Cc aren't contacts — unless it is a
  // consumer mail provider (then everyone there would be dropped) or the setting is off.
  var myDomain = crmDomainOf_(me);
  var treat = settings.TreatMyDomainAsTeam === undefined ? true : settings.TreatMyDomainAsTeam;
  if (treat && myDomain && !crmIsFreemail_(myDomain) && myDomains.indexOf(myDomain) === -1) myDomains.push(myDomain);

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
 * Fold one thread into Master + Interactions.
 * Returns the number of contact rows touched, or null if the thread was unchanged and skipped.
 */
function crmProcessThread_(thread, master, inter, ctx) {
  var threadId = thread.getId();

  // Cheap skip: nothing changed since we last stored this thread.
  var existing = inter.index[threadId] !== undefined ? inter.rows[inter.index[threadId]] : null;
  if (existing) {
    try {
      var storedLast = crmGet_(inter, existing, 'LastMessageAt');
      var storedCount = Number(crmGet_(inter, existing, 'Messages')) || 0;
      var lastDate = thread.getLastMessageDate();
      if (storedLast instanceof Date && lastDate instanceof Date && storedLast.getTime() === lastDate.getTime() && storedCount === thread.getMessageCount()) return null;
    } catch (e) { /* fall through to a full read */ }
  }

  var messages = thread.getMessages();
  if (!messages.length) return 0;
  var subject = thread.getFirstMessageSubject() || '(no subject)';

  var touched = {};   // email → per-thread aggregate
  var outbound = 0, inbound = 0, autoReplies = 0, hasOutbound = false, lastIsAuto = false;
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
        var a = touched[r.email] || (touched[r.email] = crmEmptyAgg_(r.email, r.raw));
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
      if (ctx.isIgnored(from.email)) {
        if (/mailer-daemon|postmaster/.test(from.email)) crmMarkBounceFromDaemon_(msg, touched);
        continue;
      }
      var a2 = touched[from.email] || (touched[from.email] = crmEmptyAgg_(from.email, from.raw));
      if (from.name && !a2.name) a2.name = from.name;
      if (crmIsAutoReply_(msg.getSubject(), crmMessageHeaders_(msg))) {
        a2.autoReplies++;
        autoReplies++;
        if (m === messages.length - 1) lastIsAuto = true;
        continue; // an out-of-office is not a reply
      }
      inbound++;
      a2.received++;
      var split = crmSafeSplit_(msg);
      if (!a2.lastReply || date > a2.lastReply) {
        a2.lastReply = date;
        a2.lastIn = crmClip_(split.body, ctx.snippetChars);
      }
      if (split.signature) a2.signature = split.signature;
      if (crmIsOptOut_(split.body)) a2.optOut = true;
    }
  }

  // A thread with no outbound message from us is not a relationship we own (phase 2 noise).
  if (!hasOutbound) return 0;
  var emails = Object.keys(touched);
  if (!emails.length) return 0;

  // Interactions row (one per thread; the Email column holds the primary counterpart).
  var primary = emails.slice().sort(function (x, y) { return (touched[y].sent + touched[y].received) - (touched[x].sent + touched[x].received); })[0];
  var irow = existing || crmAppendRow_(inter, threadId);
  crmSet_(inter, irow, 'ThreadId', threadId);
  crmSet_(inter, irow, 'Email', primary);
  crmSet_(inter, irow, 'Subject', subject);
  crmSet_(inter, irow, 'FirstMessageAt', messages[0].getDate());
  crmSet_(inter, irow, 'LastMessageAt', lastMsg.getDate());
  crmSet_(inter, irow, 'Messages', messages.length);
  crmSet_(inter, irow, 'Outbound', outbound);
  crmSet_(inter, irow, 'Inbound', inbound);
  crmSet_(inter, irow, 'AutoReplies', autoReplies);
  crmSet_(inter, irow, 'LastDirection', lastIsMine ? 'OUT' : (lastIsAuto ? 'AUTO' : 'IN'));
  crmSet_(inter, irow, 'LastFrom', lastFrom ? lastFrom.email : '');
  crmSet_(inter, irow, 'LastSnippet', crmSafeExcerpt_(lastMsg, ctx.snippetChars));
  var latestIn = emails.map(function (e) { return touched[e]; }).filter(function (a) { return a.lastReply; })
    .sort(function (x, y) { return y.lastReply - x.lastReply; })[0];
  crmSet_(inter, irow, 'LastInboundSnippet', latestIn ? latestIn.lastIn : '');
  crmSet_(inter, irow, 'Link', 'https://mail.google.com/mail/u/0/#all/' + threadId);
  var participants = {};
  emails.forEach(function (e) { participants[e] = [touched[e].sent, touched[e].received, touched[e].autoReplies]; });
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
    // re-scan of the same thread (overlap window / phase 2) never double counts.
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
    if (a.signature) crmSet_(master, row, 'Signature', a.signature);
    if (a.optOut && !crmBool_(crmGet_(master, row, 'DoNotContact'))) {
      crmSet_(master, row, 'DoNotContact', true);
      crmSet_(master, row, 'Tags', crmAddToList_(crmGet_(master, row, 'Tags'), 'optout'));
      crmLog_('OPT_OUT', email, 'Inbound message matched an opt-out phrase; DoNotContact set');
    }
    crmSet_(master, row, 'Subjects', crmAddToList_(crmGet_(master, row, 'Subjects'), subject.replace(/,/g, ' ').slice(0, 80), 8));

    // NeedsReply reflects the most recent thread this contact is in.
    var prevTouch = crmGet_(master, row, 'LastTouchAt');
    if (!(prevTouch instanceof Date) || lastAt >= prevTouch) {
      crmSet_(master, row, 'NeedsReply', !lastIsMine && !lastIsAuto && !!lastFrom && lastFrom.email === email);
    }
    crmSet_(master, row, 'LastTouchAt', crmMaxDate_(crmGet_(master, row, 'LastSentAt'), crmGet_(master, row, 'LastReplyAt')));
    crmSet_(master, row, 'UpdatedAt', new Date());
    count++;
  });

  // Totals come from the per-thread Participants map so overlap re-scans stay exact.
  emails.forEach(function (email) {
    crmRebuildCountsFromInteractions_(master, master.rows[master.index[email]], inter, email);
  });
  return count;
}

/** Exact SentCount/ReceivedCount/AutoReplies/ThreadCount for a contact from Interactions.Participants. */
function crmRebuildCountsFromInteractions_(master, row, inter, email) {
  var threadIds = crmSplitList_(crmGet_(master, row, 'ThreadIds'));
  var sent = 0, recv = 0, auto = 0, known = 0;
  threadIds.forEach(function (tid) {
    var idx = inter.index[tid];
    if (idx === undefined) return;
    var raw = crmGet_(inter, inter.rows[idx], 'Participants');
    if (!raw) return;
    try {
      var p = JSON.parse(raw)[email];
      if (p) { sent += Number(p[0]) || 0; recv += Number(p[1]) || 0; auto += Number(p[2]) || 0; known++; }
    } catch (e) { /* malformed cell */ }
  });
  if (known) {
    crmSet_(master, row, 'SentCount', sent);
    crmSet_(master, row, 'ReceivedCount', recv);
    crmSet_(master, row, 'AutoReplies', auto);
  }
  crmSet_(master, row, 'ThreadCount', threadIds.length);
}

function crmEmptyAgg_(email, raw) {
  return { email: email, raw: raw || email, name: '', sent: 0, received: 0, autoReplies: 0, firstSent: null, lastSent: null, lastReply: null, lastOut: '', lastIn: '', signature: '', optOut: false, bounced: false };
}

function crmSafeExcerpt_(msg, chars) {
  if (chars === 0) return '';
  try { return crmExcerpt_(msg.getPlainBody(), chars); } catch (e) { return ''; }
}

/** { body, signature } for a message; never throws. */
function crmSafeSplit_(msg) {
  try { return crmSplitBody_(msg.getPlainBody()); } catch (e) { return { body: '', signature: '' }; }
}

/** Headers relevant to auto-reply detection; getHeader is cheap but guarded for older runtimes. */
function crmMessageHeaders_(msg) {
  var h = {};
  ['Auto-Submitted', 'X-Autoreply', 'X-Autorespond', 'X-Auto-Response-Suppress', 'Precedence'].forEach(function (name) {
    try { var v = msg.getHeader(name); if (v) h[name] = v; } catch (e) { /* not supported */ }
  });
  return h;
}

/**
 * A mailer-daemon message in a sent thread: flag every recipient whose address appears in it.
 * Addresses found in the body are normalized before comparison so dotted/tagged Gmail forms match.
 */
function crmMarkBounceFromDaemon_(msg, touched) {
  try {
    var body = String(msg.getPlainBody() || '').toLowerCase();
    var found = {};
    (body.match(/[a-z0-9._%+\-']+@[a-z0-9.\-]+\.[a-z]{2,}/g) || []).forEach(function (e) { var n = crmNormalizeEmail_(e); if (n) found[n] = true; });
    Object.keys(touched).forEach(function (email) {
      if (found[email] || body.indexOf(String(touched[email].raw).toLowerCase()) !== -1) touched[email].bounced = true;
    });
  } catch (e) { /* ignore */ }
}
