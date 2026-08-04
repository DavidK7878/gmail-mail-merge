/**
 * Core sending engine: draft templates, merge, attachments, staggered batches.
 */

var BATCH_TIME_BUDGET_MS = 4.5 * 60 * 1000; // stay under the 6-minute execution cap

// ---------------------------------------------------------------- Drafts

/** Subjects of all Gmail drafts (deduped, non-empty) — the template picker list. */
function getDraftSubjects() {
  var seen = {};
  var out = [];
  GmailApp.getDrafts().forEach(function (d) {
    var subj = d.getMessage().getSubject();
    if (subj && !seen[subj]) { seen[subj] = true; out.push(subj); }
  });
  return out;
}

/**
 * Load a draft template by exact subject.
 * Returns { subject, htmlBody, attachments, inlineImages } or throws.
 */
function getDraftTemplate_(subject) {
  var drafts = GmailApp.getDrafts();
  for (var i = 0; i < drafts.length; i++) {
    var msg = drafts[i].getMessage();
    if (msg.getSubject() === subject) {
      var htmlBody = msg.getBody();
      var allInline = msg.getAttachments({ includeInlineImages: true, includeAttachments: false });
      var regular = msg.getAttachments({ includeInlineImages: false, includeAttachments: true });

      // Map cid: references in the body to their image blobs so inline images survive the merge.
      var inlineImages = {};
      var cids = [];
      htmlBody.replace(/src="cid:([^"]+)"/g, function (m, cid) { cids.push(cid); return m; });
      cids.forEach(function (cid, i) {
        if (allInline[i]) inlineImages[cid] = allInline[i].copyBlob();
      });

      return { subject: subject, htmlBody: htmlBody, attachments: regular, inlineImages: inlineImages };
    }
  }
  throw new Error('No Gmail draft found with subject "' + subject + '". Refresh the draft list in the sidebar.');
}

// ---------------------------------------------------------------- Attachments

/**
 * Resolve the row's Attachment cell into blobs.
 * Accepts Drive URLs or file names, comma-separated.
 */
function resolveAttachments_(value) {
  var blobs = [];
  String(value || '').split(',').map(function (x) { return x.trim(); }).filter(String).forEach(function (ref) {
    var idMatch = ref.match(/[-\w]{25,}/); // Drive file id inside a URL
    var file = null;
    if (/^https?:\/\//.test(ref) && idMatch) {
      file = DriveApp.getFileById(idMatch[0]);
    } else {
      var files = DriveApp.getFilesByName(ref);
      if (files.hasNext()) file = files.next();
    }
    if (!file) throw new Error('Attachment not found in Drive: "' + ref + '"');
    blobs.push(file.getBlob());
  });
  return blobs;
}

// ---------------------------------------------------------------- Start / pause

/** Called from sidebar or menu. Installs the staggered-send trigger and kicks off the first batch. */
function startSending(draftSubject) {
  getDraftTemplate_(draftSubject); // validate before committing
  var settings = getSettings_();
  if ((settings.TrackOpens || settings.TrackClicks) && !settings.WebAppUrl) {
    throw new Error('Tracking is enabled but Settings → WebAppUrl is empty. Deploy the script as a Web App and paste its URL, or turn tracking off.');
  }

  PropertiesService.getScriptProperties().setProperty('DRAFT_SUBJECT', draftSubject);

  deleteTriggersFor_('processBatch');
  ScriptApp.newTrigger('processBatch').timeBased().everyMinutes(5).create();
  if (!hasTriggerFor_('checkReplies')) {
    ScriptApp.newTrigger('checkReplies').timeBased().everyMinutes(30).create();
  }
  if (settings.MaxFollowUps > 0 && !hasTriggerFor_('processFollowUps')) {
    ScriptApp.newTrigger('processFollowUps').timeBased().everyDays(1).atHour(10).create();
  }

  logEvent('CAMPAIGN_START', draftSubject, 'Send trigger installed (every 5 min)');
  processBatch(); // start immediately rather than waiting for the first tick
  return 'Sending started.';
}

function pauseSending() {
  deleteTriggersFor_('processBatch');
  logEvent('CAMPAIGN_PAUSE', '', 'Send trigger removed');
  try { SpreadsheetApp.getUi().alert('Sending paused. Reply checking and follow-ups keep running; use "Stop everything" to remove those too.'); } catch (e) { }
}

// ---------------------------------------------------------------- Batch engine

/** Trigger handler: sends a batch of pending rows, spaced out, within window/cap/time budget. */
function processBatch() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // another batch is mid-flight
  try {
    var settings = getSettings_();
    if (!inSendWindow_(settings)) return;

    var draftSubject = PropertiesService.getScriptProperties().getProperty('DRAFT_SUBJECT');
    if (!draftSubject) return;
    var template = getDraftTemplate_(draftSubject);

    var sheet = getContactsSheet_();
    var map = colMap_(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    var started = Date.now();
    var sentThisBatch = 0;

    for (var i = 0; i < data.length; i++) {
      var row = i + 2;
      var status = String(data[i][map['status'] - 1]).trim();
      if (status) continue; // blank status = pending
      var email = String(data[i][map['email'] - 1]).trim();
      if (!email) continue;

      if (dailyAllowanceRemaining_(settings) <= 0) {
        logEvent('CAP_REACHED', '', 'Daily cap hit — remaining rows roll to the next allowed day');
        return;
      }
      if (Date.now() - started > BATCH_TIME_BUDGET_MS) return; // next trigger tick picks it up

      if (sentThisBatch > 0) {
        var pause = settings.MinSecondsBetween + Math.floor(Math.random() * (settings.MaxSecondsBetween - settings.MinSecondsBetween + 1));
        Utilities.sleep(pause * 1000);
      }

      try {
        sendToRow_(sheet, map, headers, row, data[i], email, template, settings, null);
        incrementDailyCount_();
        sentThisBatch++;
      } catch (err) {
        cell_(sheet, map, row, 'Status').setValue('FAILED');
        logEvent('SEND_FAILED', email, String(err));
      }
    }

    // Nothing pending left — campaign done, remove the send trigger.
    deleteTriggersFor_('processBatch');
    logEvent('CAMPAIGN_DONE', '', 'All rows processed; send trigger removed');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Merge + track + send one row. overrideRecipient is used by "send test to myself".
 */
function sendToRow_(sheet, map, headers, row, rowValues, email, template, settings, overrideRecipient) {
  var subject = mergeText_(template.subject, headers, rowValues);
  var htmlBody = mergeText_(template.htmlBody, headers, rowValues);

  var trackingId = Utilities.getUuid();
  if (!overrideRecipient) {
    htmlBody = injectTracking_(htmlBody, trackingId, settings);
  }

  var attachments = template.attachments.slice();
  var attCol = map['attachment'];
  if (attCol) {
    attachments = attachments.concat(resolveAttachments_(rowValues[attCol - 1]));
  }

  var options = { htmlBody: htmlBody };
  if (attachments.length) options.attachments = attachments;
  if (Object.keys(template.inlineImages).length) options.inlineImages = template.inlineImages;

  var recipient = overrideRecipient || email;
  GmailApp.sendEmail(recipient, overrideRecipient ? '[TEST] ' + subject : subject, htmlToPlainText_(htmlBody), options);

  if (overrideRecipient) {
    logEvent('TEST_SENT', recipient, 'Merged from row ' + row + ' (' + email + ')');
    return;
  }

  // Record thread id so reply detection and threaded follow-ups can find it.
  var threadId = '';
  try {
    Utilities.sleep(2000);
    var threads = GmailApp.search('in:sent to:(' + email + ') subject:"' + subject.replace(/"/g, '') + '"', 0, 1);
    if (threads.length) threadId = threads[0].getId();
  } catch (e) { /* non-fatal */ }

  cell_(sheet, map, row, 'Status').setValue('SENT');
  cell_(sheet, map, row, 'SentAt').setValue(new Date());
  cell_(sheet, map, row, 'Opens').setValue(0);
  cell_(sheet, map, row, 'Clicks').setValue(0);
  cell_(sheet, map, row, 'FollowUpStage').setValue(0);
  cell_(sheet, map, row, 'ThreadId').setValue(threadId);
  cell_(sheet, map, row, 'TrackingId').setValue(trackingId);
  logEvent('SENT', email, subject);
}

function htmlToPlainText_(html) {
  return String(html)
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|li|h[1-6])>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// ---------------------------------------------------------------- Window & cap

function inSendWindow_(settings) {
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var day = Utilities.formatDate(now, tz, 'EEE'); // Mon, Tue, ...
  var hour = Number(Utilities.formatDate(now, tz, 'H'));
  var days = settings.SendDays.split(',').map(function (d) { return d.trim().toLowerCase().slice(0, 3); });
  if (days.indexOf(day.toLowerCase()) === -1) return false;
  return hour >= settings.SendWindowStart && hour < settings.SendWindowEnd;
}

function dailyAllowanceRemaining_(settings) {
  var byCap = settings.DailyCap - todaysCount_();
  var byQuota = MailApp.getRemainingDailyQuota();
  return Math.min(byCap, byQuota);
}

function todaysCount_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('DAILY_COUNT');
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (raw) {
    var parsed = JSON.parse(raw);
    if (parsed.date === today) return parsed.count;
  }
  return 0;
}

function incrementDailyCount_() {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  props.setProperty('DAILY_COUNT', JSON.stringify({ date: today, count: todaysCount_() + 1 }));
}

// ---------------------------------------------------------------- Test send

/** Merge the first pending (or first data) row, deliver to my own inbox. */
function sendTestToMyself(draftSubject) {
  var template = getDraftTemplate_(draftSubject);
  var settings = getSettings_();
  var sheet = getContactsSheet_();
  var map = colMap_(sheet);
  if (sheet.getLastRow() < 2) throw new Error('Add at least one recipient row to the Contacts tab first.');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  var idx = 0;
  for (var i = 0; i < data.length; i++) {
    if (!String(data[i][map['status'] - 1]).trim() && String(data[i][map['email'] - 1]).trim()) { idx = i; break; }
  }
  var email = String(data[idx][map['email'] - 1]).trim() || myEmail_();
  sendToRow_(sheet, map, headers, idx + 2, data[idx], email, template, settings, myEmail_());
  return 'Test sent to ' + myEmail_() + ' using row ' + (idx + 2) + '.';
}
