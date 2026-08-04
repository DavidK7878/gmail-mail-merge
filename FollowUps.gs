/**
 * Automatic follow-up sequences.
 *
 * A daily trigger (installed by startSending when MaxFollowUps > 0) finds rows that
 * were sent N+ days ago with no reply and no bounce, and sends the follow-up template
 * as a reply in the original thread. checkReplies() runs first so nobody who answered
 * gets a follow-up.
 *
 * Threaded replies use the Gmail advanced service (Services → add "Gmail" in the
 * Apps Script editor). If it isn't enabled, the follow-up falls back to a normal
 * "Re: <subject>" email to the same address.
 */

function processFollowUps() {
  var settings = getSettings_();
  if (settings.MaxFollowUps <= 0) return;
  if (!settings.FollowUpDraftSubjects.length) {
    logEvent('FOLLOWUP_SKIP', '', 'MaxFollowUps > 0 but FollowUpDraftSubjects is empty in Settings');
    return;
  }
  if (!inSendWindow_(settings)) return;

  checkReplies(); // always sweep for replies immediately before following up

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var sheet = getContactsSheet_();
    var map = colMap_(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    var started = Date.now();

    // Cache follow-up templates by stage (stage n uses the nth subject; last repeats).
    var templates = {};

    for (var i = 0; i < data.length; i++) {
      var row = i + 2;
      if (String(data[i][map['status'] - 1]).trim() !== 'SENT') continue;
      if (toBool_(data[i][map['replied'] - 1]) || toBool_(data[i][map['bounced'] - 1])) continue;

      var stage = Number(data[i][map['followupstage'] - 1]) || 0;
      if (stage >= settings.MaxFollowUps) continue;

      // Wait period counts from the last touch: original send or previous follow-up.
      var sentAt = data[i][map['sentat'] - 1];
      if (!(sentAt instanceof Date)) continue;
      var lastTouch = sentAt.getTime() + stage * settings.FollowUpWaitDays * 24 * 3600 * 1000;
      var due = lastTouch + settings.FollowUpWaitDays * 24 * 3600 * 1000;
      if (Date.now() < due) continue;

      if (dailyAllowanceRemaining_(settings) <= 0) {
        logEvent('CAP_REACHED', '', 'Daily cap hit during follow-ups — resuming tomorrow');
        return;
      }
      if (Date.now() - started > BATCH_TIME_BUDGET_MS) return;

      var email = String(data[i][map['email'] - 1]).trim();
      var threadId = String(data[i][map['threadid'] - 1]).trim();

      var subjectKey = settings.FollowUpDraftSubjects[Math.min(stage, settings.FollowUpDraftSubjects.length - 1)];
      try {
        if (!templates[subjectKey]) templates[subjectKey] = getDraftTemplate_(subjectKey);
        var template = templates[subjectKey];
        var htmlBody = mergeText_(template.htmlBody, headers, data[i]);

        var trackingId = String(data[i][map['trackingid'] - 1]).trim();
        if (trackingId) htmlBody = injectTracking_(htmlBody, trackingId, settings);

        sendFollowUp_(email, threadId, htmlBody, mergeText_(template.subject, headers, data[i]));

        cell_(sheet, map, row, 'FollowUpStage').setValue(stage + 1);
        incrementDailyCount_();
        logEvent('FOLLOWUP_SENT', email, 'Stage ' + (stage + 1) + ' — ' + subjectKey);

        Utilities.sleep((settings.MinSecondsBetween + Math.floor(Math.random() * 15)) * 1000);
      } catch (err) {
        logEvent('FOLLOWUP_FAILED', email, String(err));
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Send inside the original thread when possible (Gmail advanced service),
 * otherwise fall back to a plain "Re:" email.
 */
function sendFollowUp_(email, threadId, htmlBody, fallbackSubject) {
  if (threadId) {
    try {
      var thread = GmailApp.getThreadById(threadId);
      if (thread) {
        var messages = thread.getMessages();
        var last = messages[messages.length - 1];
        var messageIdHeader = last.getHeader('Message-ID');
        var subject = 'Re: ' + thread.getFirstMessageSubject().replace(/^(Re:\s*)+/i, '');

        var raw =
          'To: ' + email + '\r\n' +
          'Subject: =?UTF-8?B?' + Utilities.base64Encode(subject, Utilities.Charset.UTF_8) + '?=\r\n' +
          (messageIdHeader ? 'In-Reply-To: ' + messageIdHeader + '\r\n' : '') +
          (messageIdHeader ? 'References: ' + messageIdHeader + '\r\n' : '') +
          'Content-Type: text/html; charset=UTF-8\r\n' +
          '\r\n' +
          htmlBody;

        Gmail.Users.Messages.send(
          { raw: Utilities.base64EncodeWebSafe(raw, Utilities.Charset.UTF_8), threadId: threadId },
          'me'
        );
        return;
      }
    } catch (e) {
      logEvent('FOLLOWUP_INFO', email, 'Threaded reply unavailable (' + e + ') — sending as new "Re:" email');
    }
  }
  GmailApp.sendEmail(email, 'Re: ' + fallbackSubject.replace(/^(Re:\s*)+/i, ''), htmlToPlainText_(htmlBody), { htmlBody: htmlBody });
}
