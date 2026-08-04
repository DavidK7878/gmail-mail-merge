/**
 * Reply and bounce detection. Runs on a 30-minute trigger (installed by startSending)
 * or manually via Mail Merge → Check replies & bounces now.
 */

var MAX_THREAD_CHECKS_PER_RUN = 100; // keep well under the 6-minute execution cap

function checkReplies() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var sheet = getContactsSheet_();
    var map = colMap_(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    var me = myEmail_().toLowerCase();
    var checked = 0;

    for (var i = 0; i < data.length && checked < MAX_THREAD_CHECKS_PER_RUN; i++) {
      var row = i + 2;
      var status = String(data[i][map['status'] - 1]).trim();
      if (status !== 'SENT') continue;
      if (toBool_(data[i][map['replied'] - 1]) || toBool_(data[i][map['bounced'] - 1])) continue;

      var email = String(data[i][map['email'] - 1]).trim();
      var threadId = String(data[i][map['threadid'] - 1]).trim();
      checked++;

      // 1) Replies: any message on the thread from someone who isn't me.
      if (threadId) {
        try {
          var thread = GmailApp.getThreadById(threadId);
          if (thread) {
            var messages = thread.getMessages();
            for (var m = 0; m < messages.length; m++) {
              var from = messages[m].getFrom().toLowerCase();
              if (from.indexOf(me) !== -1) continue;
              if (/mailer-daemon|postmaster/.test(from)) {
                markBounced_(sheet, map, row, email, 'Bounce message found in thread');
                break;
              }
              cell_(sheet, map, row, 'Replied').setValue(true);
              logEvent('REPLY', email, 'From: ' + messages[m].getFrom());
              break;
            }
          }
        } catch (e) {
          logEvent('CHECK_ERROR', email, 'Thread lookup failed: ' + e);
        }
      }

      // 2) Bounces that arrive as separate mailer-daemon messages (skip if already flagged).
      if (!toBool_(cell_(sheet, map, row, 'Bounced').getValue())) {
        var sentAt = data[i][map['sentat'] - 1];
        var recent = sentAt instanceof Date && (Date.now() - sentAt.getTime()) < 7 * 24 * 3600 * 1000;
        if (recent) {
          try {
            var bounces = GmailApp.search('from:(mailer-daemon OR postmaster) newer_than:7d "' + email + '"', 0, 1);
            if (bounces.length) {
              markBounced_(sheet, map, row, email, 'mailer-daemon message references this address');
            }
          } catch (e) {
            logEvent('CHECK_ERROR', email, 'Bounce search failed: ' + e);
          }
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function markBounced_(sheet, map, row, email, reason) {
  cell_(sheet, map, row, 'Bounced').setValue(true);
  logEvent('BOUNCE', email, reason);
}
