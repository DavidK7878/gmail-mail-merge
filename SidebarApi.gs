/**
 * Server-side functions called from Sidebar.html via google.script.run.
 */

function getSidebarData() {
  var sheet = getContactsSheet_();
  var map = colMap_(sheet);
  var pending = 0, sent = 0;
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var statuses = sheet.getRange(2, map['status'], lastRow - 1, 1).getValues();
    var emails = sheet.getRange(2, map['email'], lastRow - 1, 1).getValues();
    for (var i = 0; i < statuses.length; i++) {
      var s = String(statuses[i][0]).trim();
      if (s === 'SENT') sent++;
      else if (!s && String(emails[i][0]).trim()) pending++;
    }
  }
  return {
    drafts: getDraftSubjects(),
    selected: PropertiesService.getScriptProperties().getProperty('DRAFT_SUBJECT') || '',
    pending: pending,
    sent: sent,
    quotaRemaining: MailApp.getRemainingDailyQuota(),
    sending: hasTriggerFor_('processBatch')
  };
}

/** Merge the first pending row against the chosen draft — no tracking, nothing sent. */
function previewMerge(draftSubject) {
  var template = getDraftTemplate_(draftSubject);
  var sheet = getContactsSheet_();
  var map = colMap_(sheet);
  if (sheet.getLastRow() < 2) throw new Error('Add at least one recipient row to the Contacts tab first.');

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  var idx = 0;
  for (var i = 0; i < data.length; i++) {
    if (!String(data[i][map['status'] - 1]).trim() && String(data[i][map['email'] - 1]).trim()) { idx = i; break; }
  }
  return {
    row: idx + 2,
    to: String(data[idx][map['email'] - 1]).trim(),
    subject: mergeText_(template.subject, headers, data[idx]),
    htmlBody: mergeText_(template.htmlBody, headers, data[idx])
  };
}
