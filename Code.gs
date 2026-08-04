/**
 * Mail Merge for Google Sheets — main entry points, menu, setup, shared helpers.
 *
 * Tabs: Contacts | Settings | Dashboard | Log
 * Templates are ordinary Gmail drafts with {{Column}} placeholders.
 */

var MACHINE_COLUMNS = ['Status', 'SentAt', 'Opens', 'LastOpenAt', 'Clicks', 'Replied', 'Bounced', 'FollowUpStage', 'ThreadId', 'TrackingId'];
var USER_COLUMNS = ['Email', 'FirstName', 'Company', 'Attachment'];

// ---------------------------------------------------------------- Menu / UI

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mail Merge')
    .addItem('Open sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Start sending', 'menuStartSending')
    .addItem('Pause sending', 'pauseSending')
    .addSeparator()
    .addItem('Check replies & bounces now', 'checkReplies')
    .addItem('Send due follow-ups now', 'processFollowUps')
    .addSeparator()
    .addItem('Set up / repair sheet tabs', 'setupSheet')
    .addItem('Stop everything (remove all triggers)', 'stopEverything')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Mail Merge');
  SpreadsheetApp.getUi().showSidebar(html);
}

function menuStartSending() {
  var subject = PropertiesService.getScriptProperties().getProperty('DRAFT_SUBJECT');
  if (!subject) {
    SpreadsheetApp.getUi().alert('Pick a draft template in the sidebar first (Mail Merge → Open sidebar).');
    return;
  }
  startSending(subject);
}

// ---------------------------------------------------------------- Setup

/** Creates/repairs the four tabs. Safe to re-run; never deletes user data. */
function setupSheet() {
  var ss = SpreadsheetApp.getActive();

  // Contacts
  var contacts = ss.getSheetByName('Contacts') || ss.insertSheet('Contacts', 0);
  var wanted = USER_COLUMNS.concat(MACHINE_COLUMNS);
  var lastCol = Math.max(contacts.getLastColumn(), 1);
  var existing = contacts.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim(); })
    .filter(String);
  if (existing.length === 0) {
    contacts.getRange(1, 1, 1, wanted.length).setValues([wanted]);
    contacts.getRange(1, 1, 1, wanted.length).setFontWeight('bold');
    contacts.setFrozenRows(1);
  } else {
    // Append any missing machine columns to the right, leave user columns alone.
    var missing = MACHINE_COLUMNS.filter(function (c) {
      return existing.map(function (h) { return h.toLowerCase(); }).indexOf(c.toLowerCase()) === -1;
    });
    if (missing.length) {
      contacts.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
      contacts.getRange(1, existing.length + 1, 1, missing.length).setFontWeight('bold');
    }
  }

  // Settings
  var settings = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  if (settings.getLastRow() === 0) {
    var defaults = [
      ['Setting', 'Value', 'Notes'],
      ['DailyCap', 90, 'Max emails per day. Consumer Gmail allows ~100/day via Apps Script; Workspace 1,500.'],
      ['MinSecondsBetween', 30, 'Minimum pause between two sends.'],
      ['MaxSecondsBetween', 90, 'Maximum pause between two sends (actual pause is random in range).'],
      ['TrackOpens', true, 'Inject invisible open-tracking pixel. Needs WebAppUrl.'],
      ['TrackClicks', false, 'Rewrite links through tracker (visible URL becomes script.google.com). Needs WebAppUrl.'],
      ['SendWindowStart', 8, 'Earliest hour (0-23) to send, script timezone.'],
      ['SendWindowEnd', 17, 'Sending stops at this hour (0-23).'],
      ['SendDays', 'Mon,Tue,Wed,Thu,Fri', 'Days sending is allowed.'],
      ['FollowUpWaitDays', 4, 'Days after a send (or previous follow-up) before the next follow-up.'],
      ['MaxFollowUps', 2, 'Max follow-ups per recipient. 0 disables follow-ups.'],
      ['FollowUpDraftSubjects', '', 'Comma-separated Gmail draft subjects, one per follow-up stage. Last one repeats.'],
      ['WebAppUrl', '', 'Paste the Web App URL here after deploying (Deploy → New deployment → Web app).']
    ];
    settings.getRange(1, 1, defaults.length, 3).setValues(defaults);
    settings.getRange(1, 1, 1, 3).setFontWeight('bold');
    settings.setFrozenRows(1);
    settings.setColumnWidth(3, 480);
  }

  // Dashboard
  var dash = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');
  if (dash.getLastRow() === 0) {
    var C = function (header, crit) {
      return '=IFERROR(COUNTIF(INDEX(Contacts!A:Z,0,MATCH("' + header + '",Contacts!1:1,0)),' + crit + '),0)';
    };
    var rows = [
      ['Metric', 'Value'],
      ['Recipients', '=IFERROR(COUNTA(INDEX(Contacts!A:Z,0,MATCH("Email",Contacts!1:1,0)))-1,0)'],
      ['Sent', C('Status', '"SENT"')],
      ['Failed', C('Status', '"FAILED"')],
      ['Pending', '=MAX(0,B2-B3-B4)'],
      ['Opened', C('Opens', '">0"')],
      ['Open rate', '=IFERROR(B6/B3,"")'],
      ['Clicked', C('Clicks', '">0"')],
      ['Click rate', '=IFERROR(B8/B3,"")'],
      ['Replied', C('Replied', 'TRUE')],
      ['Reply rate', '=IFERROR(B10/B3,"")'],
      ['Bounced', C('Bounced', 'TRUE')],
      ['Follow-ups sent', '=IFERROR(SUM(INDEX(Contacts!A:Z,0,MATCH("FollowUpStage",Contacts!1:1,0))),0)']
    ];
    dash.getRange(1, 1, rows.length, 2).setValues(rows);
    dash.getRange(1, 1, 1, 2).setFontWeight('bold');
    dash.getRange('B7').setNumberFormat('0.0%');
    dash.getRange('B9').setNumberFormat('0.0%');
    dash.getRange('B11').setNumberFormat('0.0%');
  }

  // Log
  var log = ss.getSheetByName('Log') || ss.insertSheet('Log');
  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Event', 'Target', 'Details']]);
    log.getRange(1, 1, 1, 4).setFontWeight('bold');
    log.setFrozenRows(1);
  }

  logEvent('SETUP', '', 'Sheet tabs created/verified');
  try { SpreadsheetApp.getUi().alert('Mail Merge tabs are ready. Fill the Contacts tab, review Settings, then open the sidebar.'); } catch (e) { /* headless */ }
}

// ---------------------------------------------------------------- Shared helpers

function getSettings_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  if (!sheet) throw new Error('Settings tab missing — run Mail Merge → Set up / repair sheet tabs.');
  var values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).getValues();
  var s = {};
  values.forEach(function (r) {
    var key = String(r[0]).trim();
    if (key) s[key] = r[1];
  });
  s.DailyCap = Number(s.DailyCap) || 90;
  s.MinSecondsBetween = Number(s.MinSecondsBetween) || 30;
  s.MaxSecondsBetween = Math.max(Number(s.MaxSecondsBetween) || 90, s.MinSecondsBetween);
  s.TrackOpens = toBool_(s.TrackOpens);
  s.TrackClicks = toBool_(s.TrackClicks);
  s.SendWindowStart = Number(s.SendWindowStart);
  if (isNaN(s.SendWindowStart)) s.SendWindowStart = 8;
  s.SendWindowEnd = Number(s.SendWindowEnd);
  if (isNaN(s.SendWindowEnd)) s.SendWindowEnd = 17;
  s.SendDays = String(s.SendDays || 'Mon,Tue,Wed,Thu,Fri');
  s.FollowUpWaitDays = Number(s.FollowUpWaitDays) || 4;
  s.MaxFollowUps = Number(s.MaxFollowUps) || 0;
  s.FollowUpDraftSubjects = String(s.FollowUpDraftSubjects || '').split(',').map(function (x) { return x.trim(); }).filter(String);
  s.WebAppUrl = String(s.WebAppUrl || '').trim();
  return s;
}

function toBool_(v) {
  if (typeof v === 'boolean') return v;
  return String(v).trim().toLowerCase() === 'true';
}

function getContactsSheet_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Contacts');
  if (!sheet) throw new Error('Contacts tab missing — run Mail Merge → Set up / repair sheet tabs.');
  return sheet;
}

/** Map of lowercased header -> 1-based column index. */
function colMap_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  headers.forEach(function (h, i) {
    var key = String(h).trim().toLowerCase();
    if (key && !(key in map)) map[key] = i + 1;
  });
  MACHINE_COLUMNS.concat(['Email']).forEach(function (c) {
    if (!(c.toLowerCase() in map)) throw new Error('Contacts tab is missing the "' + c + '" column — run Mail Merge → Set up / repair sheet tabs.');
  });
  return map;
}

function cell_(sheet, map, row, header) {
  return sheet.getRange(row, map[header.toLowerCase()]);
}

/** Replace {{Header}} placeholders with row values (case-insensitive, tolerant of spaces). */
function mergeText_(text, headers, rowValues) {
  var lookup = {};
  headers.forEach(function (h, i) {
    lookup[String(h).trim().toLowerCase()] = rowValues[i];
  });
  return String(text).replace(/\{\{\s*([^{}]+?)\s*\}\}/g, function (m, name) {
    var key = name.toLowerCase();
    if (key in lookup) {
      var v = lookup[key];
      return (v === null || v === undefined) ? '' : String(v);
    }
    return ''; // unknown placeholder → blank rather than leaking {{...}}
  });
}

function logEvent(event, target, details) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('Log');
    if (sheet) sheet.appendRow([new Date(), event, target || '', details || '']);
  } catch (e) { /* never let logging break the pipeline */ }
}

function myEmail_() {
  return Session.getEffectiveUser().getEmail();
}

function stopEverything() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  logEvent('TRIGGERS', '', 'All triggers removed');
  try { SpreadsheetApp.getUi().alert('All triggers removed. Sending, reply checks, and follow-ups are stopped.'); } catch (e) { }
}

function deleteTriggersFor_(fnName) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === fnName) ScriptApp.deleteTrigger(t);
  });
}

function hasTriggerFor_(fnName) {
  return ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === fnName; });
}
