#!/usr/bin/env node
/**
 * Tests for the CRM .gs files. Run: node crm/tests/run.js
 *
 * The Apps Script files are loaded into a vm sandbox with fakes for the Google services
 * they touch. Pure helpers are tested directly; the scan and classifier are exercised
 * end-to-end through crmProcessThread_ and crmClassifyBatch_ with fake Gmail threads
 * and a fake UrlFetchApp.
 */
var fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');

// ---------------------------------------------------------------- sandbox
var logs = [];
var fetchQueue = [];   // fake UrlFetchApp responses, consumed in order
var fetchCalls = [];
var ctx = {
  console: console,
  Utilities: {
    formatDate: function (d) { return d.toISOString().slice(0, 10); },
    sleep: function () {},
    getUuid: function () { return 'uuid'; }
  },
  Session: {
    getScriptTimeZone: function () { return 'UTC'; },
    getEffectiveUser: function () { return { getEmail: function () { return 'me@startup.io'; } }; }
  },
  SpreadsheetApp: { getActive: function () { return { getSheetByName: function (n) { return n === 'Log' ? { appendRow: function (r) { logs.push(r); } } : null; }, toast: function () {} }; } },
  PropertiesService: { getScriptProperties: function () { return { getProperty: function () { return null; }, setProperty: function () {}, deleteProperty: function () {} }; } },
  ScriptApp: { getProjectTriggers: function () { return []; } },
  UrlFetchApp: {
    fetch: function (url, options) {
      fetchCalls.push({ url: url, options: options, body: JSON.parse(options.payload) });
      var next = fetchQueue.shift();
      if (!next) throw new Error('fake UrlFetchApp: no response queued');
      return { getResponseCode: function () { return next.code; }, getContentText: function () { return typeof next.body === 'string' ? next.body : JSON.stringify(next.body); } };
    }
  },
  LockService: {}, GmailApp: {}, DriveApp: {}, HtmlService: {}, MimeType: {}
};
vm.createContext(ctx);
['Crm.gs', 'CrmScan.gs', 'CrmClassify.gs', 'CrmPipeline.gs', 'CrmImport.gs'].forEach(function (f) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), ctx, { filename: f });
});
var CtxDate = vm.runInContext('Date', ctx);
function eq(a, b, msg) { assert.strictEqual(JSON.stringify(a), JSON.stringify(b), msg); }

var passed = 0, failed = 0;
function test(name, fn) {
  try { fn(); passed++; }
  catch (e) { failed++; console.error('FAIL', name, '\n  ', e.message); process.exitCode = 1; }
}

// ---------------------------------------------------------------- fakes: tables, gmail
function makeTable(headers) {
  var map = {}; headers.forEach(function (h, i) { map[h.toLowerCase()] = i + 1; });
  return { headers: headers, map: map, rows: [], index: {}, dirty: false, appended: 0 };
}
function addRow(t, o, key) {
  var row = t.headers.map(function () { return ''; });
  Object.keys(o).forEach(function (k) { row[t.map[k.toLowerCase()] - 1] = o[k]; });
  t.rows.push(row); if (key) t.index[key] = t.rows.length - 1; return row;
}
function get(t, row, h) { return row[t.map[h.toLowerCase()] - 1]; }
function rowFor(t, key) { return t.rows[t.index[key]]; }

function fakeMessage(o) {
  return {
    getFrom: function () { return o.from; },
    getTo: function () { return o.to || ''; },
    getCc: function () { return o.cc || ''; },
    getDate: function () { return new CtxDate(o.date); },
    getSubject: function () { return o.subject || 'Hello'; },
    getPlainBody: function () { return o.body || ''; },
    getHeader: function (name) { return (o.headers || {})[name] || ''; }
  };
}
function fakeThread(id, subject, messages) {
  return {
    getId: function () { return id; },
    getFirstMessageSubject: function () { return subject; },
    getMessages: function () { return messages.map(fakeMessage); },
    getMessageCount: function () { return messages.length; },
    getLastMessageDate: function () { return new CtxDate(messages[messages.length - 1].date); },
    _messages: messages
  };
}

// A minimal Sheet: 2-D array with the Range methods the code uses.
function FakeSheet(name, data) { this.name = name; this.data = data || []; }
FakeSheet.prototype.getName = function () { return this.name; };
FakeSheet.prototype.getLastRow = function () { return this.data.length; };
FakeSheet.prototype.getLastColumn = function () { return this.data.reduce(function (m, r) { return Math.max(m, r.length); }, 0); };
FakeSheet.prototype.getMaxRows = function () { return Math.max(this.data.length, 100); };
FakeSheet.prototype.setFrozenRows = function () {};
FakeSheet.prototype.setColumnWidth = function () {};
FakeSheet.prototype.appendRow = function (r) { this.data.push(r.slice()); };
FakeSheet.prototype.getRange = function (r, c, nr, nc) {
  var self = this; nr = nr || 1; nc = nc || 1;
  var chain = { setFontWeight: function () { return chain; }, setNumberFormat: function () { return chain; }, setDataValidation: function () { return chain; }, insertCheckboxes: function () { return chain; }, setWrap: function () { return chain; }, setBackgrounds: function () { return chain; }, clear: function () { for (var i = 0; i < nr; i++) if (self.data[r - 1 + i]) for (var j = 0; j < nc; j++) self.data[r - 1 + i][c - 1 + j] = ''; return chain; } };
  chain.getValues = function () {
    var out = [];
    for (var i = 0; i < nr; i++) { var row = self.data[r - 1 + i] || []; var line = []; for (var j = 0; j < nc; j++) line.push(row[c - 1 + j] === undefined ? '' : row[c - 1 + j]); out.push(line); }
    return out;
  };
  chain.setValues = function (vals) {
    for (var i = 0; i < vals.length; i++) { while (self.data.length < r + i) self.data.push([]); var row = self.data[r - 1 + i]; for (var j = 0; j < vals[i].length; j++) row[c - 1 + j] = vals[i][j]; }
    return chain;
  };
  return chain;
};
function FakeSpreadsheet(sheets) { this.sheets = sheets; }
FakeSpreadsheet.prototype.getSheetByName = function (n) { return this.sheets[n] || null; };
FakeSpreadsheet.prototype.insertSheet = function (n) { return (this.sheets[n] = new FakeSheet(n)); };
FakeSpreadsheet.prototype.toast = function () {};
function fakeProps() { var store = {}; return { getProperty: function (k) { return k in store ? store[k] : null; }, setProperty: function (k, v) { store[k] = String(v); }, deleteProperty: function (k) { delete store[k]; }, _store: store }; }
function withEnv(env, fn) {
  var saved = {}; Object.keys(env).forEach(function (k) { saved[k] = ctx[k]; ctx[k] = env[k]; });
  try { return fn(); } finally { Object.keys(saved).forEach(function (k) { ctx[k] = saved[k]; }); }
}
/** A workbook with the CRM tabs, Settings rows, and a Log that records to `logs`. */
function fakeWorkbook(settingsRows) {
  var sheets = {
    Master: new FakeSheet('Master', [ctx.CRM_MASTER_COLUMNS.slice()]),
    Interactions: new FakeSheet('Interactions', [ctx.CRM_INTERACTION_COLUMNS.slice()]),
    Import: new FakeSheet('Import', [ctx.CRM_IMPORT_USER_COLUMNS.concat(ctx.CRM_IMPORT_MACHINE_COLUMNS)]),
    Settings: new FakeSheet('Settings', [['Setting', 'Value', 'Notes']].concat(settingsRows || [])),
    Log: new FakeSheet('Log', [['Timestamp', 'Event', 'Target', 'Details']])
  };
  sheets.Log.appendRow = function (r) { logs.push(r); };
  return new FakeSpreadsheet(sheets);
}
function tableOf(sheet) { // read a FakeSheet as {header→[values]} rows for assertions
  var h = sheet.data[0]; return sheet.data.slice(1).map(function (r) { var o = {}; h.forEach(function (k, i) { o[k] = r[i] === undefined ? '' : r[i]; }); return o; });
}
function scanCtx(overrides) {
  var settings = Object.assign({ MyDomains: [], IgnoreDomains: [], SnippetChars: 400 }, overrides || {});
  return ctx.crmScanContext_(settings);
}
function freshTables() {
  return { master: makeTable(ctx.CRM_MASTER_COLUMNS), inter: makeTable(ctx.CRM_INTERACTION_COLUMNS) };
}

// ================================================================ pure helpers
test('normalize gmail dots/plus', function () {
  assert.strictEqual(ctx.crmNormalizeEmail_('  Jane.Doe+vc@GMail.com '), 'janedoe@gmail.com');
  assert.strictEqual(ctx.crmNormalizeEmail_('jane.doe@googlemail.com'), 'janedoe@gmail.com');
  assert.strictEqual(ctx.crmNormalizeEmail_('Jane.Doe@Acme.io'), 'jane.doe@acme.io');
  assert.strictEqual(ctx.crmNormalizeEmail_('not an email'), '');
});

test('parse address lists', function () {
  var r = ctx.crmParseAddresses_('"Doe, Jane" <jane@acme.io>, bob@x.com, Sam Q <SAM@y.org>');
  eq(r.map(function (a) { return a.email; }), ['jane@acme.io', 'bob@x.com', 'sam@y.org']);
  assert.strictEqual(r[0].name, 'Doe, Jane');
  assert.strictEqual(r[1].name, '');
  assert.strictEqual(r[2].name, 'Sam Q');
});

test('split names / normalize names / normalize companies', function () {
  eq(ctx.crmSplitName_('Jane Q. Doe'), { first: 'Jane', last: 'Doe' });
  eq(ctx.crmSplitName_('Doe, Jane'), { first: 'Jane', last: 'Doe' });
  eq(ctx.crmSplitName_('Cher'), { first: 'Cher', last: '' });
  assert.strictEqual(ctx.crmNormalizeName_('Dr. Jane Q. van der Doe (she/her)'), 'jane doe');
  assert.strictEqual(ctx.crmNormalizeName_('José Álvarez'), 'jose alvarez');
  assert.strictEqual(ctx.crmNormalizeName_('jane'), 'jane');
  assert.strictEqual(ctx.crmNormalizeCompany_('The Acme Company, Inc.'), 'acme');
  assert.strictEqual(ctx.crmNormalizeCompany_('Sequoia Capital'), 'sequoia');
});

test('machine address detection', function () {
  assert.ok(ctx.crmIsMachineAddress_('noreply@stripe.com'));
  assert.ok(ctx.crmIsMachineAddress_('calendar-notification@google.com'));
  assert.ok(ctx.crmIsMachineAddress_('notifications@github.com'));
  assert.ok(ctx.crmIsMachineAddress_('reply+abcdef1234567890@reply.example.com'));
  assert.ok(ctx.crmIsMachineAddress_('jane@bank.com', ['bank.com']));
  assert.ok(!ctx.crmIsMachineAddress_('jane@a16z.com'));
  assert.ok(!ctx.crmIsMachineAddress_('hello@startup.co'));
  assert.ok(!ctx.crmIsMachineAddress_('support@customer.com'));
});

test('body split: quotes stripped, signature captured', function () {
  var body = 'Hi Jane,\n\nThanks for the intro — happy to chat Tuesday.\n\nBest,\nDavid\n--\nDavid K | CEO, Startup\n+1 555 010 0000\n\nOn Mon, Jan 1, 2026 at 9:00 AM Jane <jane@acme.io>\nwrote:\n> Sure, let me know\n> when works.';
  var s = ctx.crmSplitBody_(body);
  assert.ok(s.body.indexOf('wrote:') === -1, 'quoted header removed (even when wrapped)');
  assert.ok(s.body.indexOf('let me know') === -1, 'quoted lines removed');
  assert.ok(s.body.indexOf('CEO') === -1, 'signature not in body');
  assert.ok(s.body.indexOf('happy to chat Tuesday') !== -1);
  assert.ok(s.signature.indexOf('CEO, Startup') !== -1, 'signature captured: ' + s.signature);
  // Heuristic sign-off without "--"
  var s2 = ctx.crmSplitBody_('Sounds good, see you then.\n\nThanks,\nPriya Patel\nPartner, Foo Ventures\nwww.foo.vc');
  assert.strictEqual(s2.body, 'Sounds good, see you then.');
  assert.ok(/Partner, Foo Ventures/.test(s2.signature));
  // Outlook header block
  var s3 = ctx.crmSplitBody_('Yes works for me.\n\nFrom: David <me@startup.io>\nSent: Monday\nTo: Jane\nSubject: Re: intro\n\nold text');
  assert.strictEqual(s3.body, 'Yes works for me.');
  assert.strictEqual(ctx.crmExcerpt_(body, 0), '');
  assert.ok(ctx.crmExcerpt_('word '.repeat(200), 50).length <= 52);
});

test('auto-reply detection', function () {
  assert.ok(ctx.crmIsAutoReply_('Automatic reply: Intro', {}));
  assert.ok(ctx.crmIsAutoReply_('Re: Out of Office', {}));
  assert.ok(ctx.crmIsAutoReply_('Delivery Status Notification (Failure)', {}));
  assert.ok(ctx.crmIsAutoReply_('Re: intro', { 'Auto-Submitted': 'auto-replied' }));
  assert.ok(ctx.crmIsAutoReply_('Re: intro', { 'Precedence': 'bulk' }));
  assert.ok(!ctx.crmIsAutoReply_('Re: intro', { 'Auto-Submitted': 'no' }));
  assert.ok(!ctx.crmIsAutoReply_('Re: intro — our office hours', {}));
});

test('opt-out detection is conservative', function () {
  assert.ok(ctx.crmIsOptOut_('Please remove me from your list.'));
  assert.ok(ctx.crmIsOptOut_('Unsubscribe'));
  assert.ok(ctx.crmIsOptOut_("We're no longer interested, thanks."));
  assert.ok(!ctx.crmIsOptOut_('Interesting — not sure it fits our thesis but happy to chat.'));
  assert.ok(!ctx.crmIsOptOut_('Can you remove me from the cc and add my colleague?'));
});

test('rule matching', function () {
  var m = function (rule, email, text) { return ctx.crmRuleMatches_(rule, { email: email, domain: ctx.crmDomainOf_(email), text: text || '' }); };
  assert.ok(m({ type: 'domain', pattern: '.vc' }, 'x@foo.vc'));
  assert.ok(!m({ type: 'domain', pattern: '.vc' }, 'x@foo.vcx.com'));
  assert.ok(m({ type: 'domain', pattern: 'a16z.com' }, 'x@mail.a16z.com'));
  assert.ok(!m({ type: 'domain', pattern: 'a16z.com' }, 'x@nota16z.com'));
  assert.ok(m({ type: 'email', pattern: 'Notifications@GitHub.com' }, 'notifications@github.com'));
  assert.ok(m({ type: 'keyword', pattern: 'term sheet' }, 'x@y.com', 'Re: Term Sheet draft'));
});

test('stage computation', function () {
  var S = ctx.crmComputeStage_;
  var base = { bounced: false, doNotContact: false, relationship: '', sent: 0, received: 0, daysSinceTouch: 1, currentStage: '' };
  var w = function (o) { var c = Object.assign({}, base); Object.keys(o).forEach(function (k) { c[k] = o[k]; }); return c; };
  assert.strictEqual(S(w({}), 45), 'New');
  assert.strictEqual(S(w({ sent: 1 }), 45), 'Contacted');
  assert.strictEqual(S(w({ sent: 1, received: 1 }), 45), 'Replied');
  assert.strictEqual(S(w({ sent: 3, received: 2 }), 45), 'Engaged');
  assert.strictEqual(S(w({ sent: 1, received: 1, relationship: 'engaged' }), 45), 'Engaged');
  assert.strictEqual(S(w({ sent: 1, received: 1, relationship: 'meeting_scheduled' }), 45), 'Meeting');
  assert.strictEqual(S(w({ sent: 1, received: 1, relationship: 'customer' }), 45), 'Won');
  assert.strictEqual(S(w({ sent: 1, received: 1, relationship: 'not_interested' }), 45), 'Not Interested');
  assert.strictEqual(S(w({ sent: 1, received: 1, daysSinceTouch: 90 }), 45), 'Dormant');
  assert.strictEqual(S(w({ sent: 1, daysSinceTouch: 90 }), 45), 'Contacted', 'never-replied contacts do not go dormant');
  assert.strictEqual(S(w({ sent: 1, received: 1, bounced: true }), 45), 'Bounced');
  assert.strictEqual(S(w({ sent: 1, doNotContact: true }), 45), 'Not Interested');
  assert.strictEqual(S(w({ sent: 1, currentStage: 'Won' }), 45), 'Won', 'hand-set Won survives');
  // Idempotent: applying twice gives the same answer.
  var f = w({ sent: 2, received: 1, daysSinceTouch: 3 });
  var once = S(f, 45); f.currentStage = once;
  assert.strictEqual(S(f, 45), once);
});

// ================================================================ scan (end to end through crmProcessThread_)
test('scan: outbound + reply aggregates, NeedsReply, signature, interactions', function () {
  var t = freshTables(), c = scanCtx();
  var thread = fakeThread('T1', 'Intro: Startup <> Acme', [
    { from: 'David <me@startup.io>', to: 'Jane Doe <jane@acme.io>', cc: 'cofounder@startup.io', date: '2026-08-01T10:00:00Z', body: 'Hi Jane, quick intro to what we do.\n\nBest,\nDavid' },
    { from: 'Jane Doe <jane@acme.io>', to: 'me@startup.io', date: '2026-08-02T10:00:00Z', subject: 'Re: Intro', body: 'Thanks David — interesting. Can you send the deck?\n\nJane Doe\nHead of Data, Acme\n+1 415 555 0100\n\nOn Sat David wrote:\n> Hi Jane' }
  ]);
  var touched = ctx.crmProcessThread_(thread, t.master, t.inter, c);
  assert.strictEqual(touched, 1, 'cofounder (own domain) is not a contact');
  var row = rowFor(t.master, 'jane@acme.io');
  assert.ok(row, 'master row created');
  assert.strictEqual(get(t.master, row, 'Name'), 'Jane Doe');
  assert.strictEqual(get(t.master, row, 'FirstName'), 'Jane');
  assert.strictEqual(get(t.master, row, 'SentCount'), 1);
  assert.strictEqual(get(t.master, row, 'ReceivedCount'), 1);
  assert.strictEqual(get(t.master, row, 'NeedsReply'), true, 'she spoke last');
  assert.ok(/Head of Data, Acme/.test(get(t.master, row, 'Signature')), 'signature captured from inbound');
  assert.ok(/send the deck/.test(get(t.master, row, 'LastInboundSnippet')));
  assert.ok(get(t.master, row, 'LastInboundSnippet').indexOf('Head of Data') === -1, 'signature not in excerpt');
  assert.strictEqual(get(t.master, row, 'ThreadIds'), 'T1');
  var ir = rowFor(t.inter, 'T1');
  assert.strictEqual(get(t.inter, ir, 'Outbound'), 1);
  assert.strictEqual(get(t.inter, ir, 'Inbound'), 1);
  assert.strictEqual(get(t.inter, ir, 'LastDirection'), 'IN');
  eq(JSON.parse(get(t.inter, ir, 'Participants')), { 'jane@acme.io': [1, 1, 0] });
});

test('scan: re-processing the same thread never double counts', function () {
  var t = freshTables(), c = scanCtx();
  var msgs = [
    { from: 'me@startup.io', to: 'jane@acme.io', date: '2026-08-01T10:00:00Z', body: 'Hi' },
    { from: 'jane@acme.io', to: 'me@startup.io', date: '2026-08-02T10:00:00Z', body: 'Hello back' }
  ];
  ctx.crmProcessThread_(fakeThread('T1', 'Intro', msgs), t.master, t.inter, c);
  ctx.crmProcessThread_(fakeThread('T1', 'Intro', msgs), t.master, t.inter, c);
  var row = rowFor(t.master, 'jane@acme.io');
  assert.strictEqual(get(t.master, row, 'SentCount'), 1);
  assert.strictEqual(get(t.master, row, 'ReceivedCount'), 1);
  assert.strictEqual(get(t.master, row, 'ThreadCount'), 1);
  assert.strictEqual(t.inter.rows.length, 1, 'one interactions row');
  // Thread grows: counts follow.
  msgs.push({ from: 'me@startup.io', to: 'jane@acme.io', date: '2026-08-03T10:00:00Z', body: 'Deck attached' });
  ctx.crmProcessThread_(fakeThread('T1', 'Intro', msgs), t.master, t.inter, c);
  assert.strictEqual(get(t.master, row, 'SentCount'), 2);
  assert.strictEqual(get(t.master, row, 'NeedsReply'), false, 'we spoke last now');
  // Second thread with the same contact adds up.
  ctx.crmProcessThread_(fakeThread('T2', 'Follow-up', [{ from: 'me@startup.io', to: 'jane@acme.io', date: '2026-08-10T10:00:00Z', body: 'Ping' }]), t.master, t.inter, c);
  assert.strictEqual(get(t.master, row, 'SentCount'), 3);
  assert.strictEqual(get(t.master, row, 'ThreadCount'), 2);
});

test('scan: auto-replies are not replies; bounces flagged; opt-out honoured', function () {
  var t = freshTables(), c = scanCtx();
  ctx.crmProcessThread_(fakeThread('T1', 'Intro', [
    { from: 'me@startup.io', to: 'bob@corp.com', date: '2026-08-01T10:00:00Z', body: 'Hi Bob' },
    { from: 'bob@corp.com', to: 'me@startup.io', date: '2026-08-01T10:01:00Z', subject: 'Automatic reply: Intro', body: 'I am out of the office until Monday.' }
  ]), t.master, t.inter, c);
  var bob = rowFor(t.master, 'bob@corp.com');
  assert.strictEqual(get(t.master, bob, 'ReceivedCount'), 0, 'OOO not counted as a reply');
  assert.strictEqual(get(t.master, bob, 'AutoReplies'), 1);
  assert.strictEqual(get(t.master, bob, 'NeedsReply'), false, 'an OOO does not need a reply');
  assert.strictEqual(get(t.inter, rowFor(t.inter, 'T1'), 'LastDirection'), 'AUTO');

  ctx.crmProcessThread_(fakeThread('T2', 'Intro', [
    { from: 'me@startup.io', to: 'gone@corp.com', date: '2026-08-01T10:00:00Z', body: 'Hi' },
    { from: 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>', to: 'me@startup.io', date: '2026-08-01T10:01:00Z', subject: 'Delivery Status Notification (Failure)', body: 'Address not found: gone@corp.com' }
  ]), t.master, t.inter, c);
  assert.strictEqual(get(t.master, rowFor(t.master, 'gone@corp.com'), 'Bounced'), true);
  assert.ok(!t.master.index['mailer-daemon@googlemail.com'], 'daemon is not a contact');

  ctx.crmProcessThread_(fakeThread('T3', 'Intro', [
    { from: 'me@startup.io', to: 'nope@corp.com', date: '2026-08-01T10:00:00Z', body: 'Hi' },
    { from: 'nope@corp.com', to: 'me@startup.io', date: '2026-08-02T10:00:00Z', body: 'Please remove me from your list.' }
  ]), t.master, t.inter, c);
  var nope = rowFor(t.master, 'nope@corp.com');
  assert.strictEqual(get(t.master, nope, 'DoNotContact'), true);
  assert.ok(/optout/.test(get(t.master, nope, 'Tags')));
});

test('scan: MyDomains, ignore rules, machine recipients, multi-recipient threads', function () {
  var t = freshTables(), c = scanCtx({ MyDomains: ['ourbrand.com', 'advisor@gmail.com'] });
  var touched = ctx.crmProcessThread_(fakeThread('T1', 'Group intro', [
    { from: 'me@startup.io', to: 'a@x.com, b@y.com, teammate@ourbrand.com, advisor@gmail.com, noreply@calendly.com', date: '2026-08-01T10:00:00Z', body: 'Hi all' },
    { from: 'b@y.com', to: 'me@startup.io', date: '2026-08-02T10:00:00Z', body: 'Thanks!' }
  ]), t.master, t.inter, c);
  assert.strictEqual(touched, 2, 'only a@x.com and b@y.com');
  assert.strictEqual(get(t.master, rowFor(t.master, 'a@x.com'), 'NeedsReply'), false, 'b replied, not a');
  assert.strictEqual(get(t.master, rowFor(t.master, 'b@y.com'), 'NeedsReply'), true);
  assert.strictEqual(get(t.inter, rowFor(t.inter, 'T1'), 'Email'), 'b@y.com', 'primary counterpart = most active');
});

// ================================================================ classifier (through crmClassifyBatch_)
var settings = { Model: 'claude-opus-5', Categories: ctx.CRM_DEFAULT_CATEGORIES.slice(), CompanyDescription: 'We sell X to Y.', SnippetChars: 400, MinConfidence: 0.7, RetryLowConfidence: true, EvidenceThreads: 3, ClassifyBatchSize: 8 };
function apiOk(contacts, extra) {
  return { code: 200, body: Object.assign({ model: 'claude-opus-5', stop_reason: 'end_turn', content: [{ type: 'text', text: JSON.stringify({ contacts: contacts }) }], usage: { input_tokens: 100, output_tokens: 50 } }, extra || {}) };
}
function result(o) {
  return Object.assign({ email: 'jane@acme.io', category: 'Investor', confidence: 0.9, evidence: 'signature says Partner at Acme Ventures', company: 'Acme Ventures', title: 'Partner', first_name: 'Jane', last_name: 'Doe', relationship: 'replied', sentiment: 'positive', summary: 'Investor who replied.', next_action: 'Send deck' }, o);
}

test('classifier: request shape and validated parsing', function () {
  fetchQueue = [apiOk([result({}), { email: 'stranger@nowhere.com', category: 'Other', confidence: 1 }])]; fetchCalls = []; logs = [];
  var out = ctx.crmClassifyBatch_([{ email: 'jane@acme.io' }], settings, 'sk-test', 'low');
  var req = fetchCalls[0];
  assert.strictEqual(req.url, 'https://api.anthropic.com/v1/messages');
  assert.strictEqual(req.options.headers['x-api-key'], 'sk-test');
  assert.strictEqual(req.options.headers['anthropic-beta'], 'server-side-fallback-2026-07-01');
  assert.strictEqual(req.body.model, 'claude-opus-5');
  assert.strictEqual(req.body.fallbacks, 'default');
  assert.strictEqual(req.body.output_config.effort, 'low');
  assert.strictEqual(req.body.output_config.format.type, 'json_schema');
  assert.ok(/We sell X to Y/.test(req.body.system));
  assert.strictEqual(out.length, 1, 'unexpected email dropped');
  assert.strictEqual(out[0].category, 'Investor');
  assert.ok(logs.some(function (l) { return l[1] === 'CLASSIFY_UNEXPECTED'; }));
});

test('classifier: coercion clamps and defaults', function () {
  fetchQueue = [apiOk([result({ category: 'Unicorn', confidence: 7, relationship: 'bff', sentiment: 'meh', evidence: null })])];
  var out = ctx.crmClassifyBatch_([{ email: 'jane@acme.io' }], settings, 'k', 'low')[0];
  assert.strictEqual(out.category, 'Other');
  assert.strictEqual(out.confidence, 1);
  assert.strictEqual(out.relationship, 'unclear');
  assert.strictEqual(out.sentiment, 'neutral');
  assert.strictEqual(out.evidence, '');
  assert.ok(ctx.crmNeedsReview_(out, settings), 'coerced-to-Other is flagged for review');
  assert.ok(!ctx.crmNeedsReview_(ctx.crmCoerceResult_(result({}), settings), settings));
  assert.ok(ctx.crmNeedsReview_(ctx.crmCoerceResult_(result({ confidence: 0.55 }), settings), settings));
});

test('classifier: retries on 529 then succeeds; gives up after 4 attempts', function () {
  fetchQueue = [{ code: 529, body: 'overloaded' }, { code: 500, body: 'boom' }, apiOk([result({})])]; fetchCalls = [];
  var out = ctx.crmClassifyBatch_([{ email: 'jane@acme.io' }], settings, 'k', 'low');
  assert.strictEqual(out.length, 1);
  assert.strictEqual(fetchCalls.length, 3);
  fetchQueue = [{ code: 429, body: 'x' }, { code: 429, body: 'x' }, { code: 429, body: 'x' }, { code: 429, body: 'x' }];
  assert.throws(function () { ctx.crmClassifyBatch_([{ email: 'jane@acme.io' }], settings, 'k', 'low'); }, /429 after retries/);
});

test('classifier: refusal, truncation, 401, bad JSON all throw cleanly', function () {
  fetchQueue = [apiOk([], { stop_reason: 'refusal', stop_details: { type: 'refusal', category: 'x' } })];
  assert.throws(function () { ctx.crmClassifyBatch_([{ email: 'a@b.co' }], settings, 'k', 'low'); }, /refusal/);
  fetchQueue = [apiOk([], { stop_reason: 'max_tokens' })];
  assert.throws(function () { ctx.crmClassifyBatch_([{ email: 'a@b.co' }], settings, 'k', 'low'); }, /max_tokens/);
  fetchQueue = [{ code: 401, body: '{"error":{"message":"invalid x-api-key"}}' }];
  assert.throws(function () { ctx.crmClassifyBatch_([{ email: 'a@b.co' }], settings, 'k', 'low'); }, /401/);
  fetchQueue = [{ code: 200, body: { stop_reason: 'end_turn', content: [{ type: 'text', text: 'not json' }] } }];
  assert.throws(function () { ctx.crmClassifyBatch_([{ email: 'a@b.co' }], settings, 'k', 'low'); }, /not valid JSON/);
});

test('classifier: evidence pack includes signature and recent threads', function () {
  var t = freshTables(), c = scanCtx();
  ctx.crmProcessThread_(fakeThread('T1', 'Intro', [
    { from: 'me@startup.io', to: 'jane@acme.io', date: '2026-08-01T10:00:00Z', body: 'Hi Jane' },
    { from: 'jane@acme.io', to: 'me@startup.io', date: '2026-08-02T10:00:00Z', body: 'Send the deck.\n\nJane\nPartner, Acme Ventures' }
  ]), t.master, t.inter, c);
  ctx.crmProcessThread_(fakeThread('T2', 'Deck', [
    { from: 'me@startup.io', to: 'jane@acme.io', date: '2026-08-05T10:00:00Z', body: 'Deck attached.' }
  ]), t.master, t.inter, c);
  var p = ctx.crmContactPayload_(t.master, rowFor(t.master, 'jane@acme.io'), t.inter, settings);
  assert.strictEqual(p.emails_we_sent, 2);
  assert.strictEqual(p.emails_they_sent, 1);
  assert.ok(/Partner, Acme Ventures/.test(p.their_signature));
  assert.strictEqual(p.threads.length, 2);
  assert.strictEqual(p.threads[0].subject, 'Deck', 'newest first');
  assert.strictEqual(p.threads[1].their_latest_message, 'Send the deck.');
  assert.strictEqual(p.our_latest_message, undefined, 'aggregate fallback only when threads missing');
  var p0 = ctx.crmContactPayload_(t.master, rowFor(t.master, 'jane@acme.io'), t.inter, Object.assign({}, settings, { SnippetChars: 0 }));
  assert.strictEqual(p0.threads, undefined, 'SnippetChars 0 sends metadata only');
  assert.strictEqual(p0.their_signature, undefined);
});

test('classifier: apply result respects existing values and sets review flag', function () {
  var t = freshTables();
  var row = addRow(t.master, { Email: 'jane@acme.io', Company: 'Acme (hand-entered)', Tags: 'warm' }, 'jane@acme.io');
  ctx.crmApplyAiResult_(t.master, row, ctx.crmCoerceResult_(result({ confidence: 0.6 }), settings), settings);
  assert.strictEqual(get(t.master, row, 'Company'), 'Acme (hand-entered)', 'never overwrite a filled company');
  assert.strictEqual(get(t.master, row, 'Title'), 'Partner');
  assert.strictEqual(get(t.master, row, 'Name'), 'Jane Doe');
  assert.strictEqual(get(t.master, row, 'ReviewNeeded'), true);
  assert.strictEqual(get(t.master, row, 'Tags'), 'warm, rel:replied');
  assert.strictEqual(get(t.master, row, 'ClassifiedBy'), 'AI');
});

test('classifier: candidate selection honours Lock, MANUAL, mode, and same-job stamps', function () {
  var t = freshTables();
  addRow(t.master, { Email: 'a@x.com' });
  addRow(t.master, { Email: 'b@x.com', Category: 'Investor' });
  addRow(t.master, { Email: 'c@x.com', Lock: true });
  addRow(t.master, { Email: 'd@x.com', Category: 'Other', ClassifiedBy: 'MANUAL' });
  addRow(t.master, { Email: 'e@x.com', ClassifiedAt: new CtxDate(Date.now() + 1000) });
  addRow(t.master, { Email: '' });
  eq(ctx.crmClassifyCandidates_(t.master, { mode: 'unclassified', startedAt: Date.now() }), [0]);
  eq(ctx.crmClassifyCandidates_(t.master, { mode: 'all', startedAt: Date.now() }), [0, 1]);
});

// ================================================================ dedupe
test('dedupe decisions', function () {
  var t = freshTables(); var master = t.master;
  var old = new CtxDate(Date.now() - 120 * 86400000), recent = new CtxDate(Date.now() - 3 * 86400000);
  addRow(master, { Email: 'jane@acme.io', Name: 'Jane Doe', Company: 'Acme, Inc.', Stage: 'Engaged', LastTouchAt: recent, Category: 'Investor' }, 'jane@acme.io');
  addRow(master, { Email: 'old@acme.io', Name: 'Old Timer', Stage: 'Contacted', LastTouchAt: old }, 'old@acme.io');
  addRow(master, { Email: 'fresh@beta.com', Name: 'Fresh', Stage: 'Contacted', LastTouchAt: recent }, 'fresh@beta.com');
  addRow(master, { Email: 'gone@beta.com', Name: 'Gone', Stage: 'Contacted', Bounced: true, LastTouchAt: old }, 'gone@beta.com');
  addRow(master, { Email: 'p@gmail.com', Name: 'Pat Lee', Company: 'Gamma Labs', Stage: 'Engaged', LastTouchAt: recent }, 'p@gmail.com');
  var idx = ctx.crmBuildDedupeIndexes_(master);
  idx.queued['queued@delta.com'] = true;
  var s = { ImportRecentDays: 30 };
  var D = function (email, person, seen) { return ctx.crmDedupeRow_(ctx.crmNormalizeEmail_(email), seen || {}, master, idx, s, person || {}); };

  assert.strictEqual(D('bad').status, 'INVALID');
  assert.strictEqual(D('JANE@acme.io').status, 'DUPLICATE');
  assert.strictEqual(D('jane@acme.io').recommendation, 'Skip — already in conversation');
  assert.strictEqual(D('old@acme.io').recommendation, 'Re-engage — reference prior thread');
  assert.ok(/^Skip — emailed 3 day/.test(D('fresh@beta.com').recommendation));
  assert.strictEqual(D('gone@beta.com').recommendation, 'Skip — do not contact');
  assert.strictEqual(D('queued@delta.com').status, 'QUEUED');

  var sc = D('new@acme.io', { name: 'Someone Else' });
  assert.strictEqual(sc.status, 'SAME_COMPANY');
  assert.strictEqual(sc.matchedName, 'Jane Doe', 'most advanced contact at the domain');
  assert.ok(/in conversation with Jane Doe/.test(sc.recommendation));

  var pm = D('jane.doe@gmail.com', { name: 'Dr Jane Doe', company: 'The Acme Company' });
  assert.strictEqual(pm.status, 'POSSIBLE_MATCH', 'same name + same company, other email');
  assert.ok(/likely the same person as jane@acme.io/.test(pm.recommendation));
  var pm2 = D('jdoe@acme.io', { name: 'Jane Doe' });
  assert.strictEqual(pm2.status, 'POSSIBLE_MATCH', 'same name + same domain beats SAME_COMPANY');
  assert.strictEqual(D('jane.doe@other.com', { name: 'Jane Doe', company: 'Different Co' }).status, 'NEW', 'same name alone is not a match');
  assert.strictEqual(D('pat@gammalabs.io', { name: 'Pat Lee', company: 'Gamma Labs Inc' }).status, 'POSSIBLE_MATCH', 'freemail contact matched by name+company');

  assert.strictEqual(D('someone@gmail.com').status, 'NEW', 'freemail domains never match SAME_COMPANY');
  assert.strictEqual(D('x@nowhere.org').status, 'NEW');
  assert.strictEqual(D('x@nowhere.org', {}, { 'x@nowhere.org': true }).status, 'DUPLICATE_IN_IMPORT');
});

test('import header aliases', function () {
  assert.strictEqual(ctx.crmCanonicalImportHeader_('e-mail address'), 'Email');
  assert.strictEqual(ctx.crmCanonicalImportHeader_('organization'), 'Company');
  assert.strictEqual(ctx.crmCanonicalImportHeader_('linkedin'), null);
});

// ================================================================ audit
test('audit finds invariant violations', function () {
  var t = freshTables();
  addRow(t.master, { Email: 'a@x.com', Category: 'Investor', Stage: 'Contacted', SentCount: 1, ClassifiedBy: 'AI', Confidence: 0.9, Evidence: 'ok', ThreadIds: 'T1' });
  addRow(t.master, { Email: 'A@x.com', Category: 'Investor', Stage: 'Contacted', SentCount: 1, ThreadIds: 'T1' }); // dup
  addRow(t.master, { Email: 'b@x.com', Category: 'Customer', Stage: 'Contacted', SentCount: 2, ReceivedCount: 2, ClassifiedBy: 'AI', Confidence: 0.4, Evidence: 'x', ThreadIds: 'T2', LastInboundSnippet: 'hi', Source: 'Scan' }); // stage wrong + low conf not flagged
  addRow(t.master, { Email: 'c@x.com', Stage: 'New' }); // unclassified
  addRow(t.master, { Email: 'd@x.com', Category: 'Other', Stage: 'Contacted', SentCount: 1, ThreadIds: 'T3', Tags: 'optout', DoNotContact: false, NeedsReply: true, Bounced: true });
  var rep = ctx.crmAuditMaster_(t.master, Object.assign({}, settings, { DormantDays: 45 }));
  var by = {}; rep.findings.forEach(function (f) { by[f.check] = f.count; });
  assert.strictEqual(by['Duplicate emails after normalization'], 1);
  assert.strictEqual(by['Unclassified (no Category)'], 1);
  assert.strictEqual(by['Confidence below MinConfidence but not flagged'], 1);
  assert.ok(by['Stage disagrees with facts'] >= 2, 'b (Engaged expected) and d (Bounced expected)');
  assert.strictEqual(by['NeedsReply on a suppressed contact'], 1);
  assert.strictEqual(by['Opted out but still contactable'], 1);
  assert.strictEqual(rep.total, 5);
  assert.ok(rep.problems > 0);
  var clean = freshTables();
  addRow(clean.master, { Email: 'ok@x.com', Category: 'Investor', Stage: 'Contacted', SentCount: 1, ClassifiedBy: 'RULE', Confidence: 1, Evidence: 'Rule domain:x.com', ThreadIds: 'T1' });
  assert.strictEqual(ctx.crmAuditMaster_(clean.master, settings).problems, 0);
});

test('misc helpers', function () {
  var tg = ctx.crmSetTag_('warm, rel:replied', 'rel:', 'engaged');
  assert.strictEqual(tg, 'warm, rel:engaged');
  assert.strictEqual(ctx.crmGetTag_(tg, 'rel:'), 'engaged');
  assert.strictEqual(ctx.crmAddToList_('a, b', 'b'), 'a, b');
  assert.strictEqual(ctx.crmAddToList_('a, b', 'c', 2), 'b, c');
  var sch = ctx.crmClassifySchema_(['Investor', 'Other']);
  var item = sch.properties.contacts.items;
  assert.strictEqual(item.additionalProperties, false);
  eq(Object.keys(item.properties).sort(), item.required.slice().sort());
  assert.ok(item.properties.category.enum.indexOf('Ignore') !== -1);
});


// ================================================================ review fixes
test('machine-address rules: humans at SaaS companies and initial.surname addresses are contacts', function () {
  ['jane@google.com', 'bob@stripe.com', 'sam@github.com', 'kim@slack.com', 'r.stevenson@corp.com', 'm-rodriguez@corp.com', 'b_williams@corp.com', 'reply@corp.com'].forEach(function (e) {
    assert.ok(!ctx.crmIsMachineAddress_(e), e + ' should be a contact');
  });
  ['reply+1a2b3c4d5e6f7890@reply.tool.com', 'msg-9f8e7d6c5b4a@notify.tool.com', 'drive-shares-dm-noreply@google.com', 'messages-noreply@linkedin.com', 'jobs-listings@linkedin.com', 'hello@calendly.com'].forEach(function (e) {
    assert.ok(ctx.crmIsMachineAddress_(e), e + ' should be machine');
  });
});

test('freemail covers regional providers; own freemail domain is never treated as the team', function () {
  ['yahoo.fr', 'hotmail.co.uk', 'outlook.de', 'live.jp', 'zoho.com', 'mail.com', 'yandex.com', 'gmx.at', 'web.de', 'qq.com'].forEach(function (d) { assert.ok(ctx.crmIsFreemail_(d), d); });
  assert.ok(!ctx.crmIsFreemail_('startup.io'));
  var saved = ctx.Session;
  ctx.Session = { getScriptTimeZone: function () { return 'UTC'; }, getEffectiveUser: function () { return { getEmail: function () { return 'founder@yahoo.fr'; } }; } };
  try {
    var c = scanCtx();
    assert.ok(!c.isMine('investor@yahoo.fr'), 'yahoo.fr contact kept when the user is on yahoo.fr');
    var c2 = ctx.crmScanContext_({ MyDomains: [], IgnoreDomains: [], SnippetChars: 400, TreatMyDomainAsTeam: false });
    assert.ok(!c2.isMine('x@startup.io'));
  } finally { ctx.Session = saved; }
  assert.ok(scanCtx().isMine('teammate@startup.io'), 'own company domain is the team by default');
  assert.ok(!scanCtx({ TreatMyDomainAsTeam: false }).isMine('teammate@startup.io'));
});

test('bounce detection matches dotted / tagged gmail recipients', function () {
  var t = freshTables(), c = scanCtx();
  ctx.crmProcessThread_(fakeThread('T1', 'Intro', [
    { from: 'me@startup.io', to: 'Jane <jane.doe+intro@gmail.com>', date: '2026-08-01T10:00:00Z', body: 'Hi' },
    { from: 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>', to: 'me@startup.io', date: '2026-08-01T10:01:00Z', subject: 'Delivery Status Notification (Failure)', body: 'The address jane.doe+intro@gmail.com was not found.' }
  ]), t.master, t.inter, c);
  assert.strictEqual(get(t.master, rowFor(t.master, 'janedoe@gmail.com'), 'Bounced'), true);
});

test('scan: unchanged threads are skipped without reading messages', function () {
  var t = freshTables(), c = scanCtx();
  var msgs = [{ from: 'me@startup.io', to: 'jane@acme.io', date: '2026-08-01T10:00:00Z', body: 'Hi' }];
  var th = fakeThread('T1', 'Intro', msgs);
  assert.strictEqual(ctx.crmProcessThread_(th, t.master, t.inter, c), 1);
  var reads = 0; th.getMessages = function () { reads++; return msgs.map(fakeMessage); };
  assert.strictEqual(ctx.crmProcessThread_(th, t.master, t.inter, c), null, 'skipped');
  assert.strictEqual(reads, 0, 'getMessages not called for an unchanged thread');
  msgs.push({ from: 'jane@acme.io', to: 'me@startup.io', date: '2026-08-02T10:00:00Z', body: 'Hello' });
  assert.strictEqual(ctx.crmProcessThread_(th, t.master, t.inter, c), 1, 'processed once it changed');
  assert.strictEqual(get(t.master, rowFor(t.master, 'jane@acme.io'), 'ReceivedCount'), 1);
});

test('save merges by key: hand edits, sorts, deletions, and a Lock ticked mid-job all survive', function () {
  var H = ctx.CRM_MASTER_COLUMNS, col = function (h) { return H.indexOf(h); };
  var mk = function (email, cat, notes) { var r = H.map(function () { return ''; }); r[col('Email')] = email; r[col('Category')] = cat; r[col('Notes')] = notes || ''; return r; };
  var sheet = new FakeSheet('Master', [H.slice(), mk('a@x.com', 'Investor'), mk('b@x.com', ''), mk('c@x.com', 'Other')]);
  var wb = new FakeSpreadsheet({ Master: sheet });
  withEnv({ SpreadsheetApp: { getActive: function () { return wb; } } }, function () {
    var t = ctx.crmLoadTable_('Master', H, function (r, tt) { return ctx.crmNormalizeEmail_(ctx.crmGet_(tt, r, 'Email')); });
    // Job decides: b → Customer, c → Investor with new evidence; appends d.
    ctx.crmSet_(t, t.rows[1], 'Category', 'Customer'); ctx.crmSet_(t, t.rows[1], 'Evidence', 'said they use it');
    ctx.crmSet_(t, t.rows[2], 'Category', 'Investor');
    var d = ctx.crmAppendRow_(t, 'd@x.com'); ctx.crmSet_(t, d, 'Email', 'd@x.com'); ctx.crmSet_(t, d, 'Category', 'Press');
    // Meanwhile the user: sorts (c, a, b), writes a note on b, deletes a, locks c with a hand label.
    var rc = mk('c@x.com', 'Advisor'); rc[col('Lock')] = true;
    var rb = mk('b@x.com', '', 'met at conf');
    sheet.data = [H.slice(), rc, rb];
    ctx.crmSaveTable_(t);
    var rows = tableOf(sheet), by = {}; rows.forEach(function (r) { by[r.Email] = r; });
    assert.strictEqual(rows.length, 3, 'a stays deleted, d appended');
    assert.strictEqual(rows[0].Email, 'c@x.com', 'user order kept');
    assert.strictEqual(by['b@x.com'].Category, 'Customer', 'job value applied to the matching row');
    assert.strictEqual(by['b@x.com'].Evidence, 'said they use it');
    assert.strictEqual(by['b@x.com'].Notes, 'met at conf', 'hand edit made during the job survives');
    assert.strictEqual(by['c@x.com'].Category, 'Advisor', 'Lock ticked mid-job wins over the job');
    assert.strictEqual(by['c@x.com'].Lock, true);
    assert.strictEqual(by['d@x.com'].Category, 'Press');
    assert.ok(!t.dirty && !t.rows[1]._d, 'dirty flags cleared');
  });
});

test('commit guard: Stop or a newer Start during a tick wins', function () {
  var props = fakeProps();
  withEnv({ PropertiesService: { getScriptProperties: function () { return props; } } }, function () {
    var mine = { startedAt: 100, offset: 5 };
    props.setProperty('P', JSON.stringify({ startedAt: 100, offset: 0 }));
    assert.strictEqual(ctx.crmCommitState_('P', mine), true);
    assert.strictEqual(JSON.parse(props.getProperty('P')).offset, 5);
    props.deleteProperty('P');                                  // Stop clicked
    assert.strictEqual(ctx.crmCommitState_('P', mine), false);
    assert.strictEqual(props.getProperty('P'), null, 'stopped job not resurrected');
    props.setProperty('P', JSON.stringify({ startedAt: 200, offset: 0 })); // newer job started
    assert.strictEqual(ctx.crmCommitState_('P', mine), false);
    assert.strictEqual(JSON.parse(props.getProperty('P')).startedAt, 200, 'newer job untouched');
  });
});

test('scan windows: stable date-bounded queries, phase progress', function () {
  var day = 86400000, since = Date.UTC(2026, 0, 1), until = since + 10 * day;
  var st = { full: false, sinceMs: since, untilMs: until, windowDays: 7, phase: 1, cursorMs: since };
  var w = ctx.crmScanWindow_(st);
  assert.strictEqual(w.start, since); assert.strictEqual(w.end, since + 7 * day);
  assert.strictEqual(w.query, 'in:sent after:' + since / 1000 + ' before:' + (since + 7 * day) / 1000);
  st.cursorMs = w.end; w = ctx.crmScanWindow_(st);
  assert.strictEqual(w.end, until, 'last window clipped to until');
  st.cursorMs = until; assert.strictEqual(ctx.crmScanWindow_(st), null);
  st.phase = 2; st.cursorMs = since; assert.ok(/^-in:sent /.test(ctx.crmScanWindow_(st).query));
  assert.ok(Math.abs(ctx.crmScanProgress_({ full: false, sinceMs: since, untilMs: until, phase: 2, cursorMs: since }) - 0.5) < 1e-9);
  assert.strictEqual(ctx.crmScanProgress_({ full: true, sinceMs: since, untilMs: until, phase: 1, cursorMs: until }), 1);
});

test('scan driver end to end: windows, multi-tick resume, phase 2 replies, quick-skip, stop', function () {
  var day = 86400000;
  var T = function (iso) { return Date.parse(iso); };
  // Mailbox: 4 threads across 3 weeks. T3 has only an inbound reply in the incremental window.
  var threads = [
    fakeThread('T1', 'Intro A', [{ from: 'me@startup.io', to: 'a@x.com', date: '2026-07-02T10:00:00Z', body: 'Hi A' }]),
    fakeThread('T2', 'Intro B', [{ from: 'me@startup.io', to: 'b@y.com', date: '2026-07-09T10:00:00Z', body: 'Hi B' }, { from: 'b@y.com', to: 'me@startup.io', date: '2026-07-10T10:00:00Z', body: 'Thanks' }]),
    fakeThread('T3', 'Old thread', [{ from: 'me@startup.io', to: 'c@z.com', date: '2026-05-01T10:00:00Z', body: 'Hi C' }, { from: 'c@z.com', to: 'me@startup.io', date: '2026-07-15T10:00:00Z', body: 'Finally replying!' }]),
    fakeThread('T4', 'Intro D', [{ from: 'me@startup.io', to: 'd@w.com', date: '2026-07-16T10:00:00Z', body: 'Hi D' }])
  ];
  var searches = [];
  var gmail = { search: function (q, offset, max) {
    searches.push(q);
    var m = q.match(/after:(\d+) before:(\d+)/), lo = Number(m[1]) * 1000, hi = Number(m[2]) * 1000, phase2 = /^-in:sent/.test(q);
    var hits = threads.filter(function (th) {
      return th._messages.some(function (msg) { var d = T(msg.date), mine = /me@startup\.io/.test(msg.from); return d >= lo && d < hi && (phase2 ? !mine : mine); });
    });
    hits.sort(function (a, b) { return b.getLastMessageDate() - a.getLastMessageDate(); });
    return hits.slice(offset, offset + max);
  } };
  var props = fakeProps();
  var triggers = [];
  var scriptApp = { getProjectTriggers: function () { return triggers.slice(); }, deleteTrigger: function (t) { triggers = triggers.filter(function (x) { return x !== t; }); },
    newTrigger: function (fn) { return { timeBased: function () { return { everyMinutes: function () { return { create: function () { var t = { getHandlerFunction: function () { return fn; } }; triggers.push(t); return t; } }; } }; } }; } };
  var wb = fakeWorkbook([['ScanSince', '2026-07-01'], ['MaxThreadsPerRun', 1], ['ScanWindowDays', 7], ['AutoClassifyAfterScan', false], ['SnippetChars', 400]]);
  var lock = { tryLock: function () { return true; }, releaseLock: function () {} };
  var realNow = CtxDate.now;
  var now = T('2026-07-20T00:00:00Z');
  withEnv({ GmailApp: gmail, PropertiesService: { getScriptProperties: function () { return props; } }, ScriptApp: scriptApp, LockService: { getScriptLock: function () { return lock; } },
    SpreadsheetApp: { getActive: function () { return wb; }, newDataValidation: function () { var c = { requireValueInList: function () { return c; }, setAllowInvalid: function () { return c; }, build: function () { return {}; } }; return c; } } }, function () {
    props.setProperty('CRM_LAST_SCAN_AT', String(T('2026-07-08T00:00:00Z'))); // last scan → incremental since 2026-07-06
    CtxDate.now = function () { return now; };
    try {
      ctx.crmStartScan(false);                                 // runs the first tick itself
      var st = JSON.parse(props.getProperty('CRM_SCAN_STATE'));
      assert.strictEqual(st.full, false);
      assert.strictEqual(st.sinceMs, T('2026-07-06T00:00:00Z'), '2-day overlap before the last scan');
      assert.strictEqual(st.untilMs, now);
      var ticks = 1;
      while (props.getProperty('CRM_SCAN_STATE') && ticks < 30) { ctx.crmScanStep(); ticks++; }
      assert.ok(!props.getProperty('CRM_SCAN_STATE'), 'job finished');
      assert.ok(ticks > 2, 'MaxThreadsPerRun=1 forced several ticks (' + ticks + ')');
      assert.strictEqual(triggers.length, 0, 'trigger removed when done');
      assert.strictEqual(props.getProperty('CRM_LAST_SCAN_AT'), String(now));
      var rows = tableOf(wb.sheets.Master), by = {}; rows.forEach(function (r) { by[r.Email] = r; });
      assert.ok(!by['a@x.com'], 'T1 (2026-07-02) is before the incremental window');
      assert.strictEqual(by['b@y.com'].ReceivedCount, 1);
      assert.strictEqual(by['b@y.com'].Stage, 'Replied');
      assert.strictEqual(by['c@z.com'].ReceivedCount, 1, 'phase 2 picked up a reply on an old thread');
      assert.strictEqual(by['c@z.com'].NeedsReply, true);
      assert.strictEqual(by['d@w.com'].SentCount, 1);
      assert.ok(searches.some(function (q) { return /^-in:sent/.test(q); }), 'phase 2 ran');
      assert.ok(searches.every(function (q) { return /after:\d+ before:\d+$/.test(q); }), 'every query is date-bounded');
      var inter = tableOf(wb.sheets.Interactions);
      assert.strictEqual(inter.length, 3);

      // Second incremental run: everything unchanged → all skipped, counts identical.
      logs = [];
      ctx.crmStartScan(false);
      while (props.getProperty('CRM_SCAN_STATE')) ctx.crmScanStep();
      var done = logs.filter(function (l) { return l[1] === 'SCAN_DONE'; }).pop();
      assert.ok(/0 threads read, \d+ unchanged threads skipped/.test(done[3]), done[3]);
      var rows2 = tableOf(wb.sheets.Master), by2 = {}; rows2.forEach(function (r) { by2[r.Email] = r; });
      assert.strictEqual(by2['b@y.com'].ReceivedCount, 1); assert.strictEqual(by2['c@z.com'].SentCount, 1);

      // Stop during a job: the next tick must not resurrect it.
      ctx.crmStartScan(true);
      ctx.crmStopScan();
      assert.strictEqual(props.getProperty('CRM_SCAN_STATE'), null);
      ctx.crmScanStep();
      assert.strictEqual(props.getProperty('CRM_SCAN_STATE'), null);
      assert.strictEqual(triggers.length, 0);
    } finally { CtxDate.now = realNow; }
  });
});

console.log(passed + ' passed, ' + failed + ' failed');
