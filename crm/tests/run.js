#!/usr/bin/env node
/**
 * Smoke tests for the pure helpers in the CRM .gs files. Run: node crm/tests/run.js
 * Apps Script globals are stubbed just enough for the helpers under test.
 */
var fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');

var ctx = {
  console: console,
  Utilities: { formatDate: function (d) { return d.toISOString().slice(0, 10); } },
  Session: { getScriptTimeZone: function () { return 'UTC'; }, getEffectiveUser: function () { return { getEmail: function () { return 'me@startup.io'; } }; } },
  SpreadsheetApp: { getActive: function () { return { getSheetByName: function () { return null; } }; } },
  PropertiesService: { getScriptProperties: function () { return { getProperty: function () { return null; }, setProperty: function () {}, deleteProperty: function () {} }; } },
  ScriptApp: { getProjectTriggers: function () { return []; } },
  LockService: {}, GmailApp: {}, UrlFetchApp: {}, DriveApp: {}, HtmlService: {}, MimeType: {}
};
vm.createContext(ctx);
['Crm.gs', 'CrmScan.gs', 'CrmClassify.gs', 'CrmPipeline.gs', 'CrmImport.gs'].forEach(function (f) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), ctx, { filename: f });
});

var CtxDate = vm.runInContext('Date', ctx);
function eq(a, b, msg) { assert.strictEqual(JSON.stringify(a), JSON.stringify(b), msg); }
var passed = 0;
function test(name, fn) { try { fn(); passed++; } catch (e) { console.error('FAIL', name, '\n  ', e.message); process.exitCode = 1; } }

// --- email normalization
test('normalize gmail dots/plus', function () {
  assert.strictEqual(ctx.crmNormalizeEmail_('  Jane.Doe+vc@GMail.com '), 'janedoe@gmail.com');
  assert.strictEqual(ctx.crmNormalizeEmail_('jane.doe@googlemail.com'), 'janedoe@gmail.com');
  assert.strictEqual(ctx.crmNormalizeEmail_('Jane.Doe@Acme.io'), 'jane.doe@acme.io');
  assert.strictEqual(ctx.crmNormalizeEmail_('not an email'), '');
});

// --- address parsing
test('parse address lists', function () {
  var r = ctx.crmParseAddresses_('"Doe, Jane" <jane@acme.io>, bob@x.com, Sam Q <SAM@y.org>');
  eq(r.map(function (a) { return a.email; }), ['jane@acme.io', 'bob@x.com', 'sam@y.org']);
  assert.strictEqual(r[0].name, 'Doe, Jane');
  assert.strictEqual(r[1].name, '');
  assert.strictEqual(r[2].name, 'Sam Q');
});

test('split names', function () {
  eq(ctx.crmSplitName_('Jane Q. Doe'), { first: 'Jane', last: 'Doe' });
  eq(ctx.crmSplitName_('Doe, Jane'), { first: 'Jane', last: 'Doe' });
  eq(ctx.crmSplitName_('Cher'), { first: 'Cher', last: '' });
});

// --- machine addresses
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

// --- excerpts
test('excerpt strips quotes and signatures', function () {
  var body = 'Hi Jane,\n\nThanks for the intro — happy to chat Tuesday.\n\nBest,\nDavid\n--\nDavid K | CEO\n\nOn Mon, Jan 1, 2026 at 9:00 AM Jane <jane@acme.io> wrote:\n> Sure, let me know\n> when works.';
  var e = ctx.crmExcerpt_(body, 400);
  assert.ok(e.indexOf('wrote:') === -1, 'quoted header removed');
  assert.ok(e.indexOf('let me know') === -1, 'quoted lines removed');
  assert.ok(e.indexOf('CEO') === -1, 'signature removed');
  assert.ok(e.indexOf('happy to chat Tuesday') !== -1);
  assert.strictEqual(ctx.crmExcerpt_(body, 0), '');
  assert.ok(ctx.crmExcerpt_('word '.repeat(200), 50).length <= 52);
});

// --- rules
test('rule matching', function () {
  var m = function (rule, email, text) { return ctx.crmRuleMatches_(rule, { email: email, domain: ctx.crmDomainOf_(email), text: text || '' }); };
  assert.ok(m({ type: 'domain', pattern: '.vc' }, 'x@foo.vc'));
  assert.ok(!m({ type: 'domain', pattern: '.vc' }, 'x@foo.vcx.com'));
  assert.ok(m({ type: 'domain', pattern: 'a16z.com' }, 'x@mail.a16z.com'));
  assert.ok(!m({ type: 'domain', pattern: 'a16z.com' }, 'x@nota16z.com'));
  assert.ok(m({ type: 'email', pattern: 'Notifications@GitHub.com' }, 'notifications@github.com'));
  assert.ok(m({ type: 'keyword', pattern: 'term sheet' }, 'x@y.com', 'Re: Term Sheet draft'));
});

// --- stage computation
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
});

// --- dedupe
test('dedupe decisions', function () {
  var headers = ctx.CRM_MASTER_COLUMNS;
  var map = {}; headers.forEach(function (h, i) { map[h.toLowerCase()] = i + 1; });
  var master = { headers: headers, map: map, rows: [], index: {} };
  var add = function (o) {
    var row = headers.map(function () { return ''; });
    Object.keys(o).forEach(function (k) { row[map[k.toLowerCase()] - 1] = o[k]; });
    master.rows.push(row); master.index[o.Email] = master.rows.length - 1; return row;
  };
  var old = new CtxDate(Date.now() - 120 * 86400000), recent = new CtxDate(Date.now() - 3 * 86400000); // sandbox-realm Dates so instanceof works
  add({ Email: 'jane@acme.io', Name: 'Jane', Stage: 'Engaged', LastTouchAt: recent, Category: 'Investor' });
  add({ Email: 'old@acme.io', Name: 'Old', Stage: 'Contacted', LastTouchAt: old });
  add({ Email: 'fresh@beta.com', Name: 'Fresh', Stage: 'Contacted', LastTouchAt: recent });
  add({ Email: 'gone@beta.com', Name: 'Gone', Stage: 'Contacted', Bounced: true, LastTouchAt: old });
  add({ Email: 'p@gmail.com', Name: 'P', Stage: 'Engaged', LastTouchAt: recent });
  var byDomain = { 'acme.io': master.rows[0], 'beta.com': master.rows[2] };
  var settings = { ImportRecentDays: 30 };
  var D = function (email, seen) { return ctx.crmDedupeRow_(ctx.crmNormalizeEmail_(email), seen || {}, master, byDomain, settings); };

  assert.strictEqual(D('bad').status, 'INVALID');
  assert.strictEqual(D('JANE@acme.io').status, 'DUPLICATE');
  assert.strictEqual(D('jane@acme.io').recommendation, 'Skip — already in conversation');
  assert.strictEqual(D('old@acme.io').recommendation, 'Re-engage — reference prior thread');
  assert.ok(/^Skip — emailed 3 day/.test(D('fresh@beta.com').recommendation));
  assert.strictEqual(D('gone@beta.com').recommendation, 'Skip — do not contact');
  var sc = D('new@acme.io');
  assert.strictEqual(sc.status, 'SAME_COMPANY');
  assert.strictEqual(sc.matchedName, 'Jane');
  assert.ok(/in conversation with Jane/.test(sc.recommendation));
  assert.strictEqual(D('someone@gmail.com').status, 'NEW', 'freemail domains never match SAME_COMPANY');
  assert.strictEqual(D('x@nowhere.org').status, 'NEW');
  assert.strictEqual(D('x@nowhere.org', { 'x@nowhere.org': true }).status, 'DUPLICATE_IN_IMPORT');
});

test('import header aliases', function () {
  assert.strictEqual(ctx.crmCanonicalImportHeader_('e-mail address'), 'Email');
  assert.strictEqual(ctx.crmCanonicalImportHeader_('organization'), 'Company');
  assert.strictEqual(ctx.crmCanonicalImportHeader_('linkedin'), null);
});

test('tag helpers', function () {
  var t = ctx.crmSetTag_('warm, rel:replied', 'rel:', 'engaged');
  assert.strictEqual(t, 'warm, rel:engaged');
  assert.strictEqual(ctx.crmGetTag_(t, 'rel:'), 'engaged');
  assert.strictEqual(ctx.crmAddToList_('a, b', 'b'), 'a, b');
  assert.strictEqual(ctx.crmAddToList_('a, b', 'c', 2), 'b, c');
});

test('classification schema is structured-output safe', function () {
  var s = ctx.crmClassifySchema_(['Investor', 'Other']);
  var item = s.properties.contacts.items;
  assert.strictEqual(item.additionalProperties, false);
  eq(Object.keys(item.properties).sort(), item.required.slice().sort());
  assert.ok(item.properties.category.enum.indexOf('Ignore') !== -1);
});

console.log(passed + ' test group(s) passed' + (process.exitCode ? ', with failures' : ''));
