/**
 * Contact classification: deterministic Rules first, then Claude for the rest.
 *
 * Resumable like the scan: CRM_CLASSIFY_STATE + a 1-minute trigger on crmClassifyStep.
 * Locked rows (Lock = TRUE) are never touched.
 */

var CRM_CLASSIFY_STATE_PROP = 'CRM_CLASSIFY_STATE';
var CRM_API_URL = 'https://api.anthropic.com/v1/messages';
var CRM_RELATIONSHIPS = ['no_reply', 'replied', 'engaged', 'meeting_scheduled', 'customer', 'not_interested', 'unclear'];
var CRM_SENTIMENTS = ['positive', 'neutral', 'negative'];

function crmClassifyUnclassified() { var m = crmStartClassify_('unclassified'); crmToast_(m); return m; }
function crmClassifyAll() { var m = crmStartClassify_('all'); crmToast_(m); return m; }

/** deferOnly = true installs the trigger without running a step now (used at the end of a scan). */
function crmStartClassify_(mode, deferOnly) {
  if (!crmGetApiKey_()) {
    var msg = 'No Claude API key set. Rules still apply; run CRM → Set Claude API key… for AI classification.';
    crmLog_('CLASSIFY_INFO', '', msg);
  }
  var state = { mode: mode, startedAt: Date.now(), processed: 0, rule: 0, ai: 0, failed: 0 };
  PropertiesService.getScriptProperties().setProperty(CRM_CLASSIFY_STATE_PROP, JSON.stringify(state));
  crmEnsureMinuteTrigger_('crmClassifyStep');
  crmLog_('CLASSIFY_START', mode, deferOnly ? 'queued after scan' : '');
  if (!deferOnly) crmClassifyStep();
  return 'Classification started (' + mode + ').' + (crmGetApiKey_() ? '' : ' No API key set — only Rules will apply.');
}

function crmStopClassify() {
  crmDeleteTriggersFor_('crmClassifyStep');
  PropertiesService.getScriptProperties().deleteProperty(CRM_CLASSIFY_STATE_PROP);
  return 'Classification stopped.';
}

function crmClassifyStatus_() {
  var raw = PropertiesService.getScriptProperties().getProperty(CRM_CLASSIFY_STATE_PROP);
  return { running: !!raw, state: raw ? JSON.parse(raw) : null };
}

/** Trigger handler: rules pass, then AI batches until time budget or nothing left. */
function crmClassifyStep() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(CRM_CLASSIFY_STATE_PROP);
    if (!raw) { crmDeleteTriggersFor_('crmClassifyStep'); return; }
    var state = JSON.parse(raw);
    var settings = crmSettings_();
    var started = Date.now();

    var master = crmLoadTable_('Master', CRM_MASTER_COLUMNS, function (r, t) { return crmNormalizeEmail_(crmGet_(t, r, 'Email')); });
    var rules = crmLoadRules_();

    // Candidate rows.
    var candidates = [];
    master.rows.forEach(function (row, i) {
      var email = crmNormalizeEmail_(crmGet_(master, row, 'Email'));
      if (!email) return;
      if (crmBool_(crmGet_(master, row, 'Lock'))) return;
      var cat = String(crmGet_(master, row, 'Category')).trim();
      var by = String(crmGet_(master, row, 'ClassifiedBy')).trim();
      if (by === 'MANUAL') return;
      if (state.mode === 'unclassified' && cat) return;
      // On re-run after a partial step, skip rows classified since this job started.
      var at = crmGet_(master, row, 'ClassifiedAt');
      if (at instanceof Date && at.getTime() >= state.startedAt) return;
      candidates.push(i);
    });

    // 1) Rules (free, instant).
    var aiQueue = [];
    candidates.forEach(function (i) {
      var row = master.rows[i];
      var hit = crmApplyRules_(rules, master, row);
      if (hit) {
        crmSet_(master, row, 'Category', hit.category);
        if (hit.company && !String(crmGet_(master, row, 'Company')).trim()) crmSet_(master, row, 'Company', hit.company);
        crmSet_(master, row, 'Confidence', 1);
        crmSet_(master, row, 'ClassifiedBy', 'RULE');
        crmSet_(master, row, 'ClassifiedAt', new Date());
        state.rule++; state.processed++;
      } else {
        aiQueue.push(i);
      }
    });

    // 2) Claude, in batches, within the time budget.
    var apiKey = crmGetApiKey_();
    var remaining = aiQueue.length;
    if (apiKey) {
      for (var b = 0; b < aiQueue.length; b += settings.ClassifyBatchSize) {
        if (Date.now() - started > CRM_TIME_BUDGET_MS) break;
        var batchIdx = aiQueue.slice(b, b + settings.ClassifyBatchSize);
        var payload = batchIdx.map(function (i) { return crmContactPayload_(master, master.rows[i], settings); });
        try {
          var results = crmClassifyBatch_(payload, settings, apiKey);
          var byEmail = {};
          results.forEach(function (r) { byEmail[crmNormalizeEmail_(r.email)] = r; });
          batchIdx.forEach(function (i) {
            var row = master.rows[i];
            var r = byEmail[crmNormalizeEmail_(crmGet_(master, row, 'Email'))];
            if (r) { crmApplyAiResult_(master, row, r, settings); state.ai++; }
            else { crmSet_(master, row, 'ClassifiedAt', new Date()); state.failed++; crmLog_('CLASSIFY_MISS', crmGet_(master, row, 'Email'), 'Model returned no entry for this contact'); }
            state.processed++;
          });
          remaining -= batchIdx.length;
        } catch (e) {
          crmLog_('CLASSIFY_ERROR', 'batch of ' + batchIdx.length, String(e));
          state.failed += batchIdx.length;
          // Stamp so a poisoned batch does not loop forever; leave Category blank for a manual retry.
          batchIdx.forEach(function (i) { crmSet_(master, master.rows[i], 'ClassifiedAt', new Date()); });
          remaining -= batchIdx.length;
          if (/401|403|invalid.*api.*key|authentication/i.test(String(e))) { remaining = 0; break; }
        }
        crmSaveTable_(master); // persist progress batch by batch
      }
    } else {
      remaining = 0;
    }

    crmRecomputePipeline_(master);
    crmSaveTable_(master);

    if (remaining <= 0) {
      props.deleteProperty(CRM_CLASSIFY_STATE_PROP);
      crmDeleteTriggersFor_('crmClassifyStep');
      crmLog_('CLASSIFY_DONE', state.mode, state.processed + ' processed: ' + state.rule + ' by rule, ' + state.ai + ' by AI, ' + state.failed + ' failed/skipped' + (apiKey ? '' : ' (no API key — AI skipped)'));
      crmApplyValidation_(SpreadsheetApp.getActive());
    } else {
      props.setProperty(CRM_CLASSIFY_STATE_PROP, JSON.stringify(state));
      crmEnsureMinuteTrigger_('crmClassifyStep');
    }
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- Rules

/** [{ type, pattern, category, company }] from the Rules tab. */
function crmLoadRules_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Rules');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var map = crmColMap_(sheet, ['Type', 'Pattern', 'Category']);
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var out = [];
  values.forEach(function (r) {
    var type = String(r[map['type'] - 1]).trim().toLowerCase();
    var pattern = String(r[map['pattern'] - 1]).trim().toLowerCase();
    var category = String(r[map['category'] - 1]).trim();
    if (!type || !pattern || !category) return;
    out.push({ type: type, pattern: pattern, category: category, company: map['company'] ? String(r[map['company'] - 1]).trim() : '' });
  });
  return out;
}

/** subject: { email, domain, text } — text is subjects + excerpts + title, lowercased by the matcher. */
function crmRuleMatches_(rule, subject) {
  var p = rule.pattern;
  if (rule.type === 'email') return subject.email === p || crmNormalizeEmail_(p) === subject.email;
  if (rule.type === 'domain') {
    var d = subject.domain;
    if (p[0] === '.') return d.slice(-p.length) === p || d === p.slice(1);
    return d === p || d.slice(-(p.length + 1)) === '.' + p;
  }
  if (rule.type === 'keyword') return String(subject.text || '').toLowerCase().indexOf(p) !== -1;
  return false;
}

function crmApplyRules_(rules, master, row) {
  if (!rules.length) return null;
  var email = crmNormalizeEmail_(crmGet_(master, row, 'Email'));
  var subject = {
    email: email,
    domain: crmDomainOf_(email),
    text: [crmGet_(master, row, 'Subjects'), crmGet_(master, row, 'Title'), crmGet_(master, row, 'Company'),
      crmGet_(master, row, 'LastOutboundSnippet'), crmGet_(master, row, 'LastInboundSnippet')].join(' \n ')
  };
  for (var i = 0; i < rules.length; i++) {
    if (crmRuleMatches_(rules[i], subject)) {
      return { category: rules[i].category.toUpperCase() === 'IGNORE' ? 'Ignore' : rules[i].category, company: rules[i].company };
    }
  }
  return null;
}

// ---------------------------------------------------------------- Claude

function crmContactPayload_(master, row, settings) {
  var g = function (h) { return crmGet_(master, row, h); };
  var fmt = function (d) { return d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''; };
  var p = {
    email: String(g('Email')).trim(),
    name: String(g('Name')).trim(),
    domain: String(g('Domain')).trim(),
    company_guess: String(g('Company')).trim(),
    title_guess: String(g('Title')).trim(),
    emails_we_sent: Number(g('SentCount')) || 0,
    emails_they_sent: Number(g('ReceivedCount')) || 0,
    first_contact: fmt(g('FirstSentAt')),
    last_we_wrote: fmt(g('LastSentAt')),
    last_they_wrote: fmt(g('LastReplyAt')),
    subjects: crmSplitList_(g('Subjects')).slice(-6),
    source: String(g('Source')).trim()
  };
  if (settings.SnippetChars > 0) {
    p.our_latest_message = String(g('LastOutboundSnippet')).slice(0, settings.SnippetChars);
    p.their_latest_message = String(g('LastInboundSnippet')).slice(0, settings.SnippetChars);
  }
  return p;
}

function crmClassifySchema_(categories) {
  return {
    type: 'object',
    properties: {
      contacts: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            email: { type: 'string' },
            category: { type: 'string', enum: categories.concat(['Ignore']) },
            confidence: { type: 'number' },
            company: { type: 'string' },
            title: { type: 'string' },
            first_name: { type: 'string' },
            last_name: { type: 'string' },
            relationship: { type: 'string', enum: CRM_RELATIONSHIPS },
            sentiment: { type: 'string', enum: CRM_SENTIMENTS },
            summary: { type: 'string' },
            next_action: { type: 'string' }
          },
          required: ['email', 'category', 'confidence', 'company', 'title', 'first_name', 'last_name', 'relationship', 'sentiment', 'summary', 'next_action'],
          additionalProperties: false
        }
      }
    },
    required: ['contacts'],
    additionalProperties: false
  };
}

function crmSystemPrompt_(settings) {
  var cats = settings.Categories;
  var defs = {
    'Investor': 'VCs, angels, fund partners/associates, accelerators, scouts — anyone who might write a check or who we pitched.',
    'Engineer': 'engineers, designers, technical candidates or hires, technical collaborators, open-source contacts.',
    'Customer': 'a prospect or paying user of what we sell — people we pitched the product to or who use it.',
    'Partner': 'integration, distribution, channel, or co-marketing partners; other startups we collaborate with.',
    'Advisor': 'advisors, mentors, professors, experienced operators giving guidance.',
    'Press': 'journalists, bloggers, podcast hosts, newsletter writers.',
    'Recruiter': 'recruiters and talent agencies (not the candidates themselves).',
    'Vendor': 'people selling us something: lawyers, accountants, SaaS sales reps, agencies, banks.',
    'Personal': 'friends, family, personal admin unrelated to the company.',
    'Other': 'none of the above fits.'
  };
  var lines = cats.map(function (c) { return '- ' + c + ': ' + (defs[c] || 'as the name suggests.'); });
  return [
    'You classify people from a startup founder\'s sent email into a CRM.',
    settings.CompanyDescription ? 'About the startup: ' + settings.CompanyDescription : 'The startup has not described itself; infer from the emails.',
    '',
    'Categories (pick exactly one):',
    lines.join('\n'),
    '- Ignore: automated senders, mailing lists, receipts, or anything that is not a person we have a relationship with.',
    '',
    'Relationship (from the evidence only):',
    '- no_reply: we wrote, they never answered. - replied: they answered at least once. - engaged: a real back-and-forth or clear interest.',
    '- meeting_scheduled: a call/meeting was set up or happened. - customer: they are paying/using/signed. - not_interested: they declined or asked to stop. - unclear: cannot tell.',
    '',
    'Rules: return one entry per input contact, same email string, in the same order. Use the domain and signature-style cues to fill company and title; leave a field as an empty string when unknown rather than guessing. Split the display name into first_name/last_name. summary is one sentence about who they are and where things stand. next_action is one short imperative (e.g. "Send deck follow-up", "Book intro call", "None"). confidence is 0 to 1.'
  ].join('\n');
}

/** One API call for a batch of contact payloads. Returns the parsed contacts array. */
function crmClassifyBatch_(payload, settings, apiKey) {
  var body = {
    model: settings.Model,
    max_tokens: 4096,
    system: crmSystemPrompt_(settings),
    messages: [{ role: 'user', content: 'Classify these contacts:\n' + JSON.stringify(payload, null, 1) }],
    output_config: { effort: 'low', format: { type: 'json_schema', schema: crmClassifySchema_(settings.Categories) } },
    fallbacks: 'default'
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'anthropic-beta': 'server-side-fallback-2026-07-01' },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  };

  var attempt = 0, resp, code, text;
  while (true) {
    resp = UrlFetchApp.fetch(CRM_API_URL, options);
    code = resp.getResponseCode();
    text = resp.getContentText();
    if (code === 429 || code === 529 || code >= 500) {
      if (attempt++ >= 3) throw new Error('Claude API ' + code + ' after retries: ' + text.slice(0, 300));
      Utilities.sleep(Math.min(30000, 2000 * Math.pow(2, attempt)));
      continue;
    }
    break;
  }
  if (code !== 200) throw new Error('Claude API ' + code + ': ' + text.slice(0, 500));

  var data = JSON.parse(text);
  if (data.stop_reason === 'refusal') throw new Error('Model declined this batch (refusal)' + (data.stop_details && data.stop_details.category ? ': ' + data.stop_details.category : ''));
  if (data.stop_reason === 'max_tokens') throw new Error('Output truncated (max_tokens) — lower ClassifyBatchSize');
  var textBlock = (data.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('');
  var parsed = JSON.parse(textBlock);
  if (data.usage) crmLog_('CLASSIFY_USAGE', data.model || settings.Model, payload.length + ' contacts, ' + data.usage.input_tokens + ' in / ' + data.usage.output_tokens + ' out tokens');
  return parsed.contacts || [];
}

function crmApplyAiResult_(master, row, r, settings) {
  var cats = settings.Categories.concat(['Ignore']);
  var cat = cats.indexOf(r.category) !== -1 ? r.category : 'Other';
  crmSet_(master, row, 'Category', cat);
  crmSet_(master, row, 'Confidence', Math.max(0, Math.min(1, Number(r.confidence) || 0)));
  if (r.company && !String(crmGet_(master, row, 'Company')).trim()) crmSet_(master, row, 'Company', r.company);
  if (r.title && !String(crmGet_(master, row, 'Title')).trim()) crmSet_(master, row, 'Title', r.title);
  if (!String(crmGet_(master, row, 'FirstName')).trim() && r.first_name) crmSet_(master, row, 'FirstName', r.first_name);
  if (!String(crmGet_(master, row, 'LastName')).trim() && r.last_name) crmSet_(master, row, 'LastName', r.last_name);
  if (!String(crmGet_(master, row, 'Name')).trim() && (r.first_name || r.last_name)) crmSet_(master, row, 'Name', (r.first_name + ' ' + r.last_name).trim());
  crmSet_(master, row, 'Sentiment', CRM_SENTIMENTS.indexOf(r.sentiment) !== -1 ? r.sentiment : 'neutral');
  crmSet_(master, row, 'AISummary', String(r.summary || '').slice(0, 500));
  crmSet_(master, row, 'NextAction', String(r.next_action || '').slice(0, 200));
  crmSet_(master, row, 'ClassifiedBy', 'AI');
  crmSet_(master, row, 'ClassifiedAt', new Date());
  // The relationship hint feeds the stage recompute; it lives in Tags as "rel:<value>".
  crmSet_(master, row, 'Tags', crmSetTag_(crmGet_(master, row, 'Tags'), 'rel:', CRM_RELATIONSHIPS.indexOf(r.relationship) !== -1 ? r.relationship : 'unclear'));
}

/** Tags cell keeps free-form tags plus one machine tag per prefix (e.g. "rel:engaged"). */
function crmSetTag_(existing, prefix, value) {
  var tags = crmSplitList_(existing).filter(function (t) { return t.indexOf(prefix) !== 0; });
  tags.push(prefix + value);
  return tags.join(', ');
}

function crmGetTag_(existing, prefix) {
  var hit = crmSplitList_(existing).filter(function (t) { return t.indexOf(prefix) === 0; })[0];
  return hit ? hit.slice(prefix.length) : '';
}
