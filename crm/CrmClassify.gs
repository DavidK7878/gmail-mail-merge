/**
 * Contact classification: deterministic Rules first, then Claude for the rest.
 *
 * Rigor model:
 *  - Every contact is sent with an evidence pack: identity, counts, signature block, and
 *    the last few threads (subject, direction, our message, their message).
 *  - The model must return a short `evidence` quote for each label so labels are auditable.
 *  - Results below Settings → MinConfidence (or category Other / relationship unclear) get
 *    ReviewNeeded = TRUE. With RetryLowConfidence on, they are re-run one at a time at
 *    effort "high"; the better-supported answer wins.
 *  - Locked rows and rows with ClassifiedBy = MANUAL are never touched.
 *
 * Resumable like the scan: CRM_CLASSIFY_STATE + a 1-minute trigger on crmClassifyStep.
 */

var CRM_CLASSIFY_STATE_PROP = 'CRM_CLASSIFY_STATE';
var CRM_API_URL = 'https://api.anthropic.com/v1/messages';
var CRM_RELATIONSHIPS = ['no_reply', 'replied', 'engaged', 'meeting_scheduled', 'customer', 'not_interested', 'unclear'];
var CRM_SENTIMENTS = ['positive', 'neutral', 'negative'];

function crmClassifyUnclassified() { var m = crmStartClassify_('unclassified'); crmToast_(m); return m; }
function crmClassifyAll() { var m = crmStartClassify_('all'); crmToast_(m); return m; }

/** deferOnly = true installs the trigger without running a step now (used at the end of a scan). */
function crmStartClassify_(mode, deferOnly) {
  var props = PropertiesService.getScriptProperties();
  if (deferOnly && props.getProperty(CRM_CLASSIFY_STATE_PROP)) {
    crmEnsureMinuteTrigger_('crmClassifyStep'); // a job is already running; let it finish
    return 'Classification already running.';
  }
  var state = { mode: mode, startedAt: Date.now(), processed: 0, rule: 0, ai: 0, retried: 0, review: 0, failed: 0, errors: 0 };
  props.setProperty(CRM_CLASSIFY_STATE_PROP, JSON.stringify(state));
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

/** Trigger handler. */
function crmClassifyStep() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(CRM_CLASSIFY_STATE_PROP);
    if (!raw) { crmDeleteTriggersFor_('crmClassifyStep'); return; }
    var state = JSON.parse(raw);
    var finished = false;
    try {
      finished = crmClassifyTick_(state);
      state.errors = 0;
    } catch (e) {
      state.errors = (state.errors || 0) + 1;
      crmLog_('CLASSIFY_ERROR', 'tick', String(e && e.stack || e));
      if (state.errors >= 3) {
        props.deleteProperty(CRM_CLASSIFY_STATE_PROP);
        crmDeleteTriggersFor_('crmClassifyStep');
        crmLog_('CLASSIFY_ABORT', '', 'Three consecutive tick failures — classification stopped.');
        return;
      }
    }
    if (finished) {
      if (crmCommitState_(CRM_CLASSIFY_STATE_PROP, state)) {
        props.deleteProperty(CRM_CLASSIFY_STATE_PROP);
        crmDeleteTriggersFor_('crmClassifyStep');
        crmLog_('CLASSIFY_DONE', state.mode, state.processed + ' processed: ' + state.rule + ' by rule, ' + state.ai + ' by AI (' + state.retried + ' re-run at high effort, ' + state.review + ' flagged for review), ' + state.failed + ' failed' + (crmGetApiKey_() ? '' : ' (no API key — AI skipped)'));
        crmApplyValidation_(SpreadsheetApp.getActive());
      }
    } else if (crmCommitState_(CRM_CLASSIFY_STATE_PROP, state)) {
      crmEnsureMinuteTrigger_('crmClassifyStep');
    } else {
      crmLog_('CLASSIFY_SUPERSEDED', '', 'Tick finished after Stop / a newer Start; its state was discarded');
    }
  } finally {
    lock.releaseLock();
  }
}

/** One tick: rules pass, then AI batches within the time budget. Returns true when nothing is left. */
function crmClassifyTick_(state) {
  var settings = crmSettings_();
  var started = Date.now();
  var master = crmLoadTable_('Master', CRM_MASTER_COLUMNS, function (r, t) { return crmNormalizeEmail_(crmGet_(t, r, 'Email')); });
  var inter = crmLoadTable_('Interactions', CRM_INTERACTION_COLUMNS, function (r, t) { return String(crmGet_(t, r, 'ThreadId')).trim(); });
  var rules = crmLoadRules_();
  var candidates = crmClassifyCandidates_(master, state);

  // 1) Rules (free, instant, deterministic).
  var aiQueue = [];
  candidates.forEach(function (i) {
    var row = master.rows[i];
    var hit = crmApplyRules_(rules, master, row);
    if (hit) { crmApplyRuleResult_(master, row, hit); state.rule++; state.processed++; }
    else aiQueue.push(i);
  });

  // 2) Claude, in batches, within the time budget.
  var apiKey = crmGetApiKey_();
  var remaining = apiKey ? aiQueue.length : 0;
  try {
    if (apiKey) {
      for (var b = 0; b < aiQueue.length; b += settings.ClassifyBatchSize) {
        if (Date.now() - started > CRM_TIME_BUDGET_MS) break;
        var batchIdx = aiQueue.slice(b, b + settings.ClassifyBatchSize);
        var payload = batchIdx.map(function (i) { return crmContactPayload_(master, master.rows[i], inter, settings); });
        try {
          var results = crmClassifyBatch_(payload, settings, apiKey, 'low');
          var byEmail = {};
          results.forEach(function (r) { var k = crmNormalizeEmail_(r.email); if (k) byEmail[k] = r; });
          batchIdx.forEach(function (i) {
            var row = master.rows[i];
            var email = crmNormalizeEmail_(crmGet_(master, row, 'Email'));
            var r = byEmail[email];
            if (!r) {
              crmSet_(master, row, 'ClassifiedAt', new Date());
              crmSet_(master, row, 'ReviewNeeded', true);
              state.failed++; state.processed++;
              crmLog_('CLASSIFY_MISS', email, 'Model returned no entry for this contact');
              return;
            }
            // Second pass for thin results, one contact at a time with more effort.
            if (settings.RetryLowConfidence && crmNeedsReview_(r, settings) && Date.now() - started < CRM_TIME_BUDGET_MS) {
              try {
                var again = crmClassifyBatch_([crmContactPayload_(master, row, inter, settings)], settings, apiKey, 'high')[0];
                if (again && crmNormalizeEmail_(again.email) === email) {
                  state.retried++;
                  if ((Number(again.confidence) || 0) >= (Number(r.confidence) || 0)) r = again;
                }
              } catch (e2) {
                crmLog_('CLASSIFY_RETRY_ERROR', email, String(e2));
              }
            }
            crmApplyAiResult_(master, row, r, settings);
            if (crmBool_(crmGet_(master, row, 'ReviewNeeded'))) state.review++;
            state.ai++; state.processed++;
          });
        } catch (e) {
          crmLog_('CLASSIFY_ERROR', 'batch of ' + batchIdx.length, String(e));
          state.failed += batchIdx.length; state.processed += batchIdx.length;
          // Stamp so a poisoned batch does not loop forever; leave Category blank and flag for review.
          batchIdx.forEach(function (i) { crmSet_(master, master.rows[i], 'ClassifiedAt', new Date()); crmSet_(master, master.rows[i], 'ReviewNeeded', true); });
          if (/\b(401|403)\b|invalid.*api.*key|authentication/i.test(String(e))) { remaining = 0; break; }
        }
        remaining -= batchIdx.length;
        crmSaveTable_(master); // persist progress batch by batch
      }
    }
  } finally {
    crmRecomputePipeline_(master);
    crmSaveTable_(master);
  }
  return remaining <= 0;
}

/** Row indices eligible for this job. */
function crmClassifyCandidates_(master, state) {
  var out = [];
  master.rows.forEach(function (row, i) {
    var email = crmNormalizeEmail_(crmGet_(master, row, 'Email'));
    if (!email) return;
    if (crmBool_(crmGet_(master, row, 'Lock'))) return;
    if (String(crmGet_(master, row, 'ClassifiedBy')).trim() === 'MANUAL') return;
    var cat = String(crmGet_(master, row, 'Category')).trim();
    if (state.mode === 'unclassified' && cat) return;
    // On a later tick of the same job, skip rows already handled since the job started.
    var at = crmGet_(master, row, 'ClassifiedAt');
    if (at instanceof Date && at.getTime() >= state.startedAt) return;
    out.push(i);
  });
  return out;
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
    text: [crmGet_(master, row, 'Subjects'), crmGet_(master, row, 'Title'), crmGet_(master, row, 'Company'), crmGet_(master, row, 'Signature'),
      crmGet_(master, row, 'LastOutboundSnippet'), crmGet_(master, row, 'LastInboundSnippet')].join(' \n ')
  };
  for (var i = 0; i < rules.length; i++) {
    if (crmRuleMatches_(rules[i], subject)) {
      return { category: rules[i].category.toUpperCase() === 'IGNORE' ? 'Ignore' : rules[i].category, company: rules[i].company, rule: rules[i].type + ':' + rules[i].pattern };
    }
  }
  return null;
}

function crmApplyRuleResult_(master, row, hit) {
  crmSet_(master, row, 'Category', hit.category);
  if (hit.company && !String(crmGet_(master, row, 'Company')).trim()) crmSet_(master, row, 'Company', hit.company);
  crmSet_(master, row, 'Confidence', 1);
  crmSet_(master, row, 'ReviewNeeded', false);
  crmSet_(master, row, 'Evidence', 'Rule ' + hit.rule);
  crmSet_(master, row, 'ClassifiedBy', 'RULE');
  crmSet_(master, row, 'ClassifiedAt', new Date());
}

// ---------------------------------------------------------------- Evidence pack

/** Everything the model may use for one contact. Nothing else from the mailbox is sent. */
function crmContactPayload_(master, row, inter, settings) {
  var g = function (h) { return crmGet_(master, row, h); };
  var tz = Session.getScriptTimeZone();
  var fmt = function (d) { return d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : ''; };
  var email = String(g('Email')).trim();
  var p = {
    email: email,
    name: String(g('Name')).trim(),
    domain: String(g('Domain')).trim(),
    company_guess: String(g('Company')).trim(),
    title_guess: String(g('Title')).trim(),
    emails_we_sent: Number(g('SentCount')) || 0,
    emails_they_sent: Number(g('ReceivedCount')) || 0,
    auto_replies_from_them: Number(g('AutoReplies')) || 0,
    first_contact: fmt(g('FirstSentAt')),
    last_we_wrote: fmt(g('LastSentAt')),
    last_they_wrote: fmt(g('LastReplyAt')),
    they_spoke_last: crmBool_(g('NeedsReply')),
    subjects: crmSplitList_(g('Subjects')).slice(-6),
    source: String(g('Source')).trim()
  };
  if (settings.SnippetChars > 0) {
    p.their_signature = String(g('Signature')).slice(0, 300);
    p.threads = crmEvidenceThreads_(master, row, inter, settings);
    if (!p.threads.length) {
      // Fall back to the aggregate excerpts if Interactions is missing (e.g. contact came from sync/import).
      p.our_latest_message = String(g('LastOutboundSnippet')).slice(0, settings.SnippetChars);
      p.their_latest_message = String(g('LastInboundSnippet')).slice(0, settings.SnippetChars);
    }
  }
  return p;
}

/** Up to EvidenceThreads most recent threads involving this contact, newest first. */
function crmEvidenceThreads_(master, row, inter, settings) {
  if (!inter || !settings.EvidenceThreads) return [];
  var email = crmNormalizeEmail_(crmGet_(master, row, 'Email'));
  var tz = Session.getScriptTimeZone();
  var items = [];
  crmSplitList_(crmGet_(master, row, 'ThreadIds')).forEach(function (tid) {
    var idx = inter.index[tid];
    if (idx === undefined) return;
    var ir = inter.rows[idx];
    var parts = {};
    try { parts = JSON.parse(crmGet_(inter, ir, 'Participants') || '{}'); } catch (e) { /* ignore */ }
    if (!parts[email]) return;
    var last = crmGet_(inter, ir, 'LastMessageAt');
    items.push({
      _t: last instanceof Date ? last.getTime() : 0,
      subject: String(crmGet_(inter, ir, 'Subject')).slice(0, 120),
      last_message: last instanceof Date ? Utilities.formatDate(last, tz, 'yyyy-MM-dd') : '',
      messages_from_us: Number(parts[email][0]) || 0,
      messages_from_them: Number(parts[email][1]) || 0,
      last_direction: String(crmGet_(inter, ir, 'LastDirection')),
      latest_message: String(crmGet_(inter, ir, 'LastSnippet')).slice(0, settings.SnippetChars),
      their_latest_message: String(crmGet_(inter, ir, 'LastInboundSnippet')).slice(0, settings.SnippetChars)
    });
  });
  items.sort(function (a, b) { return b._t - a._t; });
  return items.slice(0, settings.EvidenceThreads).map(function (it) { delete it._t; return it; });
}

// ---------------------------------------------------------------- Claude

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
            evidence: { type: 'string' },
            company: { type: 'string' },
            title: { type: 'string' },
            first_name: { type: 'string' },
            last_name: { type: 'string' },
            relationship: { type: 'string', enum: CRM_RELATIONSHIPS },
            sentiment: { type: 'string', enum: CRM_SENTIMENTS },
            summary: { type: 'string' },
            next_action: { type: 'string' }
          },
          required: ['email', 'category', 'confidence', 'evidence', 'company', 'title', 'first_name', 'last_name', 'relationship', 'sentiment', 'summary', 'next_action'],
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
    'Investor': 'VCs, angels, fund partners/associates, accelerators, scouts — anyone who might write a check or whom we pitched for funding.',
    'Engineer': 'engineers, designers, technical candidates or hires, technical collaborators, open-source contacts.',
    'Customer': 'a prospect or paying user of what we sell — people we pitched the product to or who use it.',
    'Partner': 'integration, distribution, channel, or co-marketing partners; other startups we collaborate with.',
    'Advisor': 'advisors, mentors, professors, experienced operators giving guidance.',
    'Press': 'journalists, bloggers, podcast hosts, newsletter writers.',
    'Recruiter': 'recruiters and talent agencies (not the candidates themselves).',
    'Vendor': 'people selling us something: lawyers, accountants, SaaS sales reps, agencies, banks.',
    'Personal': 'friends, family, personal admin unrelated to the company.',
    'Other': 'none of the above fits, or the evidence does not support any specific category.'
  };
  var lines = cats.map(function (c) { return '- ' + c + ': ' + (defs[c] || 'as the name suggests.'); });
  return [
    'You classify people from a startup founder\'s sent email into a CRM. Your labels drive who gets follow-ups, so prefer an honest low confidence over a confident guess.',
    settings.CompanyDescription ? 'About the startup: ' + settings.CompanyDescription : 'The startup has not described itself; infer from the emails and say so in evidence when that limits you.',
    '',
    'Categories (pick exactly one):',
    lines.join('\n'),
    '- Ignore: automated senders, mailing lists, receipts, or anything that is not a person we have a relationship with.',
    '',
    'Relationship, from the evidence only:',
    '- no_reply: we wrote, they never answered (auto-replies do not count). - replied: they answered at least once. - engaged: a real back-and-forth or clear interest.',
    '- meeting_scheduled: a call/meeting was set up or happened. - customer: they are paying/using/signed. - not_interested: they declined or asked to stop. - unclear: cannot tell.',
    '',
    'How to weigh evidence: what the messages say beats what the domain suggests (a person at a VC firm who is asking us for a job is an Engineer, not an Investor). A signature block is the best source for company and title. Message counts tell you about the relationship, not the category. If the only evidence is an email address and a subject line, confidence must be 0.5 or lower.',
    '',
    'Output rules: return exactly one entry per input contact, using the same email string, in the same order. evidence is a short quote or paraphrase (max 25 words) of the specific detail that justified category and relationship. Leave company/title/first_name/last_name as empty strings when unknown rather than guessing. summary is one sentence about who they are and where things stand. next_action is one short imperative, or "None". confidence is a number from 0 to 1.'
  ].join('\n');
}

/**
 * One API call for a batch of contact payloads at the given effort. Returns validated contact results.
 * Throws on transport errors, refusals, truncation, or unparseable output.
 */
function crmClassifyBatch_(payload, settings, apiKey, effort) {
  var body = {
    model: settings.Model,
    max_tokens: 16000,
    system: crmSystemPrompt_(settings),
    messages: [{ role: 'user', content: 'Classify these contacts:\n' + JSON.stringify(payload, null, 1) }],
    output_config: { effort: effort || 'low', format: { type: 'json_schema', schema: crmClassifySchema_(settings.Categories) } },
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
  var parsed;
  try { parsed = JSON.parse(textBlock); } catch (e) { throw new Error('Model output was not valid JSON: ' + textBlock.slice(0, 200)); }
  var contacts = Array.isArray(parsed.contacts) ? parsed.contacts : [];
  if (data.usage) crmLog_('CLASSIFY_USAGE', (data.model || settings.Model) + ' @' + (effort || 'low'), payload.length + ' contacts, ' + data.usage.input_tokens + ' in / ' + data.usage.output_tokens + ' out tokens');

  // Validate: only accept entries for contacts we asked about; coerce fields into range.
  var asked = {};
  payload.forEach(function (p) { asked[crmNormalizeEmail_(p.email)] = true; });
  var out = [];
  contacts.forEach(function (c) {
    var k = crmNormalizeEmail_(c && c.email);
    if (!k || !asked[k]) { crmLog_('CLASSIFY_UNEXPECTED', c && c.email, 'Model returned an email that was not in the batch'); return; }
    out.push(crmCoerceResult_(c, settings));
  });
  if (out.length !== payload.length) crmLog_('CLASSIFY_PARTIAL', '', payload.length + ' asked, ' + out.length + ' usable answers');
  return out;
}

function crmCoerceResult_(c, settings) {
  var cats = settings.Categories.concat(['Ignore']);
  var conf = Number(c.confidence);
  return {
    email: String(c.email || ''),
    category: cats.indexOf(c.category) !== -1 ? c.category : 'Other',
    confidence: isNaN(conf) ? 0 : Math.max(0, Math.min(1, conf)),
    evidence: String(c.evidence || '').slice(0, 300),
    company: String(c.company || '').trim(),
    title: String(c.title || '').trim(),
    first_name: String(c.first_name || '').trim(),
    last_name: String(c.last_name || '').trim(),
    relationship: CRM_RELATIONSHIPS.indexOf(c.relationship) !== -1 ? c.relationship : 'unclear',
    sentiment: CRM_SENTIMENTS.indexOf(c.sentiment) !== -1 ? c.sentiment : 'neutral',
    summary: String(c.summary || '').slice(0, 500),
    next_action: String(c.next_action || '').slice(0, 200)
  };
}

/** A result is "thin" when confidence is low, or the model could not commit. */
function crmNeedsReview_(r, settings) {
  return r.confidence < settings.MinConfidence || r.category === 'Other' || r.relationship === 'unclear' || !r.evidence;
}

function crmApplyAiResult_(master, row, r, settings) {
  crmSet_(master, row, 'Category', r.category);
  crmSet_(master, row, 'Confidence', r.confidence);
  crmSet_(master, row, 'Evidence', r.evidence);
  crmSet_(master, row, 'ReviewNeeded', crmNeedsReview_(r, settings));
  if (r.company && !String(crmGet_(master, row, 'Company')).trim()) crmSet_(master, row, 'Company', r.company);
  if (r.title && !String(crmGet_(master, row, 'Title')).trim()) crmSet_(master, row, 'Title', r.title);
  if (!String(crmGet_(master, row, 'FirstName')).trim() && r.first_name) crmSet_(master, row, 'FirstName', r.first_name);
  if (!String(crmGet_(master, row, 'LastName')).trim() && r.last_name) crmSet_(master, row, 'LastName', r.last_name);
  if (!String(crmGet_(master, row, 'Name')).trim() && (r.first_name || r.last_name)) crmSet_(master, row, 'Name', (r.first_name + ' ' + r.last_name).trim());
  crmSet_(master, row, 'Sentiment', r.sentiment);
  crmSet_(master, row, 'AISummary', r.summary);
  crmSet_(master, row, 'NextAction', r.next_action);
  crmSet_(master, row, 'ClassifiedBy', 'AI');
  crmSet_(master, row, 'ClassifiedAt', new Date());
  // The relationship hint feeds the stage recompute; it lives in Tags as "rel:<value>".
  crmSet_(master, row, 'Tags', crmSetTag_(crmGet_(master, row, 'Tags'), 'rel:', r.relationship));
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

// ---------------------------------------------------------------- Preview (dry run)

/**
 * Classify up to `n` unclassified contacts and return the results WITHOUT writing them,
 * so the prompt and CompanyDescription can be checked before spending on the whole list.
 */
function crmPreviewClassification(n) {
  var apiKey = crmGetApiKey_();
  if (!apiKey) throw new Error('Set the Claude API key first.');
  var settings = crmSettings_();
  var master = crmLoadTable_('Master', CRM_MASTER_COLUMNS, function (r, t) { return crmNormalizeEmail_(crmGet_(t, r, 'Email')); });
  var inter = crmLoadTable_('Interactions', CRM_INTERACTION_COLUMNS, function (r, t) { return String(crmGet_(t, r, 'ThreadId')).trim(); });
  var rules = crmLoadRules_();
  var pick = [];
  var state = { mode: 'unclassified', startedAt: Date.now() };
  crmClassifyCandidates_(master, state).forEach(function (i) {
    if (pick.length >= (Number(n) || 5)) return;
    if (crmApplyRules_(rules, master, master.rows[i])) return; // rules would take these anyway
    pick.push(i);
  });
  if (!pick.length) {
    // Nothing unclassified: preview on the most recently touched contacts instead.
    var order = master.rows.map(function (r, i) { return i; }).filter(function (i) { return crmNormalizeEmail_(crmGet_(master, master.rows[i], 'Email')); });
    order.sort(function (a, b) { var da = crmGet_(master, master.rows[a], 'LastTouchAt'), db = crmGet_(master, master.rows[b], 'LastTouchAt'); return (db instanceof Date ? db.getTime() : 0) - (da instanceof Date ? da.getTime() : 0); });
    pick = order.slice(0, Number(n) || 5);
  }
  if (!pick.length) throw new Error('Master is empty — run a scan first.');
  var payload = pick.map(function (i) { return crmContactPayload_(master, master.rows[i], inter, settings); });
  var results = crmClassifyBatch_(payload, settings, apiKey, 'low');
  return results.map(function (r) {
    return { email: r.email, category: r.category, confidence: r.confidence, relationship: r.relationship, evidence: r.evidence, summary: r.summary, review: crmNeedsReview_(r, settings) };
  });
}
