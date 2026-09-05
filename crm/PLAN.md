# Email CRM Pipeline — Plan

A Google-Sheets-native CRM that reads your **sent** Gmail, figures out who you have been talking to, uses Claude to label each contact (investor, engineer, prospect, …), keeps one **Master** sheet of every relationship, and de-duplicates any new sourcing list against it before it goes into the mail merge.

It is built with the same stack and conventions as the mail merge tool in the repo root: Apps Script bound to a Google Sheet, Gmail as the only data source, time-based triggers to stay under the 6-minute execution cap, a `Settings` tab for knobs, a `Log` tab for the audit trail, and a sidebar for the day-to-day buttons.

---

## 1. Goals

| Goal | What it means in practice |
|---|---|
| **One source of truth for every contact** | A `Master` tab with one row per email address you have ever emailed or been emailed by, with counts, dates, latest excerpts, and thread links. |
| **AI classification** | Each contact gets a `Category` (Investor, Engineer, Customer, Partner, Advisor, Press, Recruiter, Vendor, Personal, Other), a pipeline `Stage`, a one-line `AISummary`, `Sentiment`, and a suggested `NextAction`. |
| **Relationship state, not just labels** | `Stage` moves automatically: New → Contacted → Replied → Engaged → Meeting → Won, with Dormant / Not Interested / Bounced overlays. `NeedsReply` flags threads where the other person spoke last. |
| **Dedupe before you send** | Paste or import a new sourcing list into `Import`; the tool marks each row NEW / DUPLICATE / SAME_COMPANY / SUPPRESSED, explains why, and pushes only the NEW rows into the mail merge `Contacts` tab. |
| **Closed loop with the mail merge** | Campaign sends show up in Master automatically (they are sent mail). Bounces and replies recorded by the mail merge sync into Master as suppression signals. |
| **Cheap and private by default** | Deterministic `Rules` run before the AI so obvious cases cost nothing. Only short excerpts (not whole emails) go to the API. The API key lives in script properties, never in a cell. |

Non-goals for v1: inbound-only contacts you never wrote to (newsletters, cold inbound), calendar integration, multi-user ownership, writing back to Gmail (labels). These are listed in §8 as later phases.

---

## 2. Feature list

### 2.1 Gmail scan (`CrmScan.gs`)
- **Windowed, date-bounded queries.** Gmail orders search results by latest activity, so paging one big query by offset is not stable while mail arrives. The scan instead walks the date range in windows of `ScanWindowDays` with `after:<epoch> before:<epoch>` bounds, which match on the *sent message's* date and never move. Within a window, pages of 100 are fetched by offset; a window only advances once its final page has been fully processed.
- **Two phases for incremental scans.** Phase 1 is `in:sent` over the window (threads where we wrote). Phase 2 is `-in:sent …` over the same window, which finds **replies that arrived on older threads** — the case a sent-only query misses entirely. Threads without any outbound message are discarded, so phase 2 adds nothing but the reply data.
- **Incremental** scans start 2 days before the previous scan's start; **Full rescan** goes back to `Settings → ScanSince`.
- **Cheap idempotency.** A thread whose message count and last-message date match its `Interactions` row is skipped without reading its messages. Re-running a scan is therefore both safe (counts are rebuilt from per-thread participant maps) and fast.
- **Resumable and interruptible.** State (phase, window cursor, page offset, counters) lives in Script Properties; a 1-minute trigger calls `crmScanStep` until done, each tick stopping at ~4.5 minutes. State is written back only if the job is still the current one, so **Stop** or a newer **Start** issued during a tick wins. Three consecutive tick failures abort the job with a `SCAN_ABORT` log line instead of retrying forever.
- For each thread it reads every message and decides direction by comparing the sender to `MyDomains` / your own address:
  - **Outbound** → every To/Cc recipient becomes/updates a contact: `SentCount`, `FirstSentAt`, `LastSentAt`, `Subjects`, `LastOutboundSnippet`, `ThreadIds`.
  - **Inbound** → the sender's contact gets `ReceivedCount`, `LastReplyAt`, `LastInboundSnippet`.
  - Thread whose last message is inbound → that sender's `NeedsReply = TRUE`.
- **Auto-replies are not replies.** Out-of-office / vacation responders / delivery notices are detected by subject and by `Auto-Submitted`, `X-Autoreply`, `Precedence` headers. They are counted in `AutoReplies`, never in `ReceivedCount`, never set `NeedsReply`, and mark the thread `LastDirection = AUTO`.
- **Opt-outs are honoured deterministically.** An inbound body matching a conservative phrase list ("remove me from your list", "unsubscribe", "stop emailing me", …) sets `DoNotContact = TRUE`, adds the `optout` tag, and logs `OPT_OUT`. Imports then recommend Skip.
- **Signature capture.** The signature block of the contact's latest inbound message (after `--`, or a heuristic sign-off paragraph) is stored in `Signature` and fed to the classifier; it is removed from the excerpt so the excerpt is what they actually said.
- Excerpts are the first ~400 chars of the plain body **with quoted history stripped** (Gmail, Outlook, French, German, Spanish reply headers, `>` lines, forwarded-message separators).
- Skips machine addresses (`noreply@`, `mailer-daemon`, `calendar-notification@google.com`, …), anything on `IgnoreDomains`, and anything a `Rules` row marks `IGNORE`.
- Writes one row per thread to `Interactions` (subject, dates, counts, last direction, latest inbound excerpt, Gmail link) plus a `Participants` map of per-contact counts. Master counts are **rebuilt from that map**, so re-scanning a thread in the 2-day overlap window can never double count.

### 2.2 Classification (`CrmClassify.gs`)
1. **Rules first** (the `Rules` tab): `domain` (exact or suffix like `.vc`), `email`, or `keyword` (matched against subjects, excerpts, signature, title) → Category and optional Company. First match wins, costs nothing, and `IGNORE` drops the contact. Rule hits record `Evidence = Rule domain:a16z.com` so every label has a stated reason.
2. **Claude second** for everything still unclassified, with an **evidence pack** per contact: identity, counts (auto-replies separated out), dates, who spoke last, the **signature block** captured from their mail (the best source of title and company), and the last `EvidenceThreads` threads (subject, direction, our latest message, their latest message). Nothing else from the mailbox is sent.
3. The request uses **structured output** (a JSON schema) so the reply always parses, `effort: low` for the first pass, and `fallbacks: "default"` so a safety-classifier decline is re-run server-side instead of leaving a row blank. Every response is **validated**: entries for emails that were not in the batch are dropped and logged, enums are coerced, confidence is clamped to 0–1, partial answers are logged.
4. Output per contact: `category`, `confidence`, **`evidence`** (≤ 25 words quoting the detail that justified the label), `company`, `title`, `first_name`, `last_name`, `relationship` (no_reply / replied / engaged / meeting_scheduled / customer / not_interested / unclear), `sentiment`, `summary`, `next_action`.
5. **Review queue.** A result is *thin* when `confidence < MinConfidence`, or category is `Other`, or relationship is `unclear`, or evidence is empty. Thin results are re-run **one contact at a time at effort `high`** (`RetryLowConfidence`); the answer with the higher confidence wins. Whatever remains thin gets `ReviewNeeded = TRUE` and shows up on the Dashboard and in the sidebar so a human looks at it.
6. **`Lock = TRUE`** freezes Category / Stage / Company / Title / Name against both rules and AI. `ClassifiedBy = MANUAL` does the same for classification only.
7. **Preview** (menu or sidebar): classify five contacts and show category, confidence, evidence, and summary *without saving*, so the prompt and `CompanyDescription` can be sanity-checked before spending on the whole list.
8. Runs as a resumable trigger job (`crmClassifyStep`); auto-queued when a scan finishes (`AutoClassifyAfterScan`).

### 2.3 Pipeline (`CrmPipeline.gs`)
Stage is recomputed from facts after every scan/classify (`crmRecomputePipeline`):

| Stage | Rule |
|---|---|
| `Bounced` | `Bounced = TRUE` (from mail merge sync or a mailer-daemon reply) |
| `Not Interested` | `DoNotContact = TRUE` or AI relationship `not_interested` |
| `Won` | AI relationship `customer` |
| `Meeting` | AI relationship `meeting_scheduled` |
| `Engaged` | `ReceivedCount ≥ 2` or AI relationship `engaged` |
| `Replied` | `ReceivedCount ≥ 1` |
| `Contacted` | `SentCount ≥ 1` |
| `New` | imported, never emailed |
| `Dormant` | overlay: stage was Replied/Engaged/Meeting and `LastTouchAt` older than `DormantDays` |

`LastTouchAt = max(LastSentAt, LastReplyAt)`. `NeedsReply` comes from the scan. Locked rows keep their Stage.

**Audit** (`crmAudit`) checks the invariants the pipeline is supposed to maintain and reports counts with examples: duplicate normalized emails, invalid addresses, unclassified rows, low confidence not flagged, AI labels without evidence, stage disagreeing with the facts, replied-without-excerpt, counts without thread ids, `NeedsReply` on suppressed contacts, opt-outs that were un-suppressed, `Ignore` rows still present. Runs from the menu or sidebar and logs to `Log`.

**Sync from mail merge** (`crmSyncFromMailMerge`): reads the mail merge `Contacts` tab (same spreadsheet, or the sheet at `MailMergeSheetUrl`), creates missing Master rows with `Source = MailMerge`, and copies `Bounced` / `Replied` flags in.

### 2.4 Import & dedupe (`CrmImport.gs`)
- `Import` tab: paste rows with at least an `Email` header, or use **Import from Drive file…** (Google Sheet URL/ID or CSV) which maps common header variants (`E-mail`, `Full Name`, `Organization`, `Role`…).
- **Check Import against Master** fills, per row:
  - `DedupeStatus`: `NEW`, `DUPLICATE` (same normalized email), `QUEUED` (already in the mail merge `Contacts` tab), `POSSIBLE_MATCH` (same normalized name **and** same company or domain — the same person under another address), `SAME_COMPANY` (same non-freemail domain), `DUPLICATE_IN_IMPORT`, `INVALID`.
  - `MatchedName / MatchedCategory / MatchedStage / LastTouch` from the matching Master row (for SAME_COMPANY, the most advanced contact at that company).
  - `Recommendation`: `Send`, `Skip — do not contact`, `Skip — already in conversation`, `Skip — emailed N days ago`, `Re-engage — reference prior thread`, `Review — colleague already contacted (name)`.
- **Push NEW to Mail Merge** appends NEW (and, if you choose, SAME_COMPANY) rows to the `Contacts` tab with `Email`, `FirstName`, `Company`, plus any Import column whose header also exists in `Contacts`. Each pushed row is also added to Master as `Stage = New`, `Source = Import`, so the next import dedupes against it even before it is sent. `PushedAt` is stamped so a second click does not double-push.
- Email normalization for matching: lowercase, trimmed, Gmail dots and `+tags` removed, `googlemail.com → gmail.com`.

### 2.5 Dashboard & sidebar
- `Dashboard`: totals, contacts by Category × Stage matrix, reply rate per Category, `NeedsReply` count, Dormant-but-warm count, unclassified count, last scan time.
- Sidebar (`CrmSidebar.html`): status card, Scan (new / full), Classify (unclassified / all), Import (check / push), Pipeline (recompute / sync), API-key indicator, live progress line.

---

## 3. Data model

### `Master` (one row per contact)
| Group | Columns |
|---|---|
| Identity | `Email`, `Name`, `FirstName`, `LastName`, `Company`, `Domain`, `Title` |
| AI / rules | `Category`, `Stage`, `Confidence`, `ReviewNeeded`, `Evidence` (the quote that justified the label), `Sentiment`, `AISummary`, `NextAction`, `ClassifiedBy` (RULE / AI / MANUAL), `ClassifiedAt` |
| Yours to edit | `Lock`, `Tags`, `Notes`, `DoNotContact` |
| Facts from Gmail | `Source`, `FirstSentAt`, `LastSentAt`, `LastReplyAt`, `LastTouchAt`, `SentCount`, `ReceivedCount`, `AutoReplies`, `ThreadCount`, `NeedsReply`, `Bounced`, `Subjects`, `LastOutboundSnippet`, `LastInboundSnippet`, `Signature`, `ThreadIds`, `UpdatedAt` |

Extra columns you add to the right are preserved (the code addresses columns by header, like the mail merge does).

### `Interactions` (one row per thread)
`ThreadId`, `Email`, `Subject`, `FirstMessageAt`, `LastMessageAt`, `Messages`, `Outbound`, `Inbound`, `AutoReplies`, `LastDirection` (OUT / IN / AUTO), `LastFrom`, `LastSnippet`, `LastInboundSnippet`, `Link`, `Participants` (JSON `{email: [sent, received, autoReplies]}` — the authoritative per-thread counts)

### `Import` (staging)
Your columns (`Email` required) + machine columns `DedupeStatus`, `MatchedName`, `MatchedCategory`, `MatchedStage`, `LastTouch`, `Recommendation`, `PushedAt`.

### `Rules`
`Type` (domain / email / keyword), `Pattern`, `Category`, `Company`, `Notes`. Seeded with a handful of examples.

### `Settings`
| Key | Default | Purpose |
|---|---|---|
| `CompanyDescription` | (blank) | One paragraph about your startup, fed to the classifier |
| `MyDomains` | (blank) | Comma-separated domains/addresses treated as "us" (co-founders, team) |
| `ScanSince` | 12 months ago | Earliest date for a full rescan |
| `MaxThreadsPerRun` | 250 | Threads per trigger step |
| `IgnoreDomains` | (blank) | Extra domains never added to Master |
| `Model` | `claude-opus-5` | Claude model id |
| `ClassifyBatchSize` | 8 | Contacts per API request |
| `AutoClassifyAfterScan` | TRUE | Start classification when a scan completes |
| `DormantDays` | 45 | Days without touch before a warm contact goes Dormant |
| `Categories` | Investor, Engineer, Customer, Partner, Advisor, Press, Recruiter, Vendor, Personal, Other | Allowed categories (drives the schema enum and the dashboard) |
| `SnippetChars` | 400 | Excerpt length stored and sent to the API |
| `MailMergeSheetUrl` | (blank) | Blank = the `Contacts` tab in this spreadsheet |
| `ImportRecentDays` | 30 | "emailed recently" threshold for dedupe recommendations |
| `MinConfidence` | 0.7 | Below this an AI label is flagged `ReviewNeeded` |
| `RetryLowConfidence` | TRUE | Re-run thin results one at a time at effort `high` |
| `EvidenceThreads` | 3 | Recent threads per contact sent as evidence |

The **Claude API key** is set via `CRM → Set Claude API key…` and stored in Script Properties (`CLAUDE_API_KEY`), never in the sheet.

---

## 4. Architecture

```
Gmail (in:sent) ──crmScanStep (1-min trigger, resumable)──▶ Master + Interactions
                                                                 │
Rules tab ──────────────────────────────┐                        │
Claude Messages API (structured JSON) ──┴─ crmClassifyStep ──────▶ Category / Stage / Summary
                                                                 │
mail merge Contacts tab ── crmSyncFromMailMerge ─────────────────▶ Bounced / Replied / new rows
                                                                 │
new sourcing list ── Import tab ── crmCheckImport ── dedupe ─────┘
                                      │
                                      └── crmPushNewToMailMerge ──▶ Contacts tab (ready to send)
```

- **Everything runs inside the Sheet's Apps Script project.** No server, no database, no OAuth app to publish.
- **All function and file names carry a `Crm`/`crm` prefix**, so the files can be pasted into the *same* Apps Script project as the mail merge without collisions (see README for the one-line `onOpen` change).
- **Whole-tab read/modify/write** for Master: rows are loaded into memory once per step, mutated, and written back with a single `setValues`. Thousands of rows stay well under quota.
- **Time budget + trigger continuation** is the same pattern `Send.gs` uses (`BATCH_TIME_BUDGET_MS`), so scans of years of mail simply take several ticks.
- **Claude API via `UrlFetchApp`** (Apps Script has no SDK). Retries with backoff on 429/5xx/529, honours `stop_reason` (`refusal`, `max_tokens`), logs token usage to `Log`.

---

## 5. Scopes

`gmail.readonly` (scan), `spreadsheets` (Master lives here; the mail merge sheet may be a different file), `drive.readonly` (import from Drive), `script.external_request` (Claude API), `script.scriptapp` (triggers), `script.container.ui`, `userinfo.email`.

---

## 6. Correctness model & QA

### Invariants the code maintains
| Invariant | Enforced by |
|---|---|
| One Master row per normalized email | scan/import/sync all key on `crmNormalizeEmail_`; audit reports violations from hand edits |
| Counts equal the sum over threads | `Interactions.Participants` is authoritative; `crmRebuildCountsFromInteractions_` recomputes on every touch |
| Re-scanning is idempotent | same thread → same Participants map → same counts; excerpts only move forward in time |
| An auto-reply never counts as a reply or triggers NeedsReply | `crmIsAutoReply_` gate before any inbound accounting |
| An opt-out is never re-contacted | `DoNotContact` set deterministically at scan time; dedupe recommends Skip; audit flags un-suppression |
| Every label has a reason | rules write `Evidence = Rule …`; AI schema requires `evidence`; empty evidence ⇒ `ReviewNeeded` |
| Thin evidence is never presented as certain | `MinConfidence` gate, `Other`/`unclear` ⇒ review, second pass at high effort |
| Human edits win | `Lock`, `ClassifiedBy = MANUAL`, hand-set `Won`/`Not Interested` all survive recompute |
| A hand edit made *while a job runs* is not clobbered | `crmSaveTable_` re-reads the sheet at save time and merges only the cells the job changed, matched by key (so sorts and deletions during a tick are respected); a `Lock` ticked mid-job overrides the job's classification columns |
| Stop means stop | `crmCommitState_` refuses to write a job's state if it was deleted or replaced during the tick; the trigger is not re-created |
| A scan never skips a thread because new mail arrived | date-bounded windows match on immutable message dates; phase 2 covers replies on old threads |
| Nothing from the mailbox leaves except the evidence pack | `crmContactPayload_` is the only builder; `SnippetChars = 0` strips all body text |
| Model output is untrusted | emails not in the batch dropped, enums coerced, confidence clamped, refusal/truncation/bad JSON raise |

### Edge cases handled
Gmail dot/plus aliases; `googlemail.com`; quoted names with commas; wrapped `On … wrote:` headers; Outlook `From:/Sent:` blocks; forwarded-message separators; `--` and heuristic signatures; teammates on Cc (`MyDomains` + own non-freemail domain); mailer-daemon bounces inside sent threads; SaaS reply-token addresses; freemail domains never matched as a company; Excel files rejected with a clear message; API 429/5xx/529 retried with backoff; 401/403 stop the job; a model reply missing a contact is logged and flagged rather than silently skipped.

### Test suite (`tests/run.js`, 32 groups)
The `.gs` files run inside a Node `vm` sandbox with fakes for Gmail, Sheets, Properties, and `UrlFetchApp`. Beyond the pure helpers, the suite drives:
- **the real scan** (`crmProcessThread_`) with fake threads: aggregates, NeedsReply flips, signature capture, no double counting on re-scan, thread growth, multi-recipient threads, `MyDomains`, auto-replies, bounces, opt-outs;
- **the real classifier transport** (`crmClassifyBatch_`) with a fake API: request shape, header and model, unexpected-email filtering, coercion, retry/backoff, refusal, truncation, 401, malformed JSON;
- **evidence-pack construction**, **candidate selection** (Lock/MANUAL/mode/stamps), **result application**;
- **dedupe** across all seven statuses including the fuzzy same-person path;
- **audit** on a seeded table with known violations and on a clean one;
- **the whole scan driver** (`crmStartScan` → repeated `crmScanStep`) against a fake spreadsheet, fake `GmailApp.search`, fake properties and triggers: window bounds, multi-tick resume, phase 2 reply pickup, second-run quick-skip with identical counts, trigger removal, and Stop-during-job;
- **merge-by-key save** with a sheet the user sorted, edited, deleted from, and locked while the job ran; **commit guard** for stopped and replaced jobs.

### Manual acceptance checklist (first run in a real Sheet)
1. `Set up / repair CRM tabs` creates all seven tabs; Settings has every key.
2. `Preview 5` returns five rows with non-empty `evidence`; labels read right given `CompanyDescription`.
3. `Scan new mail` on a small `ScanSince` window (e.g. last 30 days): Master rows match what you see in Sent; a known OOO shows `AutoReplies = 1`, `ReceivedCount = 0`.
4. Run the same scan again: no count changes (check `Log` → `SCAN_DONE` and Dashboard totals).
5. `Audit Master` reports 0 integrity issues after a fresh scan + classify.
6. Paste a list with one known contact, one colleague at a known company, and one repeat row: statuses are `DUPLICATE`, `SAME_COMPANY`, `DUPLICATE_IN_IMPORT`; `Push NEW` adds only the new rows to `Contacts` and a second push adds nothing.

## 7. Build phases

| Phase | Deliverable | Status |
|---|---|---|
| **1. Foundation** | `Crm.gs`: menu, tab setup, settings, normalization helpers, API-key storage, logging | built |
| **2. Scan** | `CrmScan.gs`: resumable incremental sent-mail scan → Master + Interactions, snippets, NeedsReply | built |
| **3. Classify** | `CrmClassify.gs`: Rules engine + batched Claude structured-output classification, Lock respect | built |
| **4. Pipeline** | `CrmPipeline.gs`: stage recompute, dormant overlay, mail-merge sync, dashboard refresh | built |
| **5. Import/dedupe** | `CrmImport.gs`: Drive import, header mapping, dedupe statuses + recommendations, push to Contacts | built |
| **6. UI & docs** | `CrmSidebar.html`, `CrmSidebarApi.gs`, `appsscript.json`, README, node smoke tests for pure helpers | built |
| 7. Rigor pass | evidence pack, review queue, high-effort retry, auto-reply/opt-out handling, fuzzy dedupe, audit, preview, windowed two-phase scan, merge-by-key saves, commit guard, 32-group test suite | built |
| 8. Later | see §9 | — |

---

## 8. Cost & limits (so nothing surprises you)

- **Gmail read quota**: Apps Script allows ~20k Gmail read calls/day on consumer accounts; a 2,000-thread first scan uses roughly 2,000–4,000. Incremental scans are tiny.
- **Claude cost**: each batch of 8 contacts is ~1.5–2.5k input tokens and ~1k output tokens. At Opus 5 pricing, 500 contacts ≈ 60 requests ≈ well under $2. `Model` can be set to `claude-sonnet-5` to cut that further.
- **Execution cap**: 6 min per run — handled by the step/trigger design; 90 min/day total trigger runtime on consumer accounts, which is plenty for daily incremental scans.
- **Privacy**: excerpts (≤ 400 chars each, two per contact) plus names/emails/subjects are sent to Anthropic. Set `SnippetChars` to 0 to classify from metadata only.

---

## 9. Later phases (not built yet)

1. **Inbound-only contacts** — scan `in:inbox` for people who wrote first (cold inbound, warm intros) with a stricter noise filter.
2. **Gmail labels write-back** — apply `CRM/Investor`, `CRM/Needs reply` labels so the inbox itself reflects the pipeline.
3. **Weekly digest email** — "5 investors waiting on you, 12 warm contacts going dormant".
4. **Enrichment** — company size/funding from a web source; LinkedIn URL guessing.
5. **Multi-user** — `Owner` column, per-teammate `MyDomains`, shared Master across founders' inboxes.
6. **Segment export** — one click from a Dashboard cell (e.g. "Investors, Dormant") into a mail merge `Contacts` tab for a re-engagement campaign.
7. **Calendar cross-check** — mark `Meeting` from Calendar events with the contact as attendee instead of relying on the AI reading "let's meet Tuesday".
