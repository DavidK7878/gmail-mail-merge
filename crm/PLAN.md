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
- Searches `in:sent after:<since> before:<scan start>` so the result set is stable while paging.
- **Incremental**: remembers the last completed scan and only looks at threads newer than that (minus a 2-day overlap). **Full rescan** goes back to `Settings → ScanSince`.
- **Resumable**: state (query, offset, counters) lives in Script Properties; a 1-minute trigger keeps calling `crmScanStep` until done, each step stopping at ~4.5 minutes. Close the Sheet and it keeps going.
- For each thread it reads every message and decides direction by comparing the sender to `MyDomains` / your own address:
  - **Outbound** → every To/Cc recipient becomes/updates a contact: `SentCount`, `FirstSentAt`, `LastSentAt`, `Subjects`, `LastOutboundSnippet`, `ThreadIds`.
  - **Inbound** → the sender's contact gets `ReceivedCount`, `LastReplyAt`, `LastInboundSnippet`.
  - Thread whose last message is inbound → that sender's `NeedsReply = TRUE`.
- Excerpts are the first ~400 chars of the plain body **with quoted history stripped** (`On … wrote:` blocks, `>` lines, signatures after `--`).
- Skips machine addresses (`noreply@`, `mailer-daemon`, `calendar-notification@google.com`, …), anything on `IgnoreDomains`, and anything a `Rules` row marks `IGNORE`.
- Writes one row per thread to `Interactions` (subject, dates, counts, last direction, Gmail link) so you can click straight into a conversation.

### 2.2 Classification (`CrmClassify.gs`)
1. **Rules first** (the `Rules` tab): `domain` (exact or suffix like `.vc`), `email`, or `keyword` (matched against subjects, excerpts, title) → Category and optional Company. First match wins, costs nothing, and `IGNORE` drops the contact.
2. **Claude second** for everything still unclassified. Contacts are batched (default 8 per request) and sent as compact JSON: name, email, domain, counts, subjects, both excerpts. The request uses **structured output** (a JSON schema) so the reply is always parseable, `effort: low` because this is classification, and `fallbacks: "default"` so a safety-classifier decline is re-run server-side instead of leaving a row blank.
3. Output per contact: `category`, `confidence` (0–1), `company`, `title`, `first_name`, `last_name`, `relationship` (no_reply / replied / engaged / meeting_scheduled / customer / not_interested / unclear), `sentiment` (positive / neutral / negative), `summary` (one sentence), `next_action` (one short imperative).
4. **`Lock = TRUE`** on a Master row freezes Category / Stage / Company / Title / Name against both rules and AI. Edit a cell by hand and tick Lock to keep it.
5. Runs as a resumable trigger job like the scan (`crmClassifyStep`). Optionally auto-starts when a scan finishes (`AutoClassifyAfterScan`).
6. `Settings → CompanyDescription` is injected into the system prompt so "customer" means *your* customer.

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

**Sync from mail merge** (`crmSyncFromMailMerge`): reads the mail merge `Contacts` tab (same spreadsheet, or the sheet at `MailMergeSheetUrl`), creates missing Master rows with `Source = MailMerge`, and copies `Bounced` / `Replied` flags in.

### 2.4 Import & dedupe (`CrmImport.gs`)
- `Import` tab: paste rows with at least an `Email` header, or use **Import from Drive file…** (Google Sheet URL/ID or CSV) which maps common header variants (`E-mail`, `Full Name`, `Organization`, `Role`…).
- **Check Import against Master** fills, per row:
  - `DedupeStatus`: `NEW`, `DUPLICATE` (same normalized email), `SAME_COMPANY` (same non-freemail domain), `DUPLICATE_IN_IMPORT`, `INVALID`.
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
| AI / rules | `Category`, `Stage`, `Confidence`, `Sentiment`, `AISummary`, `NextAction`, `ClassifiedBy` (RULE / AI / MANUAL), `ClassifiedAt` |
| Yours to edit | `Lock`, `Tags`, `Notes`, `DoNotContact` |
| Facts from Gmail | `Source`, `FirstSentAt`, `LastSentAt`, `LastReplyAt`, `LastTouchAt`, `SentCount`, `ReceivedCount`, `ThreadCount`, `NeedsReply`, `Bounced`, `Subjects`, `LastOutboundSnippet`, `LastInboundSnippet`, `ThreadIds`, `UpdatedAt` |

Extra columns you add to the right are preserved (the code addresses columns by header, like the mail merge does).

### `Interactions` (one row per thread)
`ThreadId`, `Email`, `Subject`, `FirstMessageAt`, `LastMessageAt`, `Messages`, `Outbound`, `Inbound`, `LastDirection`, `LastFrom`, `LastSnippet`, `Link`

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

## 6. Build phases

| Phase | Deliverable | Status |
|---|---|---|
| **1. Foundation** | `Crm.gs`: menu, tab setup, settings, normalization helpers, API-key storage, logging | built |
| **2. Scan** | `CrmScan.gs`: resumable incremental sent-mail scan → Master + Interactions, snippets, NeedsReply | built |
| **3. Classify** | `CrmClassify.gs`: Rules engine + batched Claude structured-output classification, Lock respect | built |
| **4. Pipeline** | `CrmPipeline.gs`: stage recompute, dormant overlay, mail-merge sync, dashboard refresh | built |
| **5. Import/dedupe** | `CrmImport.gs`: Drive import, header mapping, dedupe statuses + recommendations, push to Contacts | built |
| **6. UI & docs** | `CrmSidebar.html`, `CrmSidebarApi.gs`, `appsscript.json`, README, node smoke tests for pure helpers | built |
| 7. Later | see §8 | — |

---

## 7. Cost & limits (so nothing surprises you)

- **Gmail read quota**: Apps Script allows ~20k Gmail read calls/day on consumer accounts; a 2,000-thread first scan uses roughly 2,000–4,000. Incremental scans are tiny.
- **Claude cost**: each batch of 8 contacts is ~1.5–2.5k input tokens and ~1k output tokens. At Opus 5 pricing, 500 contacts ≈ 60 requests ≈ well under $2. `Model` can be set to `claude-sonnet-5` to cut that further.
- **Execution cap**: 6 min per run — handled by the step/trigger design; 90 min/day total trigger runtime on consumer accounts, which is plenty for daily incremental scans.
- **Privacy**: excerpts (≤ 400 chars each, two per contact) plus names/emails/subjects are sent to Anthropic. Set `SnippetChars` to 0 to classify from metadata only.

---

## 8. Later phases (not built yet)

1. **Inbound-only contacts** — scan `in:inbox` for people who wrote first (cold inbound, warm intros) with a stricter noise filter.
2. **Gmail labels write-back** — apply `CRM/Investor`, `CRM/Needs reply` labels so the inbox itself reflects the pipeline.
3. **Weekly digest email** — "5 investors waiting on you, 12 warm contacts going dormant".
4. **Enrichment** — company size/funding from a web source; LinkedIn URL guessing.
5. **Multi-user** — `Owner` column, per-teammate `MyDomains`, shared Master across founders' inboxes.
6. **Segment export** — one click from a Dashboard cell (e.g. "Investors, Dormant") into a mail merge `Contacts` tab for a re-engagement campaign.
7. **Calendar cross-check** — mark `Meeting` from Calendar events with the contact as attendee instead of relying on the AI reading "let's meet Tuesday".
