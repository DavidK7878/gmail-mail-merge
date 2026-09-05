# Email CRM for Google Sheets

Reads your **sent** Gmail, builds a **Master** sheet of every person you have emailed, labels each one with Claude (Investor, Engineer, Customer, …), tracks the relationship stage, and de-duplicates any new sourcing list against it before it goes into the [mail merge](../README.md).

Read **[PLAN.md](PLAN.md)** for the full feature list, data model, and roadmap.

## What's in this folder

| File | Purpose |
|---|---|
| `Crm.gs` | Menu, tab setup, Settings, API-key storage, shared helpers (email normalization, excerpts, table load/save) |
| `CrmScan.gs` | Resumable, incremental scan of `in:sent` → `Master` + `Interactions` |
| `CrmClassify.gs` | `Rules` tab matching, then batched Claude classification with structured JSON output |
| `CrmPipeline.gs` | Stage computation, Dormant overlay, sync from the mail merge `Contacts` tab |
| `CrmImport.gs` | Import from Drive, dedupe against Master, push NEW rows to `Contacts` |
| `CrmSidebarApi.gs` / `CrmSidebar.html` | Sidebar backend + UI |
| `appsscript.json` | Manifest (scopes) |
| `tests/run.js` | Node test suite: pure helpers plus the real scan and classifier transport against fake Gmail / fake API (`node crm/tests/run.js`) |

## Setup (~10 minutes)

### Option A — its own spreadsheet (simplest)

1. Create a new Google Sheet ([sheets.new](https://sheets.new)) → **Extensions → Apps Script**.
2. Replace `Code.gs` with this folder's `Crm.gs`. Add script files named `CrmScan`, `CrmClassify`, `CrmPipeline`, `CrmImport`, `CrmSidebarApi` and paste in the matching files. Add an HTML file named `CrmSidebar` and paste in `CrmSidebar.html`.
3. **Project Settings (⚙️) → Show "appsscript.json"**, paste in this folder's manifest, set `timeZone` to yours. Save.
4. Back in the Sheet, reload. **CRM → Set up / repair CRM tabs** and approve the permission prompts. This creates `Master`, `Interactions`, `Import`, `Rules`, `Settings`, `Dashboard`, `Log`.
5. In **Settings**, fill `CompanyDescription` (one paragraph about what you sell, to whom, and what you are raising). Add co-founder addresses or your company domain to `MyDomains` so teammates on Cc are not treated as contacts.
6. **CRM → Set Claude API key…** and paste an Anthropic API key. It is stored in Script Properties, not in a cell. Without it, only the `Rules` tab classifies.
7. **CRM → Open sidebar → Scan new mail** (or **Full rescan**). The first scan of a year of mail takes a few minutes of background ticks; you can close the tab. When it finishes, classification starts automatically (`AutoClassifyAfterScan`).

### Option B — alongside the mail merge in one spreadsheet

Paste the same files into the mail merge's Apps Script project. There are no name collisions (everything is `crm`-prefixed) **except `onOpen`**:

1. Delete the `onOpen` function from `Crm.gs`.
2. In the mail merge's `Code.gs`, add one line at the end of `onOpen`:
   ```js
   buildCrmMenu_(SpreadsheetApp.getUi());
   ```
3. Merge the two manifests' `oauthScopes` (union of both lists; keep the Gmail advanced service block from the mail merge). Replace `spreadsheets.currentonly` with `spreadsheets`.
4. The CRM's `Settings` and `Log` tabs are shared with the mail merge — setup only appends the missing keys.

With Option B, `Push NEW to Mail Merge` writes straight into the `Contacts` tab next door. With Option A, set `Settings → MailMergeSheetUrl` to the mail merge spreadsheet's URL.

## Daily use

| You want to… | Do this |
|---|---|
| Catch up on new conversations | **Scan new mail** (incremental; only threads since the last scan) |
| See who is waiting on you | Filter `Master` by `NeedsReply = TRUE`, or check the Dashboard |
| Label new contacts | Happens automatically after a scan; or **Classify unclassified** |
| Check labels before spending | **Preview 5** in the sidebar (or the menu) — classifies five contacts, shows evidence, saves nothing |
| Work the review queue | Filter `Master` by `ReviewNeeded = TRUE`; read `Evidence`, fix, tick **`Lock`** |
| Fix a wrong label | Edit `Category` / `Stage` / `Company` by hand and tick **`Lock`** so it is never overwritten |
| Check the data is sound | **Audit Master** — duplicates, stage drift, low-confidence rows not flagged, opt-outs un-suppressed, and more |
| Never email someone again | Tick `DoNotContact` — imports will recommend Skip |
| Add a fresh sourcing list | Paste into `Import` (needs an `Email` header) or **Import from Drive file…**, then **Check Import against Master** |
| Send only the new people | **Push NEW to Mail Merge** — appends to `Contacts`, marks them `PushedAt`, registers them in Master as `Stage = New` |
| Pull bounces/replies from a campaign | **Sync from mail merge Contacts** |

### Reading the Import tab after a check

| `DedupeStatus` | Meaning | Typical `Recommendation` |
|---|---|---|
| `NEW` | Never seen this address or company | Send |
| `DUPLICATE` | Same email already in Master | Skip (in conversation / emailed recently / queued / do not contact) or Re-engage |
| `QUEUED` | Already sitting in the mail merge `Contacts` tab | Skip |
| `POSSIBLE_MATCH` | Same person, different address (same normalized name **and** same company or domain) | Review — names the other address |
| `SAME_COMPANY` | Someone else at that domain is in Master | Review — names the colleague and their stage |
| `DUPLICATE_IN_IMPORT` | Repeated within the list | Skip |
| `INVALID` | No parseable email | Fix email |

Freemail domains (gmail, outlook, …) never trigger `SAME_COMPANY`.

## How classification works

1. **Rules** (`Rules` tab) run first and cost nothing: `domain` (`a16z.com`, or a suffix like `.vc`), `email`, or `keyword` (matched against subjects, excerpts, signature, title). `Category = IGNORE` drops the contact from scans entirely. A rule hit writes `Evidence = Rule domain:a16z.com`.
2. Everything else goes to Claude in batches of `ClassifyBatchSize` with an **evidence pack**: name, email, domain, counts (auto-replies separated), dates, who spoke last, the **signature block** from their latest message, and the last `EvidenceThreads` threads (subject, direction, our message, their message). The response is constrained to a JSON schema and must include a short **`evidence`** quote per contact.
3. First pass runs at `effort: low`. Any result that is *thin* (confidence below `MinConfidence`, category `Other`, relationship `unclear`, or no evidence) is re-run **alone at `effort: high`**; the higher-confidence answer wins. Still thin → `ReviewNeeded = TRUE`.
4. The request uses `claude-opus-5` by default (`Settings → Model`) and Anthropic's server-side `fallbacks: "default"`, so a request declined by a safety classifier is re-run on a fallback model instead of leaving the row blank. Every call's token usage lands in `Log`.
5. Model output is treated as untrusted: entries for emails not in the batch are dropped, enums coerced, confidence clamped, and refusals / truncation / malformed JSON raise and flag the batch for review.

Set `SnippetChars` to `0` to classify from metadata only (nothing from message bodies leaves your account).

## Auto-replies and opt-outs

Out-of-office and vacation responders are detected by subject and by `Auto-Submitted` / `X-Autoreply` / `Precedence` headers. They count in `AutoReplies`, never in `ReceivedCount`, and never set `NeedsReply`. An inbound message matching a conservative opt-out phrase list sets `DoNotContact = TRUE` and tags the row `optout`; imports then recommend Skip.

## Stages

`New → Contacted → Replied → Engaged → Meeting → Won`, with `Dormant` (warm contact, no touch in `DormantDays`), `Not Interested`, and `Bounced` overlays. See PLAN.md §2.3 for the exact rules. Hand-set `Won` / `Not Interested` survive recomputes even without `Lock`.

## Known limits

- **Relationships you started.** People who wrote to you first and never got a reply are not in Master (planned, see PLAN.md §9). Replies on threads you started are picked up, including late replies on old threads.
- **Snippets are heuristics.** Quoted-history stripping handles Gmail, Outlook, French, German, and Spanish reply headers and `--` / sign-off signatures; exotic clients may leak a line of quoted text into an excerpt. Opt-out detection is deliberately narrow, so a politely worded decline is left to the classifier (`relationship = not_interested`).
- **6-minute executions.** Every job is chunked and resumed by a 1-minute trigger; a very large first scan just takes more ticks. Consumer accounts get 90 minutes of trigger time per day.
- **Editing while a job runs is safe** (edits, sorts, deletions, and Locks made mid-job are preserved), with one exception: if you and the job change the *same cell* at the same moment, the job's value wins unless you also ticked `Lock` on that row.
- **Gmail read quota** is ~20k calls/day on consumer accounts. A 2,000-thread first scan uses a fraction of that.
- **One inbox.** `Master` reflects the account that authorized the script. Team-wide CRM is a later phase.

## Troubleshooting

- **No "CRM" menu** → reload the Sheet; check the Apps Script editor for a red syntax banner.
- **"Settings tab missing"** → run **CRM → Set up / repair CRM tabs**.
- **Classification never starts** → sidebar shows "no key": set the API key. `Log` shows `CLASSIFY_ERROR` with the API's message for 401/400s.
- **Everything is `Other`** → fill `Settings → CompanyDescription`; add `Rules` for your main investor/customer domains.
- **Teammates appear as contacts** → add your company domain or their addresses to `Settings → MyDomains`, then **Full rescan**.
- **A scan seems stuck** → **CRM → Stop scan**, then start again; state is in Script Properties and is cleared by Stop.
