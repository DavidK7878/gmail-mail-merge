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
| `tests/run.js` | Node smoke tests for the pure helpers (`node crm/tests/run.js`) |

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
| Fix a wrong label | Edit `Category` / `Stage` / `Company` by hand and tick **`Lock`** so it is never overwritten |
| Never email someone again | Tick `DoNotContact` — imports will recommend Skip |
| Add a fresh sourcing list | Paste into `Import` (needs an `Email` header) or **Import from Drive file…**, then **Check Import against Master** |
| Send only the new people | **Push NEW to Mail Merge** — appends to `Contacts`, marks them `PushedAt`, registers them in Master as `Stage = New` |
| Pull bounces/replies from a campaign | **Sync from mail merge Contacts** |

### Reading the Import tab after a check

| `DedupeStatus` | Meaning | Typical `Recommendation` |
|---|---|---|
| `NEW` | Never seen this address or company | Send |
| `DUPLICATE` | Same email already in Master | Skip (in conversation / emailed recently / do not contact) or Re-engage |
| `SAME_COMPANY` | Someone else at that domain is in Master | Review — names the colleague and their stage |
| `DUPLICATE_IN_IMPORT` | Repeated within the list | Skip |
| `INVALID` | No parseable email | Fix email |

Freemail domains (gmail, outlook, …) never trigger `SAME_COMPANY`.

## How classification works

1. **Rules** (`Rules` tab) run first and cost nothing: `domain` (`a16z.com`, or a suffix like `.vc`), `email`, or `keyword` (matched against subjects, excerpts, title). `Category = IGNORE` drops the contact from scans entirely.
2. Everything else goes to Claude in batches of `ClassifyBatchSize`. Each contact is sent as compact JSON: name, email, domain, counts, dates, up to six subjects, and the latest outbound and inbound excerpts (each ≤ `SnippetChars`, quoted history stripped). The response is constrained to a JSON schema, so it always parses.
3. The request uses `claude-opus-5` by default (`Settings → Model`), `effort: low` (this is classification), and Anthropic's server-side `fallbacks: "default"` so a request declined by a safety classifier is re-run on a fallback model instead of leaving the row blank. Token usage for every call lands in `Log`.
4. Results fill `Category`, `Confidence`, `Company`, `Title`, `Sentiment`, `AISummary`, `NextAction`, and a `rel:` tag in `Tags` that feeds the stage.

Set `SnippetChars` to `0` to classify from metadata only (nothing from message bodies leaves your account).

## Stages

`New → Contacted → Replied → Engaged → Meeting → Won`, with `Dormant` (warm contact, no touch in `DormantDays`), `Not Interested`, and `Bounced` overlays. See PLAN.md §2.3 for the exact rules. Hand-set `Won` / `Not Interested` survive recomputes even without `Lock`.

## Known limits

- **Sent mail only.** People who wrote to you first and never got a reply are not in Master (planned, see PLAN.md §8).
- **Snippets are heuristics.** Quoted-history stripping handles Gmail, Outlook, French, and German reply headers; exotic clients may leak a line of quoted text into an excerpt.
- **6-minute executions.** Every job is chunked and resumed by a 1-minute trigger; a very large first scan just takes more ticks. Consumer accounts get 90 minutes of trigger time per day.
- **Gmail read quota** is ~20k calls/day on consumer accounts. A 2,000-thread first scan uses a fraction of that.
- **One inbox.** `Master` reflects the account that authorized the script. Team-wide CRM is a later phase.

## Troubleshooting

- **No "CRM" menu** → reload the Sheet; check the Apps Script editor for a red syntax banner.
- **"Settings tab missing"** → run **CRM → Set up / repair CRM tabs**.
- **Classification never starts** → sidebar shows "no key": set the API key. `Log` shows `CLASSIFY_ERROR` with the API's message for 401/400s.
- **Everything is `Other`** → fill `Settings → CompanyDescription`; add `Rules` for your main investor/customer domains.
- **Teammates appear as contacts** → add your company domain or their addresses to `Settings → MyDomains`, then **Full rescan**.
- **A scan seems stuck** → **CRM → Stop scan**, then start again; state is in Script Properties and is cleared by Stop.
