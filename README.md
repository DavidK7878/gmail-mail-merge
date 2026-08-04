# Mail Merge for Google Sheets

A YAMM-style mail merge that lives inside a Google Sheet, sends through your Gmail, and tracks opens, clicks, replies, and bounces — with staggered sending, per-recipient attachments, and automatic follow-ups.

**New here? Start with [TESTING.md](TESTING.md)** — a 15-minute walkthrough from blank spreadsheet to a merged test email in your own inbox.

## What's in this folder

| File | Purpose |
|---|---|
| `Code.gs` | Menu, sheet setup, shared helpers |
| `Send.gs` | Merge engine, staggered batch sending, quota/send-window enforcement, test send |
| `Tracking.gs` | Web App endpoint (`doGet`), open pixel, click redirect |
| `Replies.gs` | Reply + bounce detection (30-min trigger) |
| `FollowUps.gs` | Automatic follow-up sequences (daily trigger, threaded replies) |
| `SidebarApi.gs` | Backend for the sidebar |
| `Sidebar.html` | Sidebar UI (pick draft, preview, start/pause) |
| `appsscript.json` | Script manifest (scopes, Gmail advanced service, web app config) |

## One-time setup (~10 minutes)

1. **Create a new Google Sheet** (sheets.new) and open **Extensions → Apps Script**.
2. In the Apps Script editor:
   - Replace the default `Code.gs` content with this folder's `Code.gs`.
   - Add script files (＋ → Script) named `Send`, `Tracking`, `Replies`, `FollowUps`, `SidebarApi` and paste in the matching `.gs` files.
   - Add an HTML file (＋ → HTML) named `Sidebar` and paste in `Sidebar.html`.
   - Open **Project Settings (⚙️) → check "Show appsscript.json manifest file"**, then paste in this folder's `appsscript.json`. Set `timeZone` to yours (e.g. `America/Chicago`) — it controls the send window. Full list: [IANA time zones](https://en.wikipedia.org/wiki/List_of_tz_database_time_zones).
   - Save everything (Ctrl+S).
3. **Authorize:** back in the Sheet, reload the page. A **Mail Merge** menu appears — click **Mail Merge → Set up / repair sheet tabs**. Approve the permission prompts (it's your own account authorizing your own script). This creates the `Contacts`, `Settings`, `Dashboard`, and `Log` tabs.
4. **Deploy tracking (optional but recommended):** in the Apps Script editor, **Deploy → New deployment → Web app**, execute as **Me**, access **Anyone**. Copy the Web App URL and paste it into **Settings → WebAppUrl** in the Sheet. Without this, turn `TrackOpens`/`TrackClicks` off.
5. **Write your template:** compose a Gmail **draft** with `{{FirstName}}`-style placeholders in the subject or body. Any Contacts column header works as a placeholder. Your signature, formatting, and inline images come along for free.

## Running a campaign

1. Fill the **Contacts** tab: `Email` is required; `FirstName`, `Company`, `Attachment`, and any columns you add become merge fields. `Attachment` takes a Drive link or file name (comma-separate for multiple). Leave the machine columns (`Status` onward) alone.
2. Check the **Settings** tab — daily cap, pacing, send window, tracking toggles, follow-up rules.
3. **Mail Merge → Open sidebar**: pick your draft, **Preview row**, then **Send test to myself**. The test merges row 2's data but delivers to your inbox — this is the step that catches placeholder mistakes.
4. Click **Start sending**. Batches go out every ~5 minutes, spaced 30–90 s apart, only inside your send window, up to the daily cap. You can close the Sheet; sending continues. **Pause** stops sending; **Stop everything** also removes reply-checking and follow-up triggers.
5. Watch the **Dashboard** tab for sent / opened / clicked / replied / bounced rates, and the **Log** tab for a full event trail.

## Follow-ups

Compose each follow-up as its own Gmail draft, then in **Settings**:

- `FollowUpDraftSubjects` — comma-separated draft subjects, one per stage (last repeats).
- `FollowUpWaitDays` — days to wait after the previous touch.
- `MaxFollowUps` — how many stages (0 disables).

Follow-ups send **as a reply in the original thread** so the recipient sees them in context. Reply detection always runs first in the same pass — nobody who answered gets a follow-up. (Threaded replies use the Gmail advanced service, enabled via the manifest; if unavailable it falls back to a normal "Re:" email.)

## Known limits (by design, stated up front)

- **~100 recipients/day** on consumer @gmail.com (1,500 on Workspace). The scheduler enforces the cap and rolls the remainder to the next allowed day automatically.
- **Open counts are directional, not exact** — Gmail's image proxy and Apple Mail privacy protection can inflate or mask opens. True of every tracker, including YAMM.
- **Click-tracked links display as `script.google.com` URLs.** `TrackClicks` is off by default; turn it on per campaign when click data matters more than pristine links.
- Apps Script executions cap at 6 minutes — the batch/trigger design stays under this automatically; you'll never see it.

## Troubleshooting

- **No "Mail Merge" menu:** reload the Sheet tab; make sure the files saved without syntax errors (Apps Script editor shows a red banner if not).
- **"No Gmail draft found":** the sidebar matches drafts by exact subject — hit ↻ Refresh after editing a draft's subject.
- **Opens/clicks not recording:** confirm the Web App is deployed with access "Anyone" and the URL is pasted in Settings → WebAppUrl. Re-deploy (Deploy → Manage deployments → edit → new version) after any code change to `Tracking.gs`.
- **Follow-ups not threading:** the Gmail advanced service must be enabled (it is if you pasted the manifest; otherwise Services ＋ → Gmail API).
- Every send, open, click, reply, bounce, and error lands in the **Log** tab with a timestamp.
