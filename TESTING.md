# Testing This Tool — A Walkthrough for a New User

You do not need to know JavaScript, Apps Script, or Google Sheets internals. Everything below happens in a browser, in your own Google account. Budget about 15 minutes.

Nothing you do here can email anyone but yourself until the very last optional step.

## What you'll end up with

A Google Sheet that sends personalized emails from your own Gmail, tracks who opened them, detects replies and bounces automatically, and sends follow-ups to people who didn't answer.

## What you need

- A Google account (personal @gmail.com is fine)
- The files in this repository

---

## Step 1 — Create the spreadsheet and open the script editor

1. Go to **[sheets.new](https://sheets.new)**. A blank spreadsheet opens.
2. Rename it something like `Mail Merge` (click the title, top-left).
3. Menu: **Extensions → Apps Script**. A code editor opens in a new tab.

## Step 2 — Paste in the code

The editor starts with one file, `Code.gs`, containing an empty `myFunction()`.

1. Select everything in `Code.gs` (Ctrl+A) and paste in this repo's **`Code.gs`**.
2. For each of `Send`, `Tracking`, `Replies`, `FollowUps`, `SidebarApi`: click the **＋** next to "Files" → **Script** → type the name (no `.gs`) → paste in the matching file from this repo.
3. Click **＋ → HTML**, name it **`Sidebar`**, and paste in **`Sidebar.html`**. Delete the placeholder HTML that's already there first.
4. Show the manifest: **⚙️ Project Settings** → check **"Show 'appsscript.json' manifest file in editor"** → back in the editor, open `appsscript.json` and paste in this repo's version. **Change `timeZone`** to yours (e.g. `America/Chicago`) — it controls send hours.
5. **Ctrl+S** to save.

> **Checkpoint:** the file list should read `appsscript.json`, `Code.gs`, `Send.gs`, `Tracking.gs`, `Replies.gs`, `FollowUps.gs`, `SidebarApi.gs`, `Sidebar.html`. No red error banners.

## Step 3 — Authorize it (the scary-looking screen)

1. Go back to your spreadsheet tab and **reload the page**.
2. A new **Mail Merge** menu appears in the menu bar. Click **Mail Merge → Set up / repair sheet tabs**.
3. Google asks for permission. Click through:
   - **Continue** → choose your account
   - You'll see **"Google hasn't verified this app."** Click **Advanced** (small link, bottom-left), then **"Go to Mail Merge (unsafe)"**
   - **Allow**

   This warning is Google's standard notice for any personal script that hasn't gone through their commercial app-review process. The "developer" it names is *you* — the script runs only in your own account, and no third party is involved.

4. A dialog confirms: *"Mail Merge tabs are ready."*

> **Checkpoint:** four new tabs at the bottom — **Contacts**, **Settings**, **Dashboard**, **Log**.

## Step 4 — Add yourself as the only contact

On the **Contacts** tab, fill row 2:

| Email | FirstName | Company |
|---|---|---|
| *your own address* | your first name | Test Company |

Leave every column from `Status` onward blank — the script owns those.

## Step 5 — Write a template as a Gmail draft

Templates are ordinary Gmail drafts, which is why your signature and formatting come along for free.

1. Open Gmail, click **Compose**.
2. **Subject:** `Quick hello from {{Company}}`
3. **Body:**

   ```
   Hi {{FirstName}},

   This is a test of the mail merge tool — personalized from a spreadsheet row
   and sent through my own Gmail.

   Best,
   [your name]
   ```

4. **Close the compose window** (the X). Do not click Send — it saves as a draft, which is what we want.

Any column header on the Contacts tab works as a `{{placeholder}}`.

## Step 6 — Preview and send yourself a test

1. Back in the Sheet: **Mail Merge → Open sidebar**.
2. Pick your draft in the **Template** dropdown (hit **↻ Refresh** if it's not listed).
3. Click **Preview row** — you should see your merged email with the placeholders filled in. This is the step that catches mistakes.
4. Click **Send test to myself**.

> **Checkpoint:** an email arrives in your inbox, subject `[TEST] Quick hello from Test Company`, greeting you by name. **If the placeholders still read `{{FirstName}}`, your column headers don't match the template — fix and retry.**

**If you only wanted to evaluate the tool, you're done.** Everything below is for running a real campaign.

---

## Step 7 (optional) — Turn on open tracking

Tracking needs the script published as a web endpoint so that opened emails can phone home.

1. In the Apps Script editor: **Deploy → New deployment**.
2. Gear icon → **Web app** (it may be preselected from the manifest).
3. Confirm **Execute as: Me**, **Who has access: Anyone**. Click **Deploy**.
4. Authorize again if prompted, then **copy the Web app URL**.
5. In the Sheet: **Settings** tab → paste the URL into the **`WebAppUrl`** row (column B, last row).

"Anyone" access is required because recipients' mail clients load the tracking pixel without being logged into your account. The endpoint only records opens and clicks — it reads no data out.

**If you skip this step, set `TrackOpens` to `FALSE` in Settings**, or sending will refuse to start.

## Step 8 (optional) — Run a real campaign

1. Add real recipients to the **Contacts** tab, one row each.
2. Review **Settings** — `DailyCap`, send window hours, `SendDays`, pacing.
3. Sidebar → **Send test to myself** one more time, using a real row's data.
4. Sidebar → **Start sending**.

Emails go out in small batches every ~5 minutes, spaced 30–90 seconds apart, only inside your send window, up to the daily cap. **You can close the Sheet — sending continues.** Watch `Status` and `SentAt` fill in on the Contacts tab, rates on **Dashboard**, and a full event trail on **Log**.

**Pause** stops sending. **Mail Merge → Stop everything** also removes the reply-checking and follow-up triggers.

---

## Things that go wrong, and what they mean

| Symptom | Cause / fix |
|---|---|
| No **Mail Merge** menu | Reload the Sheet tab. Check the script saved without errors. |
| `Contacts tab missing` | Run **Mail Merge → Set up / repair sheet tabs**. |
| `No Gmail draft found` | The sidebar matches drafts by **exact subject**. Hit ↻ Refresh after editing a draft. |
| Placeholders not filled | Column header and `{{placeholder}}` must match (case-insensitive, but spelling counts). |
| `Tracking is enabled but WebAppUrl is empty` | Do Step 7, or set `TrackOpens` to `FALSE`. |
| Opens never register | Redeploy after any change to `Tracking.gs` (**Deploy → Manage deployments → ✏️ → New version**). Also: many mail clients block images, so opens undercount. |
| Nothing sends, no error | Outside your send window (Settings), or the daily cap is used up. Check the **Log** tab. |
| Follow-ups aren't threaded | The Gmail advanced service must be enabled — it is if you pasted `appsscript.json`; otherwise **Services ＋ → Gmail API**. |

Every send, open, click, reply, bounce, and error is timestamped on the **Log** tab. Start there.

## Limits worth knowing before you rely on it

- **~100 recipients/day** on a consumer @gmail.com account (1,500 on Google Workspace). The scheduler enforces this and rolls the remainder to the next allowed day.
- **Open counts are directional, not exact.** Image proxying and privacy protection in modern mail clients both inflate and mask opens. This is true of every email tracker.
- **Click tracking rewrites links** to route through `script.google.com`, which is visible to recipients. It's off by default; turn it on per campaign.
- Cold outreach is regulated in most jurisdictions (CAN-SPAM in the US, GDPR in the EU). A real physical mailing address and a working opt-out belong in your template.
