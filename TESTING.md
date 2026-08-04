# Try It Yourself — 15-Minute Setup & Test Walkthrough

This guide takes you from zero to a working mail merge with a test email in your own inbox. You need a Google account (a plain @gmail.com works) and nothing else. Nothing in this guide emails anyone but you.

## 1. Create the Sheet and paste in the code (~5 min)

1. Go to [sheets.new](https://sheets.new) to create a blank spreadsheet. Name it anything (e.g. "Mail Merge").
2. In the menu: **Extensions → Apps Script**. A script editor opens in a new tab.
3. In the editor, copy this repo's files in:
   - Replace the default `Code.gs` contents with this repo's [`Code.gs`](Code.gs).
   - For each of `Send.gs`, `Tracking.gs`, `Replies.gs`, `FollowUps.gs`, `SidebarApi.gs`: click **＋ → Script** next to "Files", name it to match (the editor adds `.gs`), then paste the file's contents.
   - Click **＋ → HTML**, name it `Sidebar`, paste in [`Sidebar.html`](Sidebar.html).
   - **Project Settings (⚙️ icon) → check "Show 'appsscript.json' manifest file in editor"**, then back in the editor open `appsscript.json` and replace it with this repo's [`appsscript.json`](appsscript.json). Change `"timeZone"` to yours ([IANA list](https://en.wikipedia.org/wiki/List_of_tz_database_time_zones)) — it controls the send window.
4. Press **Ctrl+S** to save. If a red error banner appears, a paste went wrong — re-paste that file completely.

## 2. Authorize it (~2 min, looks scarier than it is)

1. Go back to the spreadsheet tab and **reload the page**. After a few seconds a **Mail Merge** menu appears at the end of the menu bar.
2. Click **Mail Merge → Set up / repair sheet tabs**.
3. Google shows **"Authorization required"** → OK → pick your account. You'll then hit **"Google hasn't verified this app"**. This is expected: it's your own private script, not a published app, so Google hasn't reviewed it. Click **Advanced → Go to \<project name\> (unsafe) → Allow**.
4. Run **Mail Merge → Set up / repair sheet tabs** *again* (the first run got eaten by the authorization). You should see "Mail Merge tabs are ready" and four new tabs: **Contacts, Settings, Dashboard, Log**.

> What you just authorized: the script reads your Gmail drafts, sends email as you, edits this one spreadsheet, and reads Drive files you reference as attachments. It runs only in your account and only when you (or its timers) trigger it.

## 3. Write a template and add yourself as a contact (~3 min)

1. In Gmail, compose a new email. Subject: `Merge test for {{company}}`. Body: anything with `{{first}}` in it. Don't fill in "To". Close it — it auto-saves as a draft. **The draft's subject is its identity** — the tool finds it by exact subject.
2. In the **Contacts** tab of the Sheet, row 2: put **your own email** in the `Email` column, your name under `first`, and any text under `company`.
   - Any column header on the left side works as a `{{placeholder}}` (case-insensitive). Add columns as you like. Don't touch `Status` and everything right of it — the script owns those.

## 4. Send yourself the test (~1 min)

1. **Mail Merge → Open sidebar**. Pick your draft in the dropdown, click **Preview row** — you should see your name substituted in.
2. Click **Send test to myself**. Check your inbox: subject prefixed `[TEST]`, placeholders filled with row 2's values, your signature intact. That's the whole merge path working.

## 5. (Optional) Turn on open tracking (~2 min)

1. In the Apps Script editor: **Deploy → New deployment**. Type should already read "Web app" (Execute as **Me**, access **Anyone** — prefilled from the manifest). Click **Deploy**, copy the web app URL.
2. Paste that URL into **Settings → WebAppUrl** in the Sheet.
3. Skip this entirely by setting `TrackOpens` to FALSE in Settings — "Start sending" refuses to run if tracking is on but no URL is set (that's a guard, not a bug).

## 6. Running a real (small!) campaign

1. Add a few more rows to Contacts — **use your own alternate addresses or willing friends for a first run.**
2. Sidebar → pick draft → **Start sending**. Emails go out in batches every ~5 minutes, spaced 30–90 s apart, only inside the send window in Settings (default: weekdays 8am–5pm, max 90/day — leftovers roll to the next allowed day). You can close the Sheet; sending continues.
3. Watch the **Contacts** row statuses (`SENT` + timestamp), the **Dashboard** rates, and the **Log** for every event. **Pause sending** stops the batch trigger; **Stop everything** also removes reply-checking and follow-up triggers.
4. Attachments: put a Drive file name or link in a row's `Attachment` cell, or just attach files to the Gmail draft itself (draft attachments go to everyone).
5. Follow-ups: compose each follow-up as its own draft, list their subjects comma-separated in **Settings → FollowUpDraftSubjects**, set `MaxFollowUps` ≥ 1. Follow-ups send as replies in the original thread and automatically skip anyone who replied or bounced.

## What to check if something misbehaves

| Symptom | Fix |
|---|---|
| No "Mail Merge" menu | Reload the Sheet; check the script editor for a red syntax-error banner |
| "No Gmail draft found with subject…" | The dropdown matches drafts by *exact* subject; hit ↻ Refresh after editing the draft |
| "Tracking is enabled but WebAppUrl is empty" | Do step 5, or set `TrackOpens` FALSE |
| Opens never increment | Re-deploy after any code change (Deploy → Manage deployments → ✏️ → new version); confirm access is "Anyone" |
| Sends stop mid-list | Look at the Log tab — usually the daily cap (`CAP_REACHED`) or outside the send window; both resume automatically |
| Every send fails instantly | Log tab shows the reason — commonly an unresolvable `Attachment` reference |

## Known limits

- Consumer @gmail.com: ~100 recipients/day (Workspace: 1,500). The tool enforces its own `DailyCap` on top.
- Open counts are directional — Gmail's image proxy and Apple Mail privacy features distort them for every tracker.
- Click tracking (off by default) rewrites links through `script.google.com`, which recipients can see.
- Cold-emailing large lists of inferred/unverified addresses is a fast way to get a Gmail account rate-limited or flagged. Keep daily caps modest, watch the bounce column, and stop if bounces spike.
