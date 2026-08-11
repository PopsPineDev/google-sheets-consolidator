# google-sheets-consolidator

**Consolidate daily Google Sheets reports into one master sheet, automatically.** Built for AI agents and automations that can write files to Drive but can't POST to a webhook.

Two small Apps Script files. No servers, no dependencies, no OAuth app to register.

---

## The problem

You have an automation — a scheduled AI agent, a cloud job, a research bot — that produces a report every day. You want all those reports to land as dated tabs in **one** spreadsheet you can open on your phone.

The obvious design is a webhook: the agent POSTs rows to an Apps Script web app. Except many cloud agent runtimes have an outbound network allowlist, and `script.google.com` isn't on it. The POST dies with `CONNECT tunnel failed, 403` and there's nothing you can do about it from inside the sandbox.

But those same agents usually *can* write files to Google Drive, because that goes through an authenticated connector rather than raw HTTP.

## The fix

Flip the direction. The agent drops a spreadsheet into a Drive folder. A time-driven trigger **inside your own Google account** picks it up and files it.

```
agent ──writes file──> Drive folder ──timer pulls──> master sheet (dated tabs)
                                                          └─ source file trashed
```

Nothing has to reach *in* to your account, so there is no network path to block, no endpoint to secure, and no token to leak.

## What it does

- Scans an inbox folder for spreadsheets whose name starts with a prefix you choose
- Copies each into the master sheet as a new tab, named from the rest of the filename (`Daily Report — Aug 11, 2026` → tab `Aug 11, 2026`), newest tab first, header bolded and frozen
- Trashes the source file afterwards so the folder doesn't fill up
- Remembers what it already imported, so re-running is harmless

## Quick start (~3 minutes)

1. Create the master spreadsheet and the inbox folder in Drive.
2. In the master sheet: **Extensions → Apps Script**. Paste in [`src/Consolidator.gs`](src/Consolidator.gs).
3. Edit `CONFIG` at the top:

   ```js
   MASTER_SHEET_ID: '1AbC...',       // from the sheet URL: /spreadsheets/d/THIS/edit
   INBOX_FOLDER_ID: '1XyZ...',       // from the folder URL: /drive/folders/THIS
   NAME_PREFIX: 'Daily Report —',    // only files starting with this are imported
   ```

4. Pick `consolidateNow` in the function dropdown → **Run**. Authorize when prompted (it's your own script, so the "hasn't been verified" screen is expected — Advanced → Go to project).
5. Pick `installDailyTrigger` → **Run**. Done — it now runs every day on its own.

The execution log tells you what happened: `Imported 1 file(s); trashed 1.`

### Telling your agent where to write

Point it at the folder and the naming convention. For example:

> Save the report as a Google Sheet named `Daily Report — <today's date, e.g. Aug 11, 2026>` inside the Drive folder with ID `1XyZ...`.

Anything that lands there in the right shape gets filed automatically.

## Safety design

Deleting things automatically deserves care, so the ordering is deliberate:

- **`SpreadsheetApp.flush()` before trashing.** Apps Script batches writes. Flushing forces the tab to commit before the source file is touched, so a file is never removed on the strength of a write that hasn't landed.
- **Trash, not delete.** `setTrashed(true)` is recoverable from Drive Trash for about 30 days.
- **Failures keep their file.** The trash call sits inside the `try`; if an import throws, the source stays in the folder and is retried next run.
- **The master is protected twice** — skipped by file ID *and* by name prefix — so it can safely live inside the inbox folder itself.

Prefer to keep everything? Set `TRASH_AFTER_IMPORT: false`.

## Options

| Setting | Default | What it does |
|---|---|---|
| `NAME_PREFIX` | `'Daily Report —'` | Only files starting with this are imported |
| `TRASH_AFTER_IMPORT` | `true` | Trash each source after a verified import |
| `RUN_AT_HOUR` | `4` | Hour of day for the daily trigger (script timezone) |
| `IMPORT_ALL_TABS` | `false` | Import every tab of a multi-tab source, not just the first |

Schedule it after your agent runs — if the agent writes at 3 AM, leaving the default 4 AM gives it an hour of slack.

## Optional: the webhook

[`src/Webhook.gs`](src/Webhook.gs) is the push half, for callers that *can* make outbound requests — your own VPS, a CI job, a phone shortcut, a Zapier step. It appends rows to a dated tab directly.

The shared token is read from **Script Properties**, never from the file. Generate one with `openssl rand -hex 24`, then Project Settings → Script properties → `SHARED_TOKEN`.

> **Never commit a token or a deployed `/exec` URL to a public repo.** Git history keeps it even after you delete the line, and public repos are scraped continuously.

## Requirements

A Google account. That's it. Apps Script quotas are generous — one small file per day isn't close to any limit.

## License

MIT — see [LICENSE](LICENSE).
