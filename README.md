# Outlook Cloud Export

A browser bookmarklet that exports all emails from a selected Outlook folder to a JSON file — including sender, recipients, date, full body text, and attachment names.

Works with **https://outlook.cloud.microsoft/** (Microsoft 365 / Outlook on the web).

---

## What you get

For every email in the selected folder the export captures:

| Field | Example |
|---|---|
| Subject | `Q2 budget review` |
| From (name + email) | `Finance Team <finance@example.com>` |
| To (list) | `Jane Smith` |
| CC (list) | *(empty if none)* |
| Date & time | `Wed 5/20/2026 8:06 AM` / ISO timestamp |
| Full body text | complete plain-text of the email |
| Attachment names | `budget-2026-q2.pdf` |
| Attachment sizes | `282 KB` |

> **Attachment files are not downloaded.** Only names and sizes are captured — Outlook web does not expose download links in the page for scripts to use. If you need the actual files, download them manually from each email.

---

## Is it safe?

- The script runs entirely in **your own browser**, on your own Outlook session.
- It reads only what you can already see — it changes nothing.
- Nothing is sent anywhere; the JSON file is saved directly to your computer.
- No login credentials or tokens are ever read or stored.

---

## What you need

- Google Chrome (or any Chromium browser)
- Access to https://outlook.cloud.microsoft/

---

## Step 1 — Install the bookmarklet

1. Open Chrome and show the **Bookmarks Bar** (View → Always Show Bookmarks Bar, or Ctrl+Shift+B / Cmd+Shift+B).
2. Open **`outlook-export.bookmarklet.txt`** from this folder. Copy the long line that starts with `javascript:` under **BOOKMARKLET: FULL EXPORT**.
3. Right-click the bookmarks bar → **Add page…** (or **Add new bookmark**).
4. Give it a name like `Outlook Export`.
5. **Delete** whatever is in the URL/Address field and **paste** the `javascript:...` line you copied.
6. Save.

---

## Step 2 — Run the export

1. Go to **https://outlook.cloud.microsoft/** and sign in.
2. **Click the folder** you want to export in the left sidebar — make sure it is highlighted.
3. Make sure the Outlook tab is **visible and in the foreground** (do not switch to another tab while it runs — Chrome slows down background tabs).
4. Click the **Outlook Export** bookmarklet.
5. Watch the DevTools console (F12 → Console) for progress — it will show each email being processed.
6. When done, a **JSON file downloads automatically** to your Downloads folder.

Large folders (hundreds of emails) can take several minutes — one email at a time must be opened to read the full body.

---

## Step 3 — Open the JSON file

The exported file (e.g. `outlook-export-Inbox-2026-05-30T12-00-00.json`) is a standard JSON file. You can:

- Open it in any text editor (VS Code, Notepad++)
- Import it into Excel via **Data → Get Data → From JSON**
- Use it as input for other scripts or tools

---

## Troubleshooting

**"0 emails found"**
Make sure a folder is selected in the left sidebar before clicking the bookmarklet. The folder name should be highlighted.

**Script seems stuck**
Keep the Outlook tab visible. If you switched away, switch back — Chrome freezes timers in background tabs. You can also abort with `window.__OUTLOOK_EXPORT_STOP = true` in the DevTools console (F12).

**Some emails missing**
The script scrolls the email list to load all items before starting. If your folder has thousands of emails this can take a while. Try the test bookmarklet first (exports first 10 emails only) to confirm it works.

**Reading pane did not load**
Increase `emailLoadTimeoutMs` in the CONFIG if your connection is slow. Paste this in the console before running:
```javascript
window.__OUTLOOK_EXPORT_OVERRIDE = { emailLoadTimeoutMs: 20000 };
```

---

## For developers

Main script: **`outlook-export.js`** — self-contained IIFE, no dependencies.

Build the bookmarklet:
```bash
node build-bookmarklet.js
```

Override any CONFIG value at runtime:
```javascript
window.__OUTLOOK_EXPORT_OVERRIDE = { maxEmails: 5, verbose: true };
// then click the bookmarklet
```

Selectors are in the `SEL` block near the top of `outlook-export.js`. Microsoft ships DOM changes regularly — if something breaks, re-check those selectors in DevTools.

---

## License

MIT
