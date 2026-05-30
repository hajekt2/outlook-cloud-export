/* =============================================================================
 * Outlook Cloud Export  —  outlook.cloud.microsoft folder exporter
 * -----------------------------------------------------------------------------
 * Exports every email in the currently selected Outlook folder to a JSON file
 * by driving the page DOM: it scrolls the email list to capture all messages,
 * clicks each one to open the reading pane, and captures subject, sender,
 * recipients, date, full body text, and attachment metadata.
 *
 * ATTACHMENT NOTE
 *   Attachment file names and sizes are captured. Binary content is NOT
 *   exported (Outlook web does not expose download URLs in the DOM). If you
 *   need the actual files, download them manually from each email.
 *
 * HOW TO RUN
 *   A) DevTools console: paste this whole file while Outlook is open and a
 *      folder is selected. It auto-runs.
 *   B) Bookmarklet: use the minified one-liner in outlook-export.bookmarklet.txt
 *
 * Run with the Outlook tab FOCUSED/VISIBLE. Keep it in the foreground the
 * whole time — Chrome throttles timers in background tabs and Outlook pauses
 * rendering. Large folders take several minutes.
 *
 * To abort early, run in the DevTools console:
 *   window.__OUTLOOK_EXPORT_STOP = true
 *
 * Result is also saved to window.__OUTLOOK_EXPORT_RESULT after completion.
 *
 * SELECTORS verified against outlook.cloud.microsoft (May 2026). Microsoft
 * ships DOM updates regularly; if a run yields 0 emails re-check SEL block.
 * ========================================================================== */

(function () {
  "use strict";

  // ---- CONFIG ---------------------------------------------------------------
  const CONFIG = Object.assign({
    scrollDelayMs: 700,        // wait after each list scroll step
    stableRounds: 5,           // consecutive no-new-email rounds => list fully loaded
    maxScrollIters: 400,       // hard cap on list-scroll steps
    emailLoadTimeoutMs: 10000, // max wait for reading pane after click
    minSettleMs: 1200,         // min settle time after reading pane loads
    betweenEmailsMs: 400,      // extra pause between emails
    rowFindScrollMs: 700,      // wait after scrolling list to find a missing row
    maxEmails: 0,              // 0 = all; >0 = limit (handy for testing)
    download: true,            // auto-download the JSON file when done
    verbose: true,
  }, (window.__OUTLOOK_EXPORT_OVERRIDE || {}));

  // ---- SELECTORS ------------------------------------------------------------
  const SEL = {
    folderItem:     '[role="treeitem"]',
    emailRow:       '[role="option"][data-convid]',
    readingPane:    '.wide-content-host',
    fromBtn:        '[aria-label^="From:"]',
    toHeading:      '[aria-label^="To:"]',
    ccHeading:      '[aria-label^="CC:"], [aria-label^="Cc:"]',
    messageBody:    '[role="document"][aria-label*="Message body"]',
    attachContainer:'[aria-label="file attachments"]',
    attachItem:     '[role="option"]',
  };

  // ---- helpers --------------------------------------------------------------
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const log   = (...a) => CONFIG.verbose && console.log("[outlook-export]", ...a);
  const txt   = (el)  => (el ? (el.innerText || "").trim() : "");
  const norm  = (s)   => (s || "").replace(/[​‌‍﻿]/g, "").replace(/\s+/g, " ").trim();

  // Strip HTML-rendering artifacts from body text: tabs→space, whitespace-only
  // lines→empty, then collapse runs of 3+ newlines down to a single blank line.
  const cleanBody = (s) => (s || "")
    .replace(/[ \t]+/g, " ")     // tabs and repeated spaces → single space
    .replace(/^ +$/gm, "")       // lines that are only spaces → empty
    .replace(/\n{3,}/g, "\n\n")  // more than one blank line → one blank line
    .trim();

  function scrollableAncestor(el) {
    let n = el;
    while (n && n !== document.body) {
      const oy = getComputedStyle(n).overflowY;
      if (/(auto|scroll)/.test(oy) && n.scrollHeight > n.clientHeight + 40) return n;
      n = n.parentElement;
    }
    return null;
  }

  // ---- folder detection -----------------------------------------------------
  function selectedFolderName() {
    // 1. Prefer aria-selected="true" on a treeitem. The innerText often starts
    //    with a newline, so take the first line that contains a word character.
    const sel = document.querySelector('[role="treeitem"][aria-selected="true"]');
    if (sel) {
      const line = (sel.innerText || "").split("\n").map((s) => s.trim()).find((s) => /\w/.test(s));
      if (line) return norm(line);
    }
    // 2. Fall back to page title: "FolderName - User - Outlook"
    //    Only use this at startup before any email is opened (the title switches
    //    to the email subject once a message is clicked).
    const m = document.title.match(/^(.+?)\s*[-–]\s*.+?\s*[-–]\s*Outlook/i);
    if (m) return m[1].trim();
    return "unknown";
  }

  // ---- email row parsing ----------------------------------------------------
  // Row innerText structure (confirmed via DOM probe):
  //   "XPER Consulting AS\n\nSubject line\nShort date\nPreview..."
  // The separator between sender and subject is \n + U+E1B7 (private-use icon) + \n,
  // NOT a plain \n\n. Filter lines to skip icon-only entries.
  // aria-label starts with "Has attachments " when there are attachments.
  function parseRow(row) {
    // Keep only lines that contain at least one real word character.
    const lines = (row.innerText || "").split("\n")
      .map((l) => l.trim())
      .filter((l) => /\w/.test(l));
    // lines[0] = sender display name
    // lines[1] = subject
    // lines[2] = short date ("Wed 5/20")
    // lines[3+] = preview text
    return {
      id:             row.id,
      convid:         row.getAttribute("data-convid"),
      posInSet:       parseInt(row.getAttribute("aria-posinset") || "0", 10),
      senderPreview:  norm(lines[0] || ""),
      subject:        norm(lines[1] || ""),
      dateShort:      norm(lines[2] || ""),
      hasAttachments: /^Has attachments\b/i.test(row.getAttribute("aria-label") || ""),
    };
  }

  // ---- scroll email list to enumerate all rows ------------------------------
  async function loadAllEmailRows() {
    const firstRow = document.querySelector(SEL.emailRow);
    if (!firstRow) { log("ERROR: no email rows found. Is a folder selected?"); return []; }

    const totalStr   = firstRow.getAttribute("aria-setsize") || "0";
    const totalCount = parseInt(totalStr, 10);
    const scroller   = scrollableAncestor(firstRow);
    log(`List: ${totalCount} emails in folder, scroller found: ${!!scroller}`);

    const map = new Map();
    const grab = () => {
      document.querySelectorAll(SEL.emailRow).forEach((row) => {
        if (!row.id || map.has(row.id)) return;
        map.set(row.id, parseRow(row));
      });
    };

    grab();
    if (!scroller) return [...map.values()];

    let prevCount = -1, stable = 0, iters = 0;
    while (stable < CONFIG.stableRounds && iters < CONFIG.maxScrollIters) {
      if (window.__OUTLOOK_EXPORT_STOP) break;
      iters++;
      scroller.scrollTop += scroller.clientHeight * 0.75;
      await sleep(CONFIG.scrollDelayMs);
      grab();
      const count = map.size;
      if (count !== prevCount) { stable = 0; prevCount = count; } else { stable++; }
      if (totalCount > 0 && count >= totalCount) break;
      if (CONFIG.verbose && iters % 5 === 0) log(`  ...list scroll ${iters}: ${count}/${totalCount}`);
    }

    log(`List loaded: ${map.size} of ${totalCount} rows captured`);

    // Scroll back to top so we can click rows from the beginning
    scroller.scrollTop = 0;
    await sleep(600);

    return [...map.values()].sort((a, b) => a.posInSet - b.posInSet);
  }

  // ---- find a row in the current DOM (scroll if virtualized away) -----------
  async function findRow(meta, scroller) {
    let row = document.getElementById(meta.id);
    if (row) return row;

    // Row virtualized — scroll the list to its approximate position.
    // Use aria-setsize from any visible row as the total count.
    if (scroller) {
      const anyRow = document.querySelector(SEL.emailRow);
      const total  = anyRow ? parseInt(anyRow.getAttribute("aria-setsize") || "1", 10) : 1;
      const frac   = (meta.posInSet - 1) / Math.max(total - 1, 1);
      scroller.scrollTop = frac * scroller.scrollHeight;
      await sleep(CONFIG.rowFindScrollMs);
    }
    return document.getElementById(meta.id);
  }

  // ---- reading pane extraction ----------------------------------------------
  function sig() {
    const pane   = document.querySelector(SEL.readingPane);
    const from   = pane ? pane.querySelector(SEL.fromBtn) : null;
    const heads  = pane ? [...pane.querySelectorAll('[role="heading"]')] : [];
    const dateH  = heads.find((h) => /\d{1,2}\/\d{1,2}\/\d{4}/.test(txt(h)));
    return (from ? txt(from) : "") + "|" + (dateH ? txt(dateH) : "");
  }

  async function waitForPane(prevSig) {
    const t0 = Date.now();
    while (Date.now() - t0 < CONFIG.emailLoadTimeoutMs) {
      await sleep(200);
      const s = sig();
      if (s && s !== "|" && s !== prevSig) break;
    }
    await sleep(Math.max(CONFIG.minSettleMs, CONFIG.betweenEmailsMs));
    return sig();
  }

  function extractReadingPane() {
    const pane = document.querySelector(SEL.readingPane);
    if (!pane) return null;

    // --- From ---
    const fromBtn  = pane.querySelector(SEL.fromBtn);
    const fromRaw  = fromBtn ? txt(fromBtn) : "";
    // Format: "Display Name<email@domain>" or "Display Name"
    const fromMatch = fromRaw.match(/^(.+?)<([^>]+)>\s*$/);
    const from = fromMatch
      ? { name: norm(fromMatch[1]), email: norm(fromMatch[2]) }
      : { name: norm(fromRaw), email: "" };

    // --- To / CC ---
    const parseRecipients = (els) =>
      els.flatMap((el) => {
        const raw = txt(el).replace(/^(To|CC|Cc):\s*/u, "");
        // Recipients may be separated by newlines or semicolons.
        // Filter out lines that contain no real word characters (zero-width spaces, icons).
        return raw.split(/[\n;]+/).map(norm).filter((s) => /\w/.test(s));
      });
    const to = parseRecipients([...pane.querySelectorAll(SEL.toHeading)]);
    const cc = parseRecipients([...pane.querySelectorAll(SEL.ccHeading)]);

    // --- Date ---
    const heads   = [...pane.querySelectorAll('[role="heading"]')];
    const dateH   = heads.find((h) => /\d{1,2}\/\d{1,2}\/\d{4}/.test(txt(h)));
    const dateRaw = dateH ? txt(dateH) : "";
    let iso = null;
    try { const d = new Date(dateRaw); if (!isNaN(d.getTime())) iso = d.toISOString(); } catch (_) {}

    // --- Body ---
    const bodyEl = pane.querySelector(SEL.messageBody);
    const body   = cleanBody(bodyEl ? bodyEl.innerText : "");

    // --- Attachments ---
    const attachCont = pane.querySelector(SEL.attachContainer);
    const attachments = attachCont
      ? [...attachCont.querySelectorAll(SEL.attachItem)].map((el) => {
          // aria-label: "filename Open 123 KB"
          const label = el.getAttribute("aria-label") || "";
          const m     = label.match(/^(.+?)\s+Open\s+(.+)$/);
          return {
            filename: m ? m[1].trim() : label.trim(),
            size:     m ? m[2].trim() : "",
            // content not available via DOM (no download URL exposed)
          };
        })
      : [];

    return { from, to, cc, dateRaw, iso, body, attachments };
  }

  // ---- download -------------------------------------------------------------
  function downloadJSON(obj) {
    const blob  = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json" });
    const url   = URL.createObjectURL(blob);
    const a     = document.createElement("a");
    const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const folder = (obj.folder || "export").replace(/[^a-z0-9_-]/gi, "_");
    a.href     = url;
    a.download = `outlook-export-${folder}-${stamp}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  // ---- main -----------------------------------------------------------------
  async function run() {
    if (document.hidden) {
      log("WARNING: tab is in the background — keep Outlook visible/focused or " +
          "rendering will stall. Pausing 3 s for you to switch back.");
      await sleep(3000);
    }

    const folder = selectedFolderName();
    log(`Folder: "${folder}"`);

    const allMeta = await loadAllEmailRows();
    const limit = CONFIG.maxEmails > 0
      ? Math.min(CONFIG.maxEmails, allMeta.length)
      : allMeta.length;
    log(`Exporting ${limit} of ${allMeta.length} emails`);

    const result = {
      tool:        "outlook-cloud-export",
      version:     "1.0.0",
      folder,
      exportedAt:  new Date().toISOString(),
      emailCount:  0,
      emails:      [],
    };

    const firstRow = document.querySelector(SEL.emailRow);
    const scroller  = firstRow ? scrollableAncestor(firstRow) : null;
    let prevSig = "";

    for (let i = 0; i < limit; i++) {
      if (window.__OUTLOOK_EXPORT_STOP) { log("stop requested"); break; }
      const meta = allMeta[i];

      try {
        const row = await findRow(meta, scroller);
        if (!row) {
          log(`[${i+1}/${limit}] Row not found for "${meta.subject.slice(0, 50)}" — skipping`);
          result.emails.push({ index: i, ...meta, error: "row not visible in DOM" });
          continue;
        }

        row.scrollIntoView({ block: "center" });
        await sleep(100);
        row.click();
        prevSig = await waitForPane(prevSig);

        const pane = extractReadingPane();
        const email = {
          index:          i,
          id:             meta.id,
          convid:         meta.convid,
          subject:        meta.subject,
          from:           pane ? pane.from  : { name: meta.senderPreview, email: "" },
          to:             pane ? pane.to    : [],
          cc:             pane ? pane.cc    : [],
          dateRaw:        pane ? pane.dateRaw : meta.dateShort,
          iso:            pane ? pane.iso   : null,
          body:           pane ? pane.body  : "",
          hasAttachments: meta.hasAttachments,
          attachments:    pane ? pane.attachments : [],
        };

        result.emails.push(email);
        log(`[${i+1}/${limit}] "${meta.subject.slice(0, 50)}" ` +
            `← ${email.from.name} | ${email.dateRaw} | body: ${email.body.length} chars` +
            (email.attachments.length ? ` | ${email.attachments.length} attachment(s)` : ""));
      } catch (e) {
        log(`ERROR [${i+1}/${limit}] "${meta.subject}":`, e);
        result.emails.push({ index: i, ...meta, error: String(e) });
      }
    }

    result.emailCount = result.emails.length;
    window.__OUTLOOK_EXPORT_RESULT = result;
    log(`DONE: ${result.emailCount} emails exported from "${folder}".`);
    log("Result stored in window.__OUTLOOK_EXPORT_RESULT");
    if (CONFIG.download) downloadJSON(result);
    return result;
  }

  return run();
})();
