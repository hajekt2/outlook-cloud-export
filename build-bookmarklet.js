// Builds the bookmarklet version of outlook-export.js.
// Usage: node build-bookmarklet.js
const fs   = require("fs");
const path = require("path");

const src = fs.readFileSync(path.join(__dirname, "outlook-export.js"), "utf8");

function strip(code) {
  return code
    .replace(/\/\*[\s\S]*?\*\//g, "")        // block comments
    .replace(/^\s*\/\/.*$/gm, "")             // whole-line // comments
    .replace(/[ \t]+\/\/[^\n"'`]*$/gm, "")    // trailing // comments
    .replace(/\n{2,}/g, "\n")                 // collapse blank lines
    .trim();
}

const body = strip(src);

function bookmarklet(overrides) {
  const override = overrides
    ? `window.__OUTLOOK_EXPORT_OVERRIDE=${JSON.stringify(overrides)};\n`
    : "";
  return "javascript:" + encodeURIComponent(override + body);
}

const full    = bookmarklet(null);
const limited = bookmarklet({ maxEmails: 10 }); // test mode

const out =
`Outlook Cloud Export — bookmarklet
====================================
Create a new bookmark in Chrome, edit it, and paste the ENTIRE line below
(including the leading "javascript:") into the URL/location field.

Open https://outlook.cloud.microsoft/ , click on the folder you want to
export, and keep the tab FOCUSED/FOREGROUND the whole time. Then click the
bookmarklet. A JSON file downloads when done.

To abort early, paste in the DevTools console:
  window.__OUTLOOK_EXPORT_STOP = true

NOTE: Attachment file names and sizes are captured but not the file content
(Outlook web does not expose binary download URLs in the DOM). Download
attachments manually from the emails you need.

--- BOOKMARKLET: FULL EXPORT (${full.length} chars) ---
${full}

--- BOOKMARKLET: TEST (first 10 emails only, ${limited.length} chars) ---
${limited}
`;

fs.writeFileSync(path.join(__dirname, "outlook-export.bookmarklet.txt"), out);
console.log("Wrote outlook-export.bookmarklet.txt");
console.log(`  full export:  ${full.length} chars`);
console.log(`  test (10):    ${limited.length} chars`);
