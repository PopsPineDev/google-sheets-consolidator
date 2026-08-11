/**
 * sheets-agent-inbox — Webhook (OPTIONAL)
 *
 * Only useful for callers that CAN make outbound HTTP requests — your own VPS,
 * a CI job, a Zapier/Make step, a phone shortcut. Cloud AI agents often cannot
 * reach script.google.com; for those, use Consolidator.gs instead.
 *
 * Deploy: Deploy → New deployment → Web app → Execute as "Me",
 * Who has access "Anyone". Then set the SHARED_TOKEN script property.
 *
 * SECURITY: the token lives in Script Properties, never in this file.
 * Project Settings → Script properties → add SHARED_TOKEN = <a long random string>
 * Generate one with:  openssl rand -hex 24
 *
 * MIT licensed. https://github.com/PopsPineDev/sheets-agent-inbox
 */

const WEBHOOK_CONFIG = {
  TARGET_SHEET_ID: 'PUT_YOUR_MASTER_SHEET_ID_HERE',
  HEADER: ['Timestamp', 'Category', 'Title', 'Link', 'Notes'],
  TIMEZONE: 'Etc/UTC',            // e.g. 'America/Vancouver'
  TAB_DATE_FORMAT: 'MMM d, yyyy',
};

/**
 * POST JSON:
 * {
 *   "token": "<SHARED_TOKEN>",
 *   "tab":   "Aug 11, 2026",              // optional, defaults to today
 *   "rows":  [ { "Title": "...", ... } ], // keys matching HEADER; extras ignored
 *   "sections": { "NOTES": "free text\nsecond line" }   // optional
 * }
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    const token = PropertiesService.getScriptProperties().getProperty('SHARED_TOKEN');
    if (!token || body.token !== token) return json_({ ok: false, error: 'unauthorized' });

    const ss = SpreadsheetApp.openById(WEBHOOK_CONFIG.TARGET_SHEET_ID);
    const tabName = body.tab || Utilities.formatDate(
      new Date(), WEBHOOK_CONFIG.TIMEZONE, WEBHOOK_CONFIG.TAB_DATE_FORMAT);

    let sh = ss.getSheetByName(tabName);
    if (!sh) {
      sh = ss.insertSheet(tabName, 0);
      sh.appendRow(WEBHOOK_CONFIG.HEADER);
      sh.getRange(1, 1, 1, WEBHOOK_CONFIG.HEADER.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }

    const rows = (body.rows || []).map(function (r) {
      return WEBHOOK_CONFIG.HEADER.map(function (h) { return r[h] != null ? String(r[h]) : ''; });
    });
    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, WEBHOOK_CONFIG.HEADER.length).setValues(rows);
    }

    const sections = body.sections || {};
    Object.keys(sections).forEach(function (label) {
      if (!sections[label]) return;
      sh.appendRow(['']);
      sh.appendRow([label]);
      sh.getRange(sh.getLastRow(), 1).setFontWeight('bold');
      String(sections[label]).split('\n').forEach(function (line) {
        if (line.trim()) sh.appendRow(['', '', '', '', line.trim()]);
      });
    });

    return json_({ ok: true, tab: tabName, appended: rows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Health check: open the /exec URL in a browser. */
function doGet() {
  return json_({ ok: true, service: 'sheets-agent-inbox' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
