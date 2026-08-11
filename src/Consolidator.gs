/**
 * sheets-agent-inbox — Consolidator
 *
 * Pulls report spreadsheets that an agent (or any automation) dropped into a
 * Drive folder and files each one as a dated tab in a single master sheet,
 * then trashes the source so the folder stays clean.
 *
 * Runs entirely inside your Google account on a timer. Nothing has to reach
 * in from outside — which is the whole point: most cloud agents can write to
 * Drive through a connector but cannot make outbound HTTP calls.
 *
 * Setup: fill in CONFIG, run consolidateNow() once (authorize), then run
 * installDailyTrigger() once.
 *
 * MIT licensed. https://github.com/PopsPineDev/sheets-agent-inbox
 */

const CONFIG = {
  // Spreadsheet that collects everything. Copy the long id out of its URL:
  // docs.google.com/spreadsheets/d/THIS_PART/edit
  MASTER_SHEET_ID: 'PUT_YOUR_MASTER_SHEET_ID_HERE',

  // Folder your agent writes report spreadsheets into. Copy from the folder URL:
  // drive.google.com/drive/folders/THIS_PART
  INBOX_FOLDER_ID: 'PUT_YOUR_FOLDER_ID_HERE',

  // Only files whose name STARTS WITH this are imported. Keep your master
  // sheet's name outside this prefix so it can safely live in the same folder.
  // e.g. prefix 'Daily Report —' imports 'Daily Report — Aug 11, 2026'
  NAME_PREFIX: 'Daily Report —',

  // Trash each source file after its contents are safely written.
  // Trashed files sit in Drive Trash (recoverable ~30 days), not gone forever.
  TRASH_AFTER_IMPORT: true,

  // Hour of day (0-23, script timezone) for the daily run.
  RUN_AT_HOUR: 4,

  // Import every sheet in a multi-tab source file, not just the first.
  IMPORT_ALL_TABS: false,
};

const PROCESSED_KEY = 'SAI_CONSOLIDATED_FILE_IDS';

/** Run once: schedule consolidateNow() daily. Safe to re-run (replaces itself). */
function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'consolidateNow') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('consolidateNow')
    .timeBased().atHour(CONFIG.RUN_AT_HOUR).everyDays(1).create();
  Logger.log('Daily trigger installed for ~' + CONFIG.RUN_AT_HOUR + ':00 script time.');
}

/** Main entry point. Returns the number of files imported. */
function consolidateNow() {
  assertConfigured_();

  const props = PropertiesService.getScriptProperties();
  const done = JSON.parse(props.getProperty(PROCESSED_KEY) || '[]');
  const master = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
  const folder = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);

  const pending = [];
  const it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (it.hasNext()) {
    const f = it.next();
    if (f.getId() === CONFIG.MASTER_SHEET_ID) continue;            // never import the master
    if (f.getName().indexOf(CONFIG.NAME_PREFIX) !== 0) continue;   // not one of ours
    if (done.indexOf(f.getId()) !== -1) continue;                  // already imported
    pending.push(f);
  }
  // Oldest first, so the newest report ends up as the leftmost tab.
  pending.sort(function (a, b) { return a.getDateCreated() - b.getDateCreated(); });

  let imported = 0, trashed = 0;
  pending.forEach(function (f) {
    try {
      const label = f.getName().replace(CONFIG.NAME_PREFIX, '').trim() || f.getId().slice(0, 8);
      const sources = CONFIG.IMPORT_ALL_TABS
        ? SpreadsheetApp.openById(f.getId()).getSheets()
        : [SpreadsheetApp.openById(f.getId()).getSheets()[0]];

      let wroteSomething = false;
      sources.forEach(function (src, i) {
        const values = src.getDataRange().getValues();
        if (!values.length || !values[0].length) return;

        const tabName = uniqueTabName_(master, sources.length > 1 ? label + ' · ' + src.getName() : label);
        const sh = master.insertSheet(tabName, 0);
        sh.getRange(1, 1, values.length, values[0].length).setValues(values);
        sh.getRange(1, 1, 1, values[0].length).setFontWeight('bold');
        sh.setFrozenRows(1);
        wroteSomething = true;
      });

      // Force pending writes to commit BEFORE we touch the source file.
      SpreadsheetApp.flush();

      done.push(f.getId());
      if (wroteSomething) imported++;

      if (CONFIG.TRASH_AFTER_IMPORT && f.getId() !== CONFIG.MASTER_SHEET_ID) {
        f.setTrashed(true);
        trashed++;
      }
    } catch (err) {
      // Left in place on purpose: an un-trashed file is retried on the next run.
      Logger.log('Skipped "' + f.getName() + '": ' + err);
    }
  });

  props.setProperty(PROCESSED_KEY, JSON.stringify(done.slice(-500)));
  Logger.log('Imported ' + imported + ' file(s); trashed ' + trashed + '.');
  return imported;
}

/** Forget which files were already imported (re-imports anything still in the folder). */
function resetConsolidationState() {
  PropertiesService.getScriptProperties().deleteProperty(PROCESSED_KEY);
  Logger.log('State cleared.');
}

/* ---------- helpers ---------- */

function uniqueTabName_(ss, name) {
  const base = String(name).slice(0, 90);   // Sheets caps tab names at 100 chars
  let candidate = base, n = 2;
  while (ss.getSheetByName(candidate)) candidate = base + ' (' + (n++) + ')';
  return candidate;
}

function assertConfigured_() {
  ['MASTER_SHEET_ID', 'INBOX_FOLDER_ID'].forEach(function (k) {
    if (!CONFIG[k] || CONFIG[k].indexOf('PUT_YOUR_') === 0) {
      throw new Error('CONFIG.' + k + ' is not set — see the README.');
    }
  });
}
