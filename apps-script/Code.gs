/**
 * Tyrian Energy — Projects API (Google Apps Script + Google Sheet). Free.
 *
 * SETUP (about 10 minutes)
 *  1. Create a Google Sheet, e.g. "Tyrian Energy — Projets".
 *  2. In the sheet: Extensions → Apps Script. Replace the code with this file. Save.
 *  3. Project Settings (gear icon) → Script properties → Add:
 *       TEAM_KEY = a long passphrase you share with the team (used in ajouter-projet.html).
 *  4. Select the function "setup" → Run → authorise. It creates the "Projets" tab with its columns.
 *  5. Deploy → New deployment → type "Web app".
 *       Execute as: Me · Who has access: Anyone → Deploy → copy the URL ending in /exec.
 *  6. Paste that URL into PROJECTS_API in index.html AND in ajouter-projet.html.
 *
 * DAILY USE
 *  - Add a project: open ajouter-projet.html, fill in the form, save → it appears on the site.
 *  - Import existing projects (first time, sheet still empty): File → Import → Upload
 *    apps-script/projets-import.csv → "Replace current sheet". Its first row is the column header.
 *  - Hide a project: set its "published" cell to "non". Edit a project: change its cells.
 *  - After changing this code: Deploy → Manage deployments → Edit → Version: New version.
 */

var SHEET_NAME = 'Projets';
var HEADERS = ['id', 'created', 'published', 'type', 'discipline', 'sector', 'country', 'year',
  'title_fr', 'title_en', 'context_fr', 'context_en', 'action_fr', 'action_en',
  'r1_value', 'r1_label_fr', 'r1_label_en', 'r2_value', 'r2_label_fr', 'r2_label_en', 'tools'];

var LIMITS = {
  type: 10, discipline: 10, sector: 12, country: 8, year: 4,
  title_fr: 120, title_en: 120, context_fr: 260, context_en: 260, action_fr: 360, action_en: 360,
  r1_value: 24, r1_label_fr: 60, r1_label_en: 60, r2_value: 24, r2_label_fr: 60, r2_label_en: 60, tools: 160
};

var ALLOWED = {
  type: ['prior', 'tyrian'],
  discipline: ['pm', 'elec', 'meca', 'energy', 'data', 'wl'],
  sector: ['', 'manuf', 'auto', 'machines', 'tertiary', 'cement', 'utilities', 'public', 'telecom', 'firms', 'multi'],
  country: ['', 'tn', 'ca', 'fr', 'eu', 'gulf', 'africa', 'na']
};

var REQUIRED = ['type', 'discipline', 'title_fr', 'context_fr', 'action_fr'];

/** Run once from the editor: creates the sheet and its header row. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return 'OK';
}

/** GET → published projects as JSON: { ok: true, projects: [...] } */
function doGet() {
  var rows = readRows_().filter(function (p) {
    return String(p.published).trim().toLowerCase() !== 'non' && p.title_fr &&
      ALLOWED.discipline.indexOf(p.discipline) >= 0;
  });
  rows.forEach(function (p) { delete p.published; });
  return json_({ ok: true, projects: rows });
}

/** POST (text/plain JSON) → { key, confirm: true, project: {...} } */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var key = PropertiesService.getScriptProperties().getProperty('TEAM_KEY');
    if (!key || String(data.key || '') !== key) return json_({ ok: false, error: 'key' });
    if (data.confirm !== true) return json_({ ok: false, error: 'confirm' });

    var p = clean_(data.project || {});
    var missing = REQUIRED.filter(function (k) { return !p[k]; });
    if (missing.length) return json_({ ok: false, error: 'missing', fields: missing });

    p.id = Utilities.getUuid();
    p.created = new Date().toISOString();
    p.published = 'oui';

    lock.waitLock(10000);
    sheet_().appendRow(HEADERS.map(function (h) { return safe_(p[h]); }));
    return json_({ ok: true, id: p.id });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* not held */ }
  }
}

/* ---------------------------------------------------------------- helpers */

function sheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) { setup(); sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME); }
  return sh;
}

function readRows_() {
  var values = sheet_().getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0].map(function (h) { return String(h).trim(); });
  return values.slice(1).map(function (row) {
    var o = {};
    head.forEach(function (h, i) {
      if (HEADERS.indexOf(h) < 0) return;
      var v = row[i];
      o[h] = v instanceof Date ? v.toISOString() : (v === null || v === undefined ? '' : String(v));
    });
    return o;
  });
}

function clean_(src) {
  var p = {};
  Object.keys(LIMITS).forEach(function (k) {
    var v = String(src[k] === undefined || src[k] === null ? '' : src[k]).replace(/[\u0000-\u001f]+/g, ' ').trim();
    p[k] = v.slice(0, LIMITS[k]);
  });
  Object.keys(ALLOWED).forEach(function (k) { if (ALLOWED[k].indexOf(p[k]) < 0) p[k] = ''; });
  if (p.year && !/^\d{4}$/.test(p.year)) p.year = '';
  return p;
}

/** Stop spreadsheet formula injection (values starting with = + - @). */
function safe_(v) {
  var s = v === undefined || v === null ? '' : String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
