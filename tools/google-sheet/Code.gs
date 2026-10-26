/**
 * UTF site portal — Google Sheet backend.
 * Apps Script project (script.google.com) linked to the UTF Google Sheet by SHEET_ID.
 * run setup() once, then Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). Send the Web app URL back.
 *
 * Tabs it keeps:
 *   plots    – one row per edited plot (the site applies these on top of the drawing/tracker data)
 *   edits    – change log (when, plot, what changed, who)
 *   daily    – counts per day (handed over, under construction, sold, unsold, needs status, changes)
 *   current  – full list of all plots as the site last showed them
 *   Snapshot YYYY-MM-DD – copy of "current" taken every evening at 9 pm IST
 */
const PIN = '2025';                      // editor PIN — change here and re-deploy to change it
const TZ = 'Asia/Kolkata';
const SHEET_ID = '1ETbKpnj1TzfkVXNbkBnM49PZgkCFx9vnPkaEcvE86mg';   // the "UTF site data" Google Sheet
const SHEETS = {
  plots:   ['id','status','stage','customer','contractor','area','sales','villa','remarks','by','at'],
  edits:   ['at','plot','what','by'],
  daily:   ['date','done','wip','sold','nyt','needs','changes','at'],
  current: ['plot','villa','series','status','stage','customer','contractor','area','sales','remarks','edited']
};

/** Run once from the Apps Script editor (select "setup", press Run, allow access). */
function setup() {
  Object.keys(SHEETS).forEach(sh);
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailySnapshot').timeBased().atHour(21).everyDays(1).inTimezone(TZ).create();
}

function sh(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.getRange('A:Z').setNumberFormat('@');   // keep plot numbers like 1E, 40E as text
  }
  return s;
}
const today = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
function val(k, v) {
  if (v instanceof Date) return k === 'date' ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : v.toISOString();
  return v;
}
function rows(name) {
  const v = sh(name).getDataRange().getValues();
  const h = v.shift() || [];
  return v.filter(r => r.join('') !== '').map(r => Object.fromEntries(h.map((k, i) => [k, val(k, r[i])])));
}
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet() {
  const edits = rows('edits').sort((a, b) => String(b.at).localeCompare(String(a.at))).slice(0, 300);
  return out({ ok: true, today: today(), plots: rows('plots'), edits: edits, daily: rows('daily'),
               hasCurrent: sh('current').getLastRow() > 1 });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad_request' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (b.action === 'check') return out({ ok: String(b.pin) === PIN });
    if (b.action === 'sync') {                       // first visit of the day: record counts if missing
      if (sh('current').getLastRow() < 2 && b.current) writeCurrent(b.current);
      upsertDaily(b.counts, false);
      return out({ ok: true });
    }
    if (String(b.pin) !== PIN) return out({ ok: false, error: 'pin' });
    if (b.action === 'save') upsertPlot(b.plot);
    else if (b.action === 'reset') deletePlot(b.id);
    else return out({ ok: false, error: 'action' });
    sh('edits').appendRow([new Date().toISOString(), String(b.action === 'save' ? b.plot.id : b.id),
                           String(b.what || ''), String(b.editor || '')]);
    if (b.current) writeCurrent(b.current);
    upsertDaily(b.counts, true);
    return out({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function upsertPlot(p) {
  const s = sh('plots'), h = SHEETS.plots, v = s.getDataRange().getValues();
  const row = h.map(k => k === 'at' ? new Date().toISOString() : (p[k] == null ? '' : String(p[k])));
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(p.id)) { s.getRange(i + 1, 1, 1, h.length).setValues([row]); return; }
  }
  s.appendRow(row);
}
function deletePlot(id) {
  const s = sh('plots'), v = s.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === String(id)) s.deleteRow(i + 1);
}
function writeCurrent(cur) {
  const s = sh('current'), h = SHEETS.current;
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, h.length).clearContent();
  if (cur.length) s.getRange(2, 1, cur.length, h.length).setValues(cur.map(r => h.map(k => r[k] == null ? '' : String(r[k]))));
}
function changesToday() {
  const t = today();
  return rows('edits').filter(r => r.at && Utilities.formatDate(new Date(r.at), TZ, 'yyyy-MM-dd') === t).length;
}
function upsertDaily(c, force) {
  if (!c) return;
  const s = sh('daily'), t = today(), v = s.getDataRange().getValues();
  const row = [t, c.done, c.wip, c.sold, c.nyt, c.needs, changesToday(), new Date().toISOString()].map(String);
  for (let i = 1; i < v.length; i++) {
    if (String(val('date', v[i][0])) === t) { if (force) s.getRange(i + 1, 1, 1, row.length).setValues([row]); return; }
  }
  s.appendRow(row);
}

/** Runs every evening (set up by setup()): keeps a dated copy of all plots and the day's counts. */
function dailySnapshot() {
  const ss = SpreadsheetApp.openById(SHEET_ID), t = today(), cur = sh('current');
  if (cur.getLastRow() < 2) return;
  const name = 'Snapshot ' + t, old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  cur.copyTo(ss).setName(name);
  const d = rows('daily');
  if (d.length && !d.some(r => String(r.date) === t)) {
    const L = d[d.length - 1];
    sh('daily').appendRow([t, L.done, L.wip, L.sold, L.nyt, L.needs, changesToday(), new Date().toISOString()].map(String));
  } else if (d.length) {
    upsertDaily(d.find(r => String(r.date) === t), true);   // refresh the changes count
  }
}

