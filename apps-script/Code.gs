/**
 * DayBlocks backend: Google Sheet sync + weekly email summary.
 *
 * Setup (once):
 * 1. Create a new Google Sheet named "DayBlocks". Extensions > Apps Script. Paste this file.
 * 2. Run setup() once (approve permissions). It creates a random TOKEN in Script Properties,
 *    a weekly Monday-morning email trigger, and prints the token in the execution log.
 * 3. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 *    Copy the /exec URL.
 * 4. In the DayBlocks app: Settings > paste URL + token > Sync everything now.
 *
 * No secrets are hardcoded: the token lives in Project Settings > Script Properties.
 */

const LOG_HEADERS = ['Date', 'Weekday', 'Start', 'End', 'Minutes', 'Hours', 'Activity', 'Note'];

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!token || body.token !== token) return json_({ ok: false, error: 'Bad token' });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      upsertDays_(body.days || []);
      if (body.acts) saveActs_(body.acts);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true, days: (body.days || []).length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json_({ ok: true, app: 'dayblocks' });
}

/** Replace every row for the incoming dates with the new blocks (one row per block). */
function upsertDays_(days) {
  if (!days.length) return;
  const sh = sheet_('Log', LOG_HEADERS);
  const tz = Session.getScriptTimeZone();
  const incoming = new Set(days.map(d => d.date));
  const last = sh.getLastRow();
  const existing = last > 1 ? sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getDisplayValues() : [];
  const keep = existing.filter(r => !incoming.has(r[0]));
  const fresh = [];
  days.forEach(d => (d.blocks || []).forEach(b => {
    const wd = Utilities.formatDate(parseYmd_(b.date), tz, 'EEE');
    fresh.push([b.date, wd, b.start, b.end, b.minutes, +(b.minutes / 60).toFixed(2), b.activity, b.note || '']);
  }));
  const rows = keep.concat(fresh).sort((a, b) => (a[0] + a[2]).localeCompare(b[0] + b[2]));
  if (last > 1) sh.getRange(2, 1, last - 1, LOG_HEADERS.length).clearContent();
  if (rows.length) {
    const rng = sh.getRange(2, 1, rows.length, LOG_HEADERS.length);
    rng.setNumberFormat('@');                     // keep dates/times as plain text, no auto-conversion
    sh.getRange(2, 5, rows.length, 2).setNumberFormat('0.##');
    rng.setValues(rows.map(r => [r[0], r[1], r[2], r[3], Number(r[4]), Number(r[5]), r[6], r[7]]));
  }
}

function saveActs_(acts) {
  const sh = sheet_('Activities', ['Activity', 'Color']);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  if (acts.length) sh.getRange(2, 1, acts.length, 2).setValues(acts.map(a => [a.name, a.color]));
}

/** Weekly email: last Monday to Sunday, compared with the week before. */
function weeklySummary() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const dow = (now.getDay() + 6) % 7;                       // 0 = Monday
  const thisMon = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dow);
  const lastMon = addDays_(thisMon, -7), prevMon = addDays_(thisMon, -14);
  const fmt = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');

  const rows = readLog_();
  const cur = aggregate_(rows, fmt(lastMon), fmt(addDays_(lastMon, 6)));
  const prev = aggregate_(rows, fmt(prevMon), fmt(addDays_(prevMon, 6)));
  if (!cur.total) return; // nothing logged, no email

  const names = Object.keys(cur.byAct).sort((a, b) => cur.byAct[b] - cur.byAct[a]);
  const h = m => (m / 60).toFixed(1) + 'h';
  const tr = names.map(n => {
    const d = cur.byAct[n] - (prev.byAct[n] || 0);
    const delta = prev.total ? (d === 0 ? 'same' : (d > 0 ? '+' : '-') + h(Math.abs(d))) : '';
    const avg = h(cur.byAct[n] / Math.max(1, cur.days));
    return `<tr><td>${n}</td><td align="right"><b>${h(cur.byAct[n])}</b></td><td align="right">${avg}/day</td>` +
           `<td align="right">${Math.round(cur.byAct[n] / cur.total * 100)}%</td><td align="right">${delta}</td></tr>`;
  }).join('');
  const range = `${Utilities.formatDate(lastMon, tz, 'MMM d')} to ${Utilities.formatDate(addDays_(lastMon, 6), tz, 'MMM d')}`;
  const html =
    `<div style="font-family:system-ui,Arial,sans-serif;max-width:560px">` +
    `<h2 style="margin:0 0 4px">Your week: ${range}</h2>` +
    `<p style="color:#666;margin:0 0 12px">${h(cur.total)} logged across ${cur.days} day(s) (${Math.round(cur.total / (7 * 1440) * 100)}% of the week)</p>` +
    `<table cellpadding="6" style="border-collapse:collapse;width:100%;font-size:14px">` +
    `<tr style="background:#f1f3f5"><th align="left">Activity</th><th align="right">Total</th><th align="right">Avg</th><th align="right">Share</th><th align="right">vs prior week</th></tr>${tr}</table>` +
    `<p style="color:#888;font-size:12px">Open DayBlocks for day-by-day detail. Data: ${SpreadsheetApp.getActive().getUrl()}</p></div>`;

  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: `DayBlocks weekly summary: ${range}`, htmlBody: html });
}

/** Run once. Safe to re-run: keeps the existing token, never duplicates the trigger. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('TOKEN');
  if (!token) { token = Utilities.getUuid().replace(/-/g, ''); props.setProperty('TOKEN', token); }
  sheet_('Log', LOG_HEADERS);
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'weeklySummary').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('weeklySummary').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  Logger.log('Your DayBlocks sync token (paste into the app): ' + token);
}

/* ---------- helpers ---------- */
function readLog_() {
  const sh = sheet_('Log', LOG_HEADERS);
  return sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, LOG_HEADERS.length).getDisplayValues() : [];
}
function aggregate_(rows, from, to) {
  const byAct = {}, days = new Set(); let total = 0;
  rows.forEach(r => {
    if (r[0] < from || r[0] > to) return;
    const m = Number(r[4]) || 0;
    byAct[r[6]] = (byAct[r[6]] || 0) + m; total += m; days.add(r[0]);
  });
  return { byAct, total, days: days.size };
}
function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function parseYmd_(s) { const p = s.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
function addDays_(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
