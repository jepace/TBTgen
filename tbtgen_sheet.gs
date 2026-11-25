// ===== TBTgen sheet script · VERSION 2026-11-25 23:58 · paste ALL of this file, replacing everything =====
/**
 * TBTgen helpers for the Trivia Bank Google Sheet.
 *
 * Install (once): in the sheet, Extensions -> Apps Script, replace everything with this file, Save.
 * Then reload the sheet. A "TBTgen" menu appears. The first time you use "Pull schedule from calendar",
 * Google asks you to allow the script to read your calendars; that's expected.
 *
 * 1) IDs
 *
 * - Type a question (Questions tab, column B) or a category title (Categories tab, column B) in a row whose
 *   ID cell is empty, and the next free ID (Q00047, C0011, ...) is written in column A immediately.
 * - Works for pasted blocks of rows too.
 * - IDs are plain values, not formulas, so sorting/moving rows never changes them.
 * - A "TBTgen" menu has "Assign missing IDs" to fill any gaps (e.g. rows added before installing this).
 *
 * Don't delete question/category rows: mark them retired instead, so their IDs and history stay intact.
 *
 * 2) Schedule
 * - "Pull schedule from calendar" asks for a week (any date in it) and reads the calendar named by the
 *   `calendar_id` row on the Settings tab. Each event "Venue - Host" becomes a Schedule row:
 *   date, venue_id (matched on the Venues tab's calendar_name column), host, calendar_event. The set_id cell
 *   is left empty for Troy to choose (drop-down of the sets on the Sets tab).
 * - Running it again for the same week only adds events that aren't there yet.
 * - Events whose venue isn't on the Venues tab are still added (venue_id left empty) and listed at the end,
 *   so you can add the venue and fill in its venue_id.
 *
 * 3) Generator access (lets build.py read this sheet and log Usage, with no Google Cloud setup)
 * - TBTgen > Set up generator access… creates a secret key and shows the steps:
 *   Deploy > New deployment > type "Web app", Execute as: Me, Who has access: Anyone > Deploy,
 *   then put the Web app URL and the key into tbtgen_config.json next to build.py.
 * - The web app only answers requests that carry the key. Anyone without the key gets an error.
 *   To revoke access: "Make a new generator key" (old key stops working) or archive the deployment.
 * - After changing this script, use Deploy > Manage deployments > edit > Version: New version, so the
 *   web app runs the new code (the URL stays the same).
 */
const ID_RULES = {
  Questions:  { idCol: 1, textCol: 2, prefix: 'Q', width: 5 },
  Categories: { idCol: 1, textCol: 2, prefix: 'C', width: 4 },
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('TBTgen')
    .addItem('Pull schedule from calendar…', 'pullScheduleFromCalendar')
    .addItem('Assign missing IDs', 'assignAllMissingIds')
    .addSeparator()
    .addItem('Set up generator access…', 'setupGeneratorAccess')
    .addItem('Make a new generator key', 'newGeneratorKey')
    .addToUi();
}

function onEdit(e) {
  const sh = e.range.getSheet();
  const rule = ID_RULES[sh.getName()];
  if (!rule) return;
  // only react when the edit touches the text column
  if (e.range.getColumn() > rule.textCol || e.range.getLastColumn() < rule.textCol) return;
  fillIds_(sh, rule, Math.max(e.range.getRow(), 2), e.range.getLastRow());
}

function assignAllMissingIds() {
  const ss = SpreadsheetApp.getActive();
  let total = 0;
  Object.keys(ID_RULES).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() >= 2) total += fillIds_(sh, ID_RULES[name], 2, sh.getLastRow());
  });
  ss.toast(total + ' ID(s) assigned', 'TBTgen');
}

function fillIds_(sh, rule, first, last) {
  if (last < first) return 0;
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const n = last - first + 1;
    const idRange = sh.getRange(first, rule.idCol, n, 1);
    const ids = idRange.getValues();
    const texts = sh.getRange(first, rule.textCol, n, 1).getValues();
    let next = nextNumber_(sh, rule);
    let count = 0;
    for (let i = 0; i < n; i++) {
      if (ids[i][0] === '' && String(texts[i][0]).trim() !== '') {
        ids[i][0] = rule.prefix + String(next++).padStart(rule.width, '0');
        count++;
      }
    }
    if (count) idRange.setValues(ids);
    return count;
  } finally {
    lock.releaseLock();
  }
}

function nextNumber_(sh, rule) {
  const last = sh.getLastRow();
  if (last < 2) return 1;
  const re = new RegExp('^' + rule.prefix + '(\\d+)$');
  let max = 0;
  sh.getRange(2, rule.idCol, last - 1, 1).getValues().forEach(r => {
    const m = String(r[0]).match(re);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return max + 1;
}

// ---------------------------------------------------------------- schedule

function pullScheduleFromCalendar() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  const calId = settingValue_(ss, 'calendar_id');
  if (!calId) { ui.alert('Add a calendar_id row to the Settings tab first.'); return; }
  const cal = CalendarApp.getCalendarById(calId);
  if (!cal) { ui.alert('Could not open calendar ' + calId + '. Is it shared with you?'); return; }

  // which week?
  const nextMon = mondayOf_(new Date(Date.now() + 7 * 864e5));
  const resp = ui.prompt('Pull schedule from calendar',
    'Any date in the week to pull (yyyy-mm-dd). Leave blank for the week of ' +
    Utilities.formatDate(nextMon, tz, 'yyyy-MM-dd') + '.', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const typed = resp.getResponseText().trim();
  const start = typed ? mondayOf_(new Date(typed + 'T12:00:00')) : nextMon;
  if (isNaN(start)) { ui.alert('Could not read the date "' + typed + '".'); return; }
  const end = new Date(start.getTime() + 7 * 864e5);

  // venue lookup: calendar_name -> venue_id
  const venues = tableRows_(ss.getSheetByName('Venues'));
  const byCalName = {};
  venues.forEach(v => { if (v.calendar_name) byCalName[norm_(v.calendar_name)] = v.venue_id; });

  const sh = ss.getSheetByName('Schedule');
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = name => hdr.indexOf(name);
  ['date', 'venue_id', 'set_id', 'host', 'calendar_event'].forEach(n => {
    if (col(n) < 0) throw new Error('Schedule tab needs a "' + n + '" column');
  });
  const existing = {};
  tableRows_(sh).forEach(r => {
    if (r.date) existing[Utilities.formatDate(new Date(r.date), tz, 'yyyy-MM-dd') + '|' + r.calendar_event] = true;
  });

  const rows = [], unknown = [];
  cal.getEvents(start, end).forEach(ev => {
    if (ev.isAllDayEvent()) return;
    const title = ev.getTitle().trim();
    const parts = title.split(/\s+[-–—]\s+/);
    const venueName = parts[0], host = parts.length > 1 ? parts.slice(1).join(' - ') : '';
    const day = Utilities.formatDate(ev.getStartTime(), tz, 'yyyy-MM-dd');
    if (existing[day + '|' + title]) return;
    const venueId = byCalName[norm_(venueName)] || '';
    if (!venueId) unknown.push(day + '  ' + title);
    const row = new Array(hdr.length).fill('');
    row[col('date')] = new Date(day + 'T12:00:00');
    row[col('venue_id')] = venueId;
    row[col('host')] = host;
    row[col('calendar_event')] = title;
    rows.push(row);
  });

  if (rows.length) {
    const first = lastDataRow_(sh) + 1;
    sh.getRange(first, 1, rows.length, hdr.length).setValues(rows);
    sh.getRange(first, col('date') + 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  }
  let msg = rows.length + ' night(s) added for the week of ' + Utilities.formatDate(start, tz, 'MMM d') + '.';
  if (rows.length) msg += '\nPick a set for each one in the set_id column.';
  if (unknown.length) msg += '\n\nThese venues are not on the Venues tab (calendar_name column):\n' + unknown.join('\n');
  ui.alert(msg);
}

function settingValue_(ss, key) {
  const row = tableRows_(ss.getSheetByName('Settings')).find(r => r.key === key);
  return row ? String(row.value).trim() : '';
}

function tableRows_(sh) {
  const last = lastDataRow_(sh);
  if (last < 2) return [];
  const vals = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  const hdr = vals[0].map(String);
  return vals.slice(1).map(r => Object.fromEntries(hdr.map((h, i) => [h, r[i]])));
}

function lastDataRow_(sh) {
  // last row with something in column A (helper formulas further right may run further down)
  const colA = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = colA.length - 1; i >= 0; i--) if (String(colA[i][0]).trim() !== '') return i + 1;
  return 1;
}

function mondayOf_(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  x.setDate(x.getDate() - ((x.getDay() + 6) % 7));
  return x;
}

function norm_(s) {
  return String(s).toLowerCase().replace(/[’']/g, "'").replace(/\s+/g, ' ').trim();
}

// ---------------------------------------------------------------- generator access (web app)

const KEY_PROP = 'TBTGEN_KEY';

function setupGeneratorAccess() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty(KEY_PROP);
  if (!key) { key = makeKey_(); props.setProperty(KEY_PROP, key); }
  showKey_(key);
}

function newGeneratorKey() {
  const key = makeKey_();
  PropertiesService.getScriptProperties().setProperty(KEY_PROP, key);
  showKey_(key);
}

function makeKey_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function showKey_(key) {
  const html = HtmlService.createHtmlOutput(
    '<div style="font:14px sans-serif;line-height:1.5">' +
    '<p><b>1.</b> In the Apps Script editor: <b>Deploy → New deployment</b>, gear icon → <b>Web app</b>.<br>' +
    'Execute as: <b>Me</b>. Who has access: <b>Anyone</b>. Click <b>Deploy</b> and copy the <b>Web app URL</b>.</p>' +
    '<p><b>2.</b> Next to build.py, create <code>tbtgen_config.json</code>:</p>' +
    '<textarea style="width:100%;height:90px;font:12px monospace" onclick="this.select()">' +
    '{\n  "sheet_url": "PASTE THE WEB APP URL HERE",\n  "token": "' + key + '"\n}</textarea>' +
    '<p>Keep the key private: anyone with the URL <i>and</i> the key can read this sheet.</p></div>')
    .setWidth(560).setHeight(330);
  SpreadsheetApp.getUi().showModalDialog(html, 'TBTgen generator access');
}

// Bump when this script changes; TBTgen compares it with the version it expects (build.SCRIPT_VERSION).
const SCRIPT_VERSION = '2026-11-25 23:58';

function keyOk_(given) {
  const key = PropertiesService.getScriptProperties().getProperty(KEY_PROP);
  return !!key && typeof given === 'string' && given === key;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** GET is not used: the key must never be in a URL (URLs end up in logs and histories). Reads are POST 'tables'. */
function doGet(e) {
  // Opening the web app URL in a browser shows which script version is deployed (no data, no key needed).
  return json_({ ok: false, version: SCRIPT_VERSION,
                 error: 'unknown action: this TBTgen reads the sheet with POST (update TBTgen)' });
}

/** Every tab's values (dates as yyyy-mm-dd). */
function tables_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const tables = {};
  ss.getSheets().forEach(sh => {
    const rows = lastDataRow_(sh), cols = sh.getLastColumn();
    tables[sh.getName()] = rows < 1 || cols < 1 ? [] :
      sh.getRange(1, 1, rows, cols).getValues().map(r => r.map(v =>
        v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v));
  });
  return { ok: true, tables: tables, url: ss.getUrl(), version: SCRIPT_VERSION };
}

/**
 * POST {token, action, ...}. Actions:
 *  - tables: every tab's values (read only; no lock)
 *  - append_usage: rows [[date, venue_id, set_id, category_id, question_id], ...]
 *  - upsert:  {tab, key:{col:value,...}|null, values:{col:value,...}}
 *             Updates the row whose key columns match; with no key on Questions/Categories, adds a new row and
 *             gives it the next ID. Returns {id, row}.
 *  - append:  {tab, rows:[{col:value,...}, ...]}  adds rows in one go (new IDs on Questions/Categories); returns {ids}
 *  - replace: {tab, match:{col:value}, rows:[{col:value,...}, ...]}
 *             The rows matching `match` become exactly `rows` (rewritten in place, extras cleared, more appended).
 * Only the tabs/columns in EDITABLE can be written; "(auto)" formula columns are never touched.
 */
const EDITABLE = {
  Questions: ['question_id', 'question', 'answer', 'question_image', 'answer_image', 'media', 'clip_start', 'clip_end',
              'seconds', 'retired', 'notes', 'generations', 'difficulty', 'question_note', 'answer_note'],
  Categories: ['category_id', 'title', 'blurb', 'type', 'season_tag', 'retired', 'notes'],
  CategoryQuestions: ['category_id', 'position', 'question_id'],
  Sets: ['set_id', 'round', 'slot', 'category_id', 'bonus', 'tiebreaker_id'],
  Tiebreakers: ['tiebreaker_id', 'order', 'question', 'answer'],
  Schedule: ['date', 'venue_id', 'set_id', 'host', 'calendar_event'],
  Venues: ['venue_id', 'name', 'logo', 'region', 'night', 'upcoming_nights', 'active', 'calendar_name', 'audience'],
  Settings: ['key', 'value', 'meaning'],
};

function doPost(e) {
  const lock = LockService.getDocumentLock();
  try {
    const body = JSON.parse(e.postData.contents);
    if (!keyOk_(body.token)) return json_({ ok: false, error: 'wrong or missing key' });
    if (body.action === 'tables') return json_(tables_());
    lock.waitLock(20000);
    if (body.action === 'append_usage') return json_(appendUsage_(body.rows || []));
    if (body.action === 'upsert') return json_(upsert_(body.tab, body.key, body.values || {}));
    if (body.action === 'replace') return json_(replaceRows_(body.tab, body.match || {}, body.rows || []));
    if (body.action === 'append') return json_(appendRows_(body.tab, body.rows || []));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function appendUsage_(raw) {
  const rows = raw.map(r => [new Date(String(r[0]) + 'T12:00:00'), r[1], r[2], r[3], r[4]]);
  if (!rows.length) return { ok: true, added: 0 };
  const sh = SpreadsheetApp.getActive().getSheetByName('Usage');
  const first = lastDataRow_(sh) + 1;
  sh.getRange(first, 1, rows.length, 5).setValues(rows);
  sh.getRange(first, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  return { ok: true, added: rows.length };
}

function tabInfo_(tab) {
  if (!EDITABLE[tab]) throw new Error('tab not editable: ' + tab);
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(tab);
  if (!sh) throw new Error('no tab named ' + tab);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  return { sh: sh, hdr: hdr, tz: ss.getSpreadsheetTimeZone() };
}

function cellText_(v, tz) {
  return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).trim();
}

function toCell_(col, v) {
  if (v === null || v === undefined) return '';
  if (col === 'date' && /^\d{4}-\d{2}-\d{2}$/.test(String(v))) return new Date(v + 'T12:00:00');
  // Text that starts with "=" (e.g. from AI or an imported deck) must stay text, never become a formula.
  if (typeof v === 'string' && v.charAt(0) === '=') return "'" + v;
  return v;
}

function checkCols_(tab, obj) {
  Object.keys(obj).forEach(c => { if (EDITABLE[tab].indexOf(c) < 0) throw new Error(tab + ': column not editable: ' + c); });
}

function findRows_(info, match) {
  const keys = Object.keys(match);
  const last = lastDataRow_(info.sh);
  if (last < 2 || !keys.length) return [];
  const idx = keys.map(k => { const i = info.hdr.indexOf(k); if (i < 0) throw new Error('no column ' + k); return i; });
  const vals = info.sh.getRange(2, 1, last - 1, info.hdr.length).getValues();
  const out = [];
  vals.forEach((r, n) => {
    if (keys.every((k, j) => cellText_(r[idx[j]], info.tz) === String(match[k]).trim())) out.push(n + 2);
  });
  return out;
}

function writeRow_(info, rowNum, values) {
  Object.keys(values).forEach(c => {
    const i = info.hdr.indexOf(c);
    if (i < 0) throw new Error('no column ' + c);
    const cell = info.sh.getRange(rowNum, i + 1);
    cell.setValue(toCell_(c, values[c]));
    if (c === 'date') cell.setNumberFormat('yyyy-mm-dd');
  });
}

function upsert_(tab, key, values) {
  const info = tabInfo_(tab);
  checkCols_(tab, values);
  if (key) checkCols_(tab, key);
  const rule = ID_RULES[tab];
  if (key && Object.keys(key).length) {
    const rows = findRows_(info, key);
    if (rows.length > 1) throw new Error(tab + ': more than one row matches ' + JSON.stringify(key));
    if (rows.length === 1) {
      writeRow_(info, rows[0], values);
      return { ok: true, row: rows[0], id: rule ? key[EDITABLE[tab][0]] : null };
    }
    if (rule) throw new Error(tab + ': ' + JSON.stringify(key) + ' not found');
  }
  const rowNum = lastDataRow_(info.sh) + 1;
  const all = Object.assign({}, key || {}, values);
  let id = null;
  if (rule && !all[EDITABLE[tab][0]]) {
    id = rule.prefix + String(nextNumber_(info.sh, rule)).padStart(rule.width, '0');
    all[EDITABLE[tab][0]] = id;
  }
  writeRow_(info, rowNum, all);
  return { ok: true, row: rowNum, id: id || all[EDITABLE[tab][0]] || null };
}

function replaceRows_(tab, match, rows) {
  const info = tabInfo_(tab);
  checkCols_(tab, match);
  rows.forEach(r => checkCols_(tab, r));
  if (!Object.keys(match).length) throw new Error('replace needs a match');
  const existing = findRows_(info, match);
  const editableIdx = EDITABLE[tab].map(c => info.hdr.indexOf(c)).filter(i => i >= 0);
  rows.forEach((r, n) => {
    const full = Object.assign({}, match, r);
    const rowNum = n < existing.length ? existing[n] : lastDataRow_(info.sh) + 1;
    if (n < existing.length) editableIdx.forEach(i => info.sh.getRange(rowNum, i + 1).setValue(''));
    writeRow_(info, rowNum, full);
  });
  existing.slice(rows.length).forEach(rowNum =>
    editableIdx.forEach(i => info.sh.getRange(rowNum, i + 1).setValue('')));
  return { ok: true, written: rows.length, cleared: Math.max(0, existing.length - rows.length) };
}

function appendRows_(tab, rows) {
  const info = tabInfo_(tab);
  rows.forEach(r => checkCols_(tab, r));
  if (!rows.length) return { ok: true, ids: [] };
  const rule = ID_RULES[tab];
  const idCol = EDITABLE[tab][0];
  let next = rule ? nextNumber_(info.sh, rule) : 0;
  const ids = [];
  const width = info.hdr.length;
  const first = lastDataRow_(info.sh) + 1;
  // only touch editable columns, so "(auto)" formula columns keep their formulas
  const cols = EDITABLE[tab].map(c => info.hdr.indexOf(c)).filter(i => i >= 0);
  const block = rows.map(r => {
    const full = Object.assign({}, r);
    if (rule && !full[idCol]) full[idCol] = rule.prefix + String(next++).padStart(rule.width, '0');
    ids.push(rule ? full[idCol] : null);
    return full;
  });
  cols.forEach(ci => {
    const name = info.hdr[ci];
    const values = block.map(full => [toCell_(name, full[name])]);
    info.sh.getRange(first, ci + 1, values.length, 1).setValues(values);
    if (name === 'date') info.sh.getRange(first, ci + 1, values.length, 1).setNumberFormat('yyyy-mm-dd');
  });
  return { ok: true, ids: ids, first_row: first, width: width };
}
