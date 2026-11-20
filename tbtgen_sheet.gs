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
 */
const ID_RULES = {
  Questions:  { idCol: 1, textCol: 2, prefix: 'Q', width: 5 },
  Categories: { idCol: 1, textCol: 2, prefix: 'C', width: 4 },
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('TBTgen')
    .addItem('Pull schedule from calendar…', 'pullScheduleFromCalendar')
    .addItem('Assign missing IDs', 'assignAllMissingIds')
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
