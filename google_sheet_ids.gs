/**
 * TBTgen: automatic, permanent IDs for the Google Sheet version of the Trivia Bank.
 *
 * Install (once): in the sheet, Extensions -> Apps Script, replace everything with this file, Save.
 * Then reload the sheet. No permissions are needed for the automatic part.
 *
 * - Type a question (Questions tab, column B) or a category title (Categories tab, column B) in a row whose
 *   ID cell is empty, and the next free ID (Q00047, C0011, ...) is written in column A immediately.
 * - Works for pasted blocks of rows too.
 * - IDs are plain values, not formulas, so sorting/moving rows never changes them.
 * - A "TBTgen" menu has "Assign missing IDs" to fill any gaps (e.g. rows added before installing this).
 *
 * Don't delete question/category rows: mark them retired instead, so their IDs and history stay intact.
 */
const ID_RULES = {
  Questions:  { idCol: 1, textCol: 2, prefix: 'Q', width: 5 },
  Categories: { idCol: 1, textCol: 2, prefix: 'C', width: 4 },
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('TBTgen')
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
