"""
Writing to the question bank: the same two operations for the Google Sheet (through its web app) and for
the local trivia_bank.xlsx.

  upsert(tab, key, values)   update the row whose key columns match; with no key on Questions/Categories,
                             add a row with the next ID. Returns the ID (or None).
  replace(tab, match, rows)  the rows matching `match` become exactly `rows`.

Only the tabs/columns in EDITABLE can be written (kept in step with EDITABLE in tbtgen_sheet.gs), and "(auto)"
formula columns are never touched.
"""
import datetime as dt
from pathlib import Path

from openpyxl import load_workbook

import build

EDITABLE = {
    "Questions": ["question_id", "question", "answer", "question_image", "answer_image", "media", "clip_start",
                  "clip_end", "seconds", "retired", "notes", "generations", "difficulty", "question_note",
                  "answer_note"],
    "Categories": ["category_id", "title", "blurb", "type", "season_tag", "retired", "notes"],
    "CategoryQuestions": ["category_id", "position", "question_id"],
    "Sets": ["set_id", "round", "slot", "category_id", "bonus", "tiebreaker_id"],
    "Tiebreakers": ["tiebreaker_id", "order", "question", "answer"],
    "Schedule": ["date", "venue_id", "set_id", "host", "calendar_event"],
    "Venues": ["venue_id", "name", "logo", "region", "night", "upcoming_nights", "active", "calendar_name", "audience",
               "team_max", "size_penalty", "penalty_per"],
    "Teams": ["team_id", "name", "venue_id", "aliases", "active", "notes"],
    "Results": ["date", "venue_id", "team_id", "players", "round1", "round2", "round3", "tiebreak", "photo", "notes"],
    "Seasons": ["venue_id", "season", "start_date", "weeks", "drop_lowest"],
    "Settings": ["key", "value", "meaning"],
}
ID_RULES = {"Questions": ("Q", 5), "Categories": ("C", 4), "Teams": ("TM", 4)}


class WriteError(Exception):
    pass


def check_cols(tab, obj):
    if tab not in EDITABLE:
        raise WriteError(f"tab not editable: {tab}")
    bad = [c for c in obj if c not in EDITABLE[tab]]
    if bad:
        raise WriteError(f"{tab}: column(s) not editable: {', '.join(bad)}")


def clean(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "Y" if v else ""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


class SheetWriter:
    def __init__(self, cfg):
        self.cfg = cfg

    def _call(self, action, **payload):
        try:
            return build.sheet_call(self.cfg, action, payload)
        except build.SheetError as e:
            msg = str(e)
            if "unknown action" in msg or "not editable" in msg:
                msg += " (the sheet's script is older than this TBTgen: paste in the new tbtgen_sheet.gs and " \
                       "use Deploy > Manage deployments > edit > Version: New version)"
            raise WriteError(msg)

    def upsert(self, tab, key, values):
        check_cols(tab, values)
        check_cols(tab, key or {})
        r = self._call("upsert", tab=tab, key=key or None, values={k: clean(v) for k, v in values.items()})
        return r.get("id")

    def append(self, tab, rows, batch=200):
        """Adds many rows; returns the new IDs (Questions/Categories) in order."""
        ids = []
        for i in range(0, len(rows), batch):
            chunk = rows[i:i + batch]
            for r in chunk:
                check_cols(tab, r)
            r = self._call("append", tab=tab, rows=[{k: clean(v) for k, v in x.items()} for x in chunk])
            ids += r.get("ids") or [None] * len(chunk)
        return ids

    def replace(self, tab, match, rows):
        check_cols(tab, match)
        for r in rows:
            check_cols(tab, r)
        self._call("replace", tab=tab, match=match, rows=[{k: clean(v) for k, v in r.items()} for r in rows])


class XlsxWriter:
    def __init__(self, path):
        self.path = Path(path)

    def _open(self, tab):
        try:
            wb = load_workbook(self.path)  # not data_only: keeps formulas
        except PermissionError:
            raise WriteError("can't open the spreadsheet: close it in Excel/Numbers and try again")
        if tab not in wb.sheetnames:
            raise WriteError(f"no tab named {tab}")
        ws = wb[tab]
        hdr = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
        return wb, ws, hdr

    @staticmethod
    def _text(v):
        if isinstance(v, (dt.date, dt.datetime)):
            return str(build.as_date(v))
        return "" if v is None else str(v).strip()

    def _last_row(self, ws):
        last = 1
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 1).value not in (None, ""):
                last = r
        return last

    def _find(self, ws, hdr, match):
        idx = {k: hdr.index(k) + 1 for k in match}
        return [r for r in range(2, self._last_row(ws) + 1)
                if all(self._text(ws.cell(r, idx[k]).value) == str(v).strip() for k, v in match.items())]

    def _write(self, ws, hdr, row, values):
        for k, v in values.items():
            if k not in hdr:
                raise WriteError(f"no column {k}")
            c = ws.cell(row, hdr.index(k) + 1)
            if k == "date" and isinstance(v, str) and len(v) == 10:
                c.value = dt.date.fromisoformat(v)
                c.number_format = "yyyy-mm-dd"
            else:
                c.value = clean(v) if clean(v) != "" else None
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.data_type = "s"  # text that starts with "=" stays text, never a formula

    def _save(self, wb):
        try:
            wb.save(self.path)
        except PermissionError:
            raise WriteError("can't save the spreadsheet: close it in Excel/Numbers and try again")

    def upsert(self, tab, key, values):
        check_cols(tab, values)
        check_cols(tab, key or {})
        wb, ws, hdr = self._open(tab)
        if key:
            rows = self._find(ws, hdr, key)
            if len(rows) > 1:
                raise WriteError(f"{tab}: more than one row matches {key}")
            if rows:
                self._write(ws, hdr, rows[0], values)
                self._save(wb)
                return key.get(EDITABLE[tab][0]) if tab in ID_RULES else None
            if tab in ID_RULES:
                raise WriteError(f"{tab}: {key} not found")
        full = {**(key or {}), **values}
        new_id = None
        if tab in ID_RULES and not full.get(EDITABLE[tab][0]):
            prefix, width = ID_RULES[tab]
            used = [int(str(ws.cell(r, 1).value)[len(prefix):]) for r in range(2, ws.max_row + 1)
                    if str(ws.cell(r, 1).value or "").startswith(prefix) and str(ws.cell(r, 1).value)[len(prefix):].isdigit()]
            new_id = f"{prefix}{max(used, default=0) + 1:0{width}d}"
            full[EDITABLE[tab][0]] = new_id
        self._write(ws, hdr, self._last_row(ws) + 1, full)
        self._save(wb)
        return new_id or full.get(EDITABLE[tab][0])

    def append(self, tab, rows):
        for r in rows:
            check_cols(tab, r)
        if not rows:
            return []
        wb, ws, hdr = self._open(tab)
        ids = []
        nxt = None
        if tab in ID_RULES:
            prefix, width = ID_RULES[tab]
            used = [int(str(ws.cell(r, 1).value)[len(prefix):]) for r in range(2, ws.max_row + 1)
                    if str(ws.cell(r, 1).value or "").startswith(prefix) and str(ws.cell(r, 1).value)[len(prefix):].isdigit()]
            nxt = max(used, default=0) + 1
        row = self._last_row(ws) + 1
        for r in rows:
            full = dict(r)
            if tab in ID_RULES and not full.get(EDITABLE[tab][0]):
                full[EDITABLE[tab][0]] = f"{ID_RULES[tab][0]}{nxt:0{ID_RULES[tab][1]}d}"
                nxt += 1
            ids.append(full.get(EDITABLE[tab][0]) if tab in ID_RULES else None)
            self._write(ws, hdr, row, full)
            row += 1
        self._save(wb)
        return ids

    def replace(self, tab, match, rows):
        check_cols(tab, match)
        for r in rows:
            check_cols(tab, r)
        if not match:
            raise WriteError("replace needs a match")
        wb, ws, hdr = self._open(tab)
        existing = self._find(ws, hdr, match)
        editable = [c for c in EDITABLE[tab] if c in hdr]
        for n, r in enumerate(rows):
            row = existing[n] if n < len(existing) else self._last_row(ws) + 1
            if n < len(existing):
                for c in editable:
                    ws.cell(row, hdr.index(c) + 1).value = None
            self._write(ws, hdr, row, {**match, **r})
        for row in existing[len(rows):]:
            for c in editable:
                ws.cell(row, hdr.index(c) + 1).value = None
        self._save(wb)


def writer(root):
    cfg = build.load_config()
    return SheetWriter(cfg) if cfg else XlsxWriter(Path(root) / "trivia_bank.xlsx")
