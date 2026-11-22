#!/usr/bin/env python3
"""
TBTgen: Trivia By Troy deck generator.

Reads the Trivia Bank Google Sheet (or a trivia_bank.xlsx file) and builds one PowerPoint deck per
scheduled venue-night that has a set assigned.

The Google Sheet is used when tbtgen_config.json exists next to this file:
  {"sheet_url": "<the TBTgen web app URL>", "token": "<the key shown by TBTgen > Set up generator access>"}
Otherwise trivia_bank.xlsx is used; --bank FILE.xlsx forces a file.

  python3 build.py                         # every night on the Schedule tab that has a set
  python3 build.py --week next             # next Mon-Sun (or --week this, --week 2026-11-30)
  python3 build.py --week next --day wed   # just next week's Wednesday decks
  python3 build.py --date 2026-12-02       # one night
  python3 build.py --day wed               # every Wednesday on the schedule
  python3 build.py --from 2026-11-23 --to 2026-12-06
  python3 build.py --venue wine_press,pops
  python3 build.py --video 2026-W48-A      # at-home video of a set -> out/video/
  python3 build.py --record                # also log the questions used to the Usage tab
  python3 build.py --check                 # only run the repeat/consistency checks
  python3 build.py --ids                   # only fill in IDs for new questions/categories
"""
import argparse
import datetime as dt
import hashlib
import json
import subprocess
import sys
import urllib.parse
import urllib.request
from collections import defaultdict
from pathlib import Path

from lxml import etree
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageEnhance, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
CACHE = ROOT / "build"
CONFIG = ROOT / "tbtgen_config.json"   # {"sheet_url": "...", "token": "..."}; not in Git

# ----------------------------------------------------------------- look & feel
# Fonts taken from Troy's Keynote template (TBT Base.key). All three come with macOS; Windows has
# Trebuchet MS and Brush Script MT, but not Savoye LET (PowerPoint substitutes another font there).
FONT = "Trebuchet MS"         # questions, answers, headers, rules
SCRIPT = "Savoye LET"         # script titles: Trivia Night, Round One, category names, host
BRUSH = "Brush Script MT"     # bonus labels: "Tonight's Bonus Category", "Double Points!"
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GOLD = RGBColor(0xF5, 0xC2, 0x42)
W, H = Inches(13.333), Inches(7.5)
LOGO_BOX = (Inches(11.85), Inches(0.12), Inches(1.3), Inches(1.3))
NUM_WORDS = {1: "One", 2: "Two", 3: "Three"}
ROUND_SUB = {1: "Good luck!", 2: "Hang in there…", 3: "It all comes down to this…"}


# ======================================================================= data
YES = ("Y", "YES", "TRUE", "1", "X")
AUTO = "(auto)"  # prefix of helper columns filled by formulas; the generator ignores them


def is_yes(v):
    return str(v).strip().upper() in YES


def table_rows(values):
    """Rows of a tab (list of lists, header first) as dicts keyed by header. Columns with no header are
    ignored; "(auto)" helper columns are kept but don't make an otherwise empty row count."""
    if not values:
        return []
    hdr = [str(h).strip() if h not in (None, "") else None for h in values[0]]
    out = []
    for r in values[1:]:
        r = list(r) + [""] * (len(hdr) - len(r))          # short rows: missing cells are empty
        rec = {h: ("" if v is None else v) for h, v in zip(hdr, r) if h}
        if all(v in (None, "") for k, v in rec.items() if not k.startswith(AUTO)):
            continue
        out.append(rec)
    return out


# ------------------------------------------------------------ sources: .xlsx or the Google Sheet
def tables_from_xlsx(path):
    wb = load_workbook(path, data_only=True)
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


class SheetError(Exception):
    pass


def sheet_call(cfg, action, payload=None):
    """Talk to the TBTgen web app attached to the Google Sheet (see tbtgen_sheet.gs)."""
    if payload is None:
        url = cfg["sheet_url"] + "?" + urllib.parse.urlencode({"token": cfg["token"], "action": action})
        req = urllib.request.Request(url)
    else:
        body = json.dumps({"token": cfg["token"], "action": action, **payload}).encode()
        req = urllib.request.Request(cfg["sheet_url"], data=body, headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:   # follows Google's redirect
            raw = resp.read().decode("utf-8")
    except Exception as e:
        raise SheetError(f"could not reach the Google Sheet web app: {e}")
    try:
        data = json.loads(raw)
    except ValueError:
        hint = " (got a Google sign-in page: the web app's access must be 'Anyone')" if "<html" in raw[:200].lower() else ""
        raise SheetError("the Google Sheet web app did not return data" + hint)
    if not data.get("ok"):
        raise SheetError(data.get("error", "unknown error from the Google Sheet web app"))
    return data


SHEET_URL = None  # the Google Sheet's own address, when the web app reports it


def tables_from_sheet(cfg):
    global SHEET_URL
    data = sheet_call(cfg, "tables")
    SHEET_URL = data.get("url") or SHEET_URL
    CACHE.mkdir(exist_ok=True)
    (CACHE / "sheet_snapshot.json").write_text(json.dumps(data["tables"]))   # last copy, for reference
    return data["tables"]


def load_config():
    if CONFIG.exists():
        cfg = json.loads(CONFIG.read_text())
        if cfg.get("sheet_url") and cfg.get("token"):
            return cfg
    return None


def load_bank(tables):
    bank = {name: table_rows(vals) for name, vals in tables.items()}
    bank["settings"] = {r["key"]: r["value"] for r in bank["Settings"]}
    bank["venues"] = {r["venue_id"]: r for r in bank["Venues"]}
    bank["cats"] = {r["category_id"]: r for r in bank["Categories"] if r["category_id"]}
    bank["qs"] = {r["question_id"]: r for r in bank["Questions"] if r["question_id"]}
    # category -> its questions, in position order (a question may belong to many categories)
    links = defaultdict(list)
    for r in bank["CategoryQuestions"]:
        if r["category_id"] and r["question_id"]:
            links[r["category_id"]].append(r)
    for v in links.values():
        v.sort(key=lambda r: int(r["position"] or 0))
    bank["links"] = links
    sets = defaultdict(list)
    for r in bank["Sets"]:
        sets[r["set_id"]].append(r)
    bank["sets"] = sets
    tb = defaultdict(list)
    for r in bank["Tiebreakers"]:
        tb[r["set_id"]].append(r)
    bank["tiebreakers"] = tb
    return bank


def cat_questions(bank, cid):
    """The question rows of a category, in order (each with its 'position')."""
    return [dict(bank["qs"][l["question_id"]], position=l["position"])
            for l in bank["links"].get(cid, []) if l["question_id"] in bank["qs"]]


def set_question_ids(bank, set_id):
    return {l["question_id"] for sl in bank["sets"].get(set_id, []) for l in bank["links"].get(sl["category_id"], [])}


def assign_ids(path):
    """Give every category / question that has no ID yet the next free C#### / Q##### ID, and save.
    IDs are never reused or renumbered. Returns how many were assigned."""
    wb = load_workbook(path)  # not data_only: keeps the helper formulas intact
    n = 0
    for tab, col, prefix, width, needs in (("Categories", "category_id", "C", 4, "title"),
                                           ("Questions", "question_id", "Q", 5, "question")):
        ws = wb[tab]
        hdr = [c.value for c in ws[1]]
        ci, ni = hdr.index(col), hdr.index(needs)
        used = [int(str(r[ci].value)[1:]) for r in ws.iter_rows(min_row=2)
                if r[ci].value and str(r[ci].value).startswith(prefix) and str(r[ci].value)[1:].isdigit()]
        nxt = max(used, default=0) + 1
        for r in ws.iter_rows(min_row=2):
            if not r[ci].value and r[ni].value not in (None, ""):
                r[ci].value = f"{prefix}{nxt:0{width}d}"
                nxt += 1
                n += 1
    if n:
        wb.save(path)
    return n


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return dt.date.fromisoformat(str(v).strip()[:10])


def short(text, n=60):
    t = " ".join(str(text).split())
    return t if len(t) <= n else t[: n - 1] + "…"


# ===================================================================== checks
def check_sets(bank, set_ids):
    """Problems inside the sets themselves (no schedule needed). Returns (errors, warnings)."""
    errs, warns = [], []
    cat_title = lambda cid: bank["cats"].get(cid, {}).get("title", cid)
    for set_id in sorted(set_ids):
        slots = bank["sets"].get(set_id)
        if not slots:
            errs.append(f"Set {set_id}: not found on the Sets tab")
            continue
        if len(slots) != 9:
            errs.append(f"Set {set_id}: has {len(slots)} categories, expected 9 (3 rounds x 3)")
        bonus = [r for r in slots if is_yes(r["bonus"])]
        if len(bonus) != 1:
            warns.append(f"Set {set_id}: has {len(bonus)} bonus categories (expected exactly 1)")
        seen_q = {}
        for r in slots:
            cid = r["category_id"]
            if cid not in bank["cats"]:
                errs.append(f"Set {set_id}: unknown category '{cid}'")
                continue
            if is_yes(bank["cats"][cid].get("retired")):
                errs.append(f"Set {set_id}: category {cid} '{cat_title(cid)}' is retired")
            links = bank["links"].get(cid, [])
            positions = sorted(int(l["position"] or 0) for l in links)
            if positions != [1, 2, 3, 4, 5]:
                errs.append(f"Category {cid} '{cat_title(cid)}': positions are {positions}, expected 1-5 "
                            f"(CategoryQuestions tab)")
            for l in links:
                qid = l["question_id"]
                qq = bank["qs"].get(qid)
                if not qq:
                    errs.append(f"Category {cid}: question {qid} is not on the Questions tab")
                    continue
                if is_yes(qq.get("retired")):
                    errs.append(f"Category {cid}: question {qid} '{short(qq['question'])}' is retired")
                if qid in seen_q:
                    errs.append(f"Set {set_id}: question {qid} is in both {seen_q[qid]} and {cid}")
                seen_q[qid] = cid
                for col in ("question_image", "answer_image", "media"):
                    if qq[col] and not (ASSETS / qq[col]).exists():
                        errs.append(f"Question {qid}: missing file assets/{qq[col]}")
        if not bank["tiebreakers"].get(set_id):
            warns.append(f"Set {set_id}: no tiebreaker questions")
    return errs, warns


def check(bank, schedule):
    """Set problems plus schedule problems (repeats at a venue, team hopping). Returns (errors, warnings)."""
    errs, warns = check_sets(bank, {r["set_id"] for r in schedule})
    lookback = int(bank["settings"].get("repeat_lookback_days", 365))
    cat_title = lambda cid: bank["cats"].get(cid, {}).get("title", cid)

    # 1) has this venue already seen these categories / questions (in any category)?
    for r in schedule:
        d = as_date(r["date"])
        vname = bank["venues"][r["venue_id"]]["name"]
        set_cats = {sl["category_id"] for sl in bank["sets"].get(r["set_id"], [])}
        cat_of = {l["question_id"]: sl["category_id"] for sl in bank["sets"].get(r["set_id"], [])
                  for l in bank["links"].get(sl["category_id"], [])}
        warned_cats, warned_qs = set(), set()
        for u in bank["Usage"]:
            if u["venue_id"] != r["venue_id"]:
                continue
            ud = as_date(u["date"])
            if (ud == d and u["set_id"] == r["set_id"]) or not (0 <= (d - ud).days <= lookback):
                continue
            if u["category_id"] in set_cats and u["category_id"] not in warned_cats:
                warned_cats.add(u["category_id"])
                warns.append(f"{d} {vname}: category {u['category_id']} '{cat_title(u['category_id'])}' "
                             f"was already played there on {ud}")
            qid = u.get("question_id")
            if qid in cat_of and cat_of[qid] != u["category_id"] and qid not in warned_qs:
                warned_qs.add(qid)
                warns.append(f"{d} {vname}: question {qid} '{short(bank['qs'][qid]['question'])}' "
                             f"(now in '{cat_title(cat_of[qid])}') was already asked there on {ud} "
                             f"in '{cat_title(u['category_id'])}'")

    # 2) shared questions at venues in the same region on back-to-back nights (team hopping)
    for a in schedule:
        for b in schedule:
            if a is b or a["venue_id"] == b["venue_id"]:
                continue
            va, vb = bank["venues"][a["venue_id"]], bank["venues"][b["venue_id"]]
            da, db = as_date(a["date"]), as_date(b["date"])
            if va["region"] and va["region"] == vb["region"] and 0 < (db - da).days <= 1:
                shared = set_question_ids(bank, a["set_id"]) & set_question_ids(bank, b["set_id"])
                if shared:
                    warns.append(f"{va['name']} ({da}) and {vb['name']} ({db}) are both in region {va['region']} "
                                 f"on back-to-back nights and share {len(shared)} questions")
    return errs, warns


# ================================================================ media helpers
def ffprobe_dur(path):
    out = subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", str(path)],
                         capture_output=True, text=True, check=True).stdout
    return float(out.strip())


def cache_name(*parts, ext):
    h = hashlib.sha1(json.dumps([str(p) for p in parts]).encode()).hexdigest()[:12]
    CACHE.mkdir(exist_ok=True)
    return CACHE / f"{h}{ext}"


def dim_logo():
    out = CACHE / "tbt_logo_dim.png"
    if not out.exists():
        CACHE.mkdir(exist_ok=True)
        im = Image.open(ASSETS / "brand/tbt_logo.png").convert("RGBA")
        im = ImageEnhance.Color(im).enhance(0.0)
        im = ImageEnhance.Brightness(im).enhance(0.35)
        im.save(out)
    return out


BOLD_FONTS = [  # first one found is used; all of these include the ♪ ♫ characters
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",            # Linux
    "/usr/local/share/fonts/dejavu/DejaVuSans-Bold.ttf",               # FreeBSD (x11-fonts/dejavu)
    "/System/Library/Fonts/Supplemental/Arial Bold.ttf",               # macOS
    "/Library/Fonts/Arial Bold.ttf",                                   # older macOS
    "C:/Windows/Fonts/arialbd.ttf",                                    # Windows
]


def bold_font(size):
    for path in BOLD_FONTS:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default(size=size)   # Pillow's built-in font


def music_frame():
    out = CACHE / "music_frame.png"
    if not out.exists():
        CACHE.mkdir(exist_ok=True)
        im = Image.new("RGB", (1280, 720), "black")
        d = ImageDraw.Draw(im)
        f = bold_font(260)
        d.text((640, 330), "♪ ♫", font=f, fill="#d63ad6", anchor="mm")
        f2 = bold_font(60)
        d.text((640, 600), "Name that song!", font=f2, fill="white", anchor="mm")
        im.save(out)
    return out


def still_plus_audio(image, audio, start, end):
    """Make an mp4 of a still image + (trimmed) audio, so audio plays the same way as video."""
    out = cache_name("still", image, audio, start, end, ext=".mp4")
    if not out.exists():
        trim = []
        if start not in ("", None):
            trim += ["-ss", str(start)]
        if end not in ("", None):
            trim += ["-to", str(end)]
        subprocess.run(["ffmpeg", "-v", "error", "-y", "-loop", "1", "-i", str(image), *trim, "-i", str(audio),
                        "-vf", "scale=1280:720:force_original_aspect_ratio=decrease,pad=1280:720:(ow-iw)/2:(oh-ih)/2,format=yuv420p",
                        "-c:v", "libx264", "-tune", "stillimage", "-r", "2", "-c:a", "aac", "-b:a", "192k",
                        "-shortest", str(out)], check=True)
    return out


def trimmed_video(video, start, end):
    if start in ("", None) and end in ("", None):
        return Path(video)
    out = cache_name("trim", video, start, end, ext=".mp4")
    if not out.exists():
        trim = (["-ss", str(start)] if start not in ("", None) else []) + (["-to", str(end)] if end not in ("", None) else [])
        subprocess.run(["ffmpeg", "-v", "error", "-y", *trim, "-i", str(video), "-c:v", "libx264", "-pix_fmt", "yuv420p",
                        "-c:a", "aac", str(out)], check=True)
    return out


def poster_of(video):
    out = cache_name("poster", video, ext=".jpg")
    if not out.exists():
        subprocess.run(["ffmpeg", "-v", "error", "-y", "-ss", "0.5", "-i", str(video), "-frames:v", "1", str(out)], check=True)
    return out


# ============================================================ slide primitives
class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = W, H
        self.blank = self.prs.slide_layouts[6]
        # drop the template's unused sample layouts (one contains a placeholder image Keynote complains about)
        for layout in list(self.prs.slide_layouts):
            if layout is not self.blank:
                self.prs.slide_layouts.remove(layout)

    def slide(self):
        s = self.prs.slides.add_slide(self.blank)
        bg = s.background.fill
        bg.solid()
        bg.fore_color.rgb = RGBColor(0, 0, 0)
        return s


def text(slide, x, y, w, h, content, size=20, bold=False, italic=False, font=FONT, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, color=WHITE, spacing=None):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, line in enumerate(str(content).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if spacing:
            p.space_after = Pt(spacing)
        r = p.add_run()
        r.text = line
        f = r.font
        f.size, f.bold, f.italic, f.name = Pt(size), bold, italic, font
        f.color.rgb = color
    return tb


def bullets(slide, x, y, w, h, items, size=20):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        head, _, small = item.partition("|")
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(10)
        r = p.add_run()
        r.text = "•  " + head
        r.font.size, r.font.bold, r.font.name, r.font.color.rgb = Pt(size), True, FONT, WHITE
        if small:
            r = p.add_run()
            r.text = "  " + small
            r.font.size, r.font.italic, r.font.name, r.font.color.rgb = Pt(size * 0.6), True, FONT, WHITE
    return tb


def picture(slide, path, x, y, w, h):
    """Place an image fitted (aspect preserved) and centred inside the box."""
    with Image.open(path) as im:
        iw, ih = im.size
    scale = min(w / iw, h / ih)
    pw, ph = int(iw * scale), int(ih * scale)
    return slide.shapes.add_picture(str(path), Emu(x + (w - pw) // 2), Emu(y + (h - ph) // 2), Emu(pw), Emu(ph))


def fit_box(path, x, y, w, h):
    with Image.open(path) as im:
        iw, ih = im.size
    scale = min(w / iw, h / ih)
    pw, ph = int(iw * scale), int(ih * scale)
    return Emu(x + (w - pw) // 2), Emu(y + (h - ph) // 2), Emu(pw), Emu(ph)


def question_block(slide, num, question):
    """'Question N' + question text; returns the y (EMU) where the content area can start."""
    import math
    text(slide, Inches(0.35), Inches(0.2), Inches(11), Inches(0.7), f"Question {num}", size=28)
    lines = sum(max(1, math.ceil(len(par) / 95)) for par in str(question).split("\n"))
    h = Inches(0.36) * lines + Inches(0.1)
    text(slide, Inches(0.35), Inches(0.85), Inches(11.3), h, question, size=20, bold=True)
    return max(Inches(2.25), Inches(0.85) + h + Inches(0.15))


def header(slide, title, logo, size=28):
    text(slide, Inches(0.35), Inches(0.2), Inches(11), Inches(0.7), title, size=size)
    picture(slide, logo, *LOGO_BOX)


# ====================================================== animation / timing XML
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


class Timing:
    """Builds PowerPoint <p:timing>: effects that run when the slide opens, and click-triggered groups."""

    def __init__(self):
        self.auto, self.clicks, self.media, self.text_ids = [], [], [], set()

    def _id(self):
        self.n += 1
        return self.n

    def _effect(self, e, node_type):
        kind = e[0]
        spid = e[1]
        if kind == "play":
            dur = e[2]
            return (f'<p:par><p:cTn id="{self._id()}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" '
                    f'nodeType="{node_type}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                    f'<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="{self._id()}" dur="{dur}" fill="hold"/>'
                    f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>')
        grp = ' grpId="0"' if spid in self.text_ids else ""
        if kind == "appear":
            preset, sub, extra = 1, 0, ""
        elif kind == "wipe":  # e = ("wipe", spid, dur_ms, "up"|"down"|"left"|"right")
            dur, direction = e[2], e[3]
            preset, sub = 22, {"up": 4, "down": 1, "right": 8, "left": 2}[direction]
            extra = (f'<p:animEffect transition="in" filter="wipe({direction})"><p:cBhvr><p:cTn id="{{ID}}" dur="{dur}"/>'
                     f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>')
        cid = self._id()
        setid = self._id()
        xml = (f'<p:par><p:cTn id="{cid}" presetID="{preset}" presetClass="entr" presetSubtype="{sub}" fill="hold"{grp} '
               f'nodeType="{node_type}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
               f'<p:set><p:cBhvr><p:cTn id="{setid}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
               f'</p:cTn><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility'
               f'</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>')
        if extra:
            xml += extra.replace("{ID}", str(self._id()))
        return xml + "</p:childTnLst></p:cTn></p:par>"

    def _group(self, effects, on_click):
        conds = '<p:cond delay="indefinite"/>' + ("" if on_click else '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>')
        outer, inner = self._id(), self._id()
        body = "".join(self._effect(e, ("clickEffect" if (on_click and i == 0) else "withEffect"))
                       for i, e in enumerate(effects))
        return (f'<p:par><p:cTn id="{outer}" fill="hold"><p:stCondLst>{conds}</p:stCondLst><p:childTnLst>'
                f'<p:par><p:cTn id="{inner}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'{body}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')

    def apply(self, slide):
        sld = slide._element
        old = sld.find(qn("p:timing"))
        if old is not None:
            sld.remove(old)
        if not (self.auto or self.clicks or self.media):
            return
        self.n = 2
        groups = ""
        if self.auto:
            groups += self._group(self.auto, on_click=False)
        for c in self.clicks:
            groups += self._group(c, on_click=True)
        seq = ""
        if groups:
            seq = ('<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
                   f'{groups}</p:childTnLst></p:cTn>'
                   '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
                   '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
                   '</p:seq>')
        vids = "".join(
            f'<p:video><p:cMediaNode vol="80000"><p:cTn id="{self._id()}" fill="hold" display="0"><p:stCondLst>'
            f'<p:cond delay="indefinite"/></p:stCondLst><p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl>'
            f'<p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn><p:tgtEl><p:spTgt spid="{v}"/></p:tgtEl>'
            f'</p:cMediaNode></p:video>' for v in self.media)
        animated_text = sorted({e[1] for g in [self.auto, *self.clicks] for e in g if e[1] in self.text_ids and e[0] != "play"})
        bld = ("<p:bldLst>" + "".join(f'<p:bldP spid="{i}" grpId="0"/>' for i in animated_text) + "</p:bldLst>"
               if animated_text else "")
        xml = (f'<p:timing xmlns:p="{P_NS}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" '
               f'nodeType="tmRoot"><p:childTnLst>{seq}{vids}</p:childTnLst></p:cTn></p:par></p:tnLst>{bld}</p:timing>')
        el = etree.fromstring(xml)
        ext = sld.find(qn("p:extLst"))
        if ext is not None:
            ext.addprevious(el)
        else:
            sld.append(el)


def auto_advance(slide, seconds, curl=True):
    """Advance automatically after `seconds` (the host can still click/press a key to move on early)."""
    ms = int(seconds * 1000)
    xml = ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
           f'xmlns:p="{P_NS}"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
           f'Requires="p14"><p:transition spd="slow" p14:dur="1250" advTm="{ms}">'
           f'{"<p14:pageCurlDouble/>" if curl else "<p:fade/>"}</p:transition></mc:Choice>'
           f'<mc:Fallback><p:transition spd="slow" advTm="{ms}"><p:fade/></p:transition></mc:Fallback>'
           '</mc:AlternateContent>')
    el = etree.fromstring(xml)
    sld = slide._element
    anchor = sld.find(qn("p:clrMapOvr"))
    if anchor is None:
        anchor = sld.find(qn("p:cSld"))
    anchor.addnext(el)


# ================================================================ slide types
def countdown_logo(slide, timing, seconds):
    """Grey TBT logo with the full-colour logo wiping up over it for `seconds` — the on-screen countdown."""
    picture(slide, dim_logo(), *LOGO_BOX)
    color = picture(slide, ASSETS / "brand/tbt_logo.png", *LOGO_BOX)
    timing.auto.append(("wipe", color.shape_id, int(seconds * 1000), "up"))


def venue_logo(venue):
    """The venue's logo, or the Trivia By Troy logo if the venue has none yet."""
    p = ASSETS / str(venue.get("logo") or "")
    return p if venue.get("logo") and p.exists() else ASSETS / "brand/tbt_logo.png"


def slide_venue_title(deck, venue):
    s = deck.slide()
    picture(s, venue_logo(venue), Inches(4.4), Inches(0.5), Inches(4.5), Inches(4.0))
    text(s, 0, Inches(4.6), W, Inches(2), "Trivia Night", size=110, font=SCRIPT, align=PP_ALIGN.CENTER)


def slide_host(deck, venue, host):
    """Slide 2: "Tonight's host <name>" in the script font."""
    s = deck.slide()
    header(s, "Trivia Night", venue_logo(venue))
    text(s, 0, Inches(1.9), W, Inches(1.2), "Tonight’s host", size=64, font=SCRIPT, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, 0, Inches(3.1), W, Inches(2.2), str(host).strip(), size=130, font=SCRIPT, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, color=GOLD)


def slide_bullets(deck, venue, title, items, size=20):
    s = deck.slide()
    header(s, title, venue_logo(venue))
    bullets(s, Inches(0.6), Inches(1.3), Inches(11), Inches(5.8), items, size=size)


def slide_script(deck, venue, script, sub="", head="Trivia Night"):
    s = deck.slide()
    header(s, head, venue_logo(venue))
    text(s, 0, Inches(2.2), W, Inches(1.8), script, size=96, font=SCRIPT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if sub:
        text(s, 0, Inches(4.6), W, Inches(0.8), sub, size=24, bold=True, italic=True, align=PP_ALIGN.CENTER)


def slide_upcoming(deck, venue, settings):
    s = deck.slide()
    header(s, "Upcoming Trivia Nights", venue_logo(venue))
    text(s, Inches(0.8), Inches(1.6), Inches(11), Inches(3.4), venue["upcoming_nights"], size=22, bold=True,
         italic=True, spacing=14)
    text(s, Inches(0.8), Inches(5.1), Inches(11), Inches(0.5), "Always follow our social media for updated schedule!",
         size=18, italic=True)
    picture(s, ASSETS / "brand/instagram.png", Inches(1.5), Inches(6.0), Inches(0.7), Inches(0.7))
    text(s, Inches(2.3), Inches(6.0), Inches(4), Inches(0.7), settings.get("instagram", ""), size=26, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
    picture(s, ASSETS / "brand/facebook.png", Inches(7.0), Inches(6.0), Inches(0.7), Inches(0.7))
    text(s, Inches(7.8), Inches(6.0), Inches(5), Inches(0.7), settings.get("facebook", ""), size=26, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)


def slide_qr(deck, top, qr, bottom):
    s = deck.slide()
    text(s, 0, Inches(0.3), W, Inches(1.4), top, size=28, bold=True, italic=True, align=PP_ALIGN.CENTER)
    picture(s, qr, Inches(4.9), Inches(1.8), Inches(3.5), Inches(4.0))
    text(s, 0, Inches(6.2), W, Inches(0.8), bottom, size=26, bold=True, italic=True, align=PP_ALIGN.CENTER)


def slide_results(deck, venue, title, lines=10):
    """Placeholder for the host to type scores into (phase 2: fill from a score sheet)."""
    s = deck.slide()
    header(s, title, venue_logo(venue))
    text(s, Inches(0.6), Inches(1.3), Inches(11), Inches(5.8),
         "\n".join(f"{i}.  " for i in range(1, lines + 1)), size=22, bold=True, spacing=4)


def slide_category(deck, cat, first_q, bonus):
    s = deck.slide()
    header(s, f"Questions {first_q} - {first_q + 4}", ASSETS / "brand/tbt_logo.png")
    text(s, Inches(1.2), Inches(2.0), Inches(10.9), Inches(1.8), cat["title"], size=80, font=SCRIPT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if bonus:
        text(s, Inches(0.5), Inches(1.4), Inches(4), Inches(1.0), "Tonight’s\nBonus Category", size=28, font=BRUSH,
             color=GOLD)
        text(s, Inches(9.6), Inches(3.9), Inches(3.4), Inches(0.8), "Double Points!", size=32, font=BRUSH, color=GOLD)
    text(s, Inches(0.5), Inches(4.8), Inches(11.5), Inches(1.9), cat["blurb"], size=16, bold=True,
         anchor=MSO_ANCHOR.BOTTOM)
    text(s, Inches(0.5), Inches(6.75), Inches(6), Inches(0.5), "Good Luck!", size=20, bold=True, italic=True)


def q_media(qq, kind):
    """Return (question_video, question_poster, answer_video|None, answer_poster|None, clip_seconds)."""
    src = ASSETS / qq["media"]
    if kind == "music" or src.suffix.lower() in (".mp3", ".m4a", ".wav", ".aac"):
        qv = still_plus_audio(music_frame(), src, qq["clip_start"], qq["clip_end"])
        art = ASSETS / qq["answer_image"] if qq["answer_image"] else music_frame()
        av = still_plus_audio(art, src, qq["clip_start"], qq["clip_end"])
        return qv, music_frame(), av, art, ffprobe_dur(qv)
    v = trimmed_video(src, qq["clip_start"], qq["clip_end"])
    return v, poster_of(v), None, None, ffprobe_dur(v)


def slide_question(deck, cat, qq, num, settings):
    s = deck.slide()
    t = Timing()
    kind = cat["type"]
    top = question_block(s, num, qq["question"])
    seconds = float(qq["seconds"] or settings.get("default_seconds", 60))
    box = (Inches(2.4), top, Inches(8.5), Inches(7.2) - top)
    if qq["media"]:
        qv, poster, _, _, clip = q_media(qq, kind)
        if not qq["seconds"]:
            seconds = clip + float(settings.get("media_extra_seconds", 30))
        vid = s.shapes.add_movie(str(qv), *fit_box(poster, *box), poster_frame_image=str(poster), mime_type="video/mp4")
        t.auto.append(("play", vid.shape_id, int(clip * 1000)))
        t.media.append(vid.shape_id)
    elif qq["question_image"]:
        picture(s, ASSETS / qq["question_image"], *box)  # shown immediately, no animation
    countdown_logo(s, t, seconds)
    t.apply(s)
    auto_advance(s, seconds)
    return seconds


def slide_answer(deck, cat, qq, num):
    """Question again; one click reveals the answer (and replays the song for music)."""
    s = deck.slide()
    t = Timing()
    kind = cat["type"]
    top = question_block(s, num, qq["question"])
    picture(s, ASSETS / "brand/tbt_logo.png", *LOGO_BOX)
    right = (Inches(5.6), top, Inches(7.4), Inches(7.2) - top)
    reveal = []
    if qq["question_image"] and not qq["media"]:
        picture(s, ASSETS / qq["question_image"], *right)
    ans = text(s, Inches(0.5), top + Inches(0.15), Inches(4.9), Inches(1.4), qq["answer"], size=26, bold=True, color=GOLD)
    t.text_ids.add(ans.shape_id)
    reveal.append(("appear", ans.shape_id))
    left_img = (Inches(0.5), top + Inches(1.6), Inches(4.9), Inches(7.2) - top - Inches(1.6))
    if qq["media"] and (kind == "music" or Path(qq["media"]).suffix.lower() in (".mp3", ".m4a", ".wav", ".aac")):
        _, _, av, art, clip = q_media(qq, kind)
        vid = s.shapes.add_movie(str(av), *fit_box(art, *right), poster_frame_image=str(art), mime_type="video/mp4")
        reveal += [("appear", vid.shape_id), ("play", vid.shape_id, int(clip * 1000))]
        t.media.append(vid.shape_id)
    elif qq["answer_image"]:
        box = left_img if (qq["question_image"] and not qq["media"]) else right
        img = picture(s, ASSETS / qq["answer_image"], *box)
        reveal.append(("appear", img.shape_id))
    t.clicks.append(reveal)
    t.apply(s)


def slide_tiebreaker(deck, venue, tbs, with_answers):
    s = deck.slide()
    text(s, 0, Inches(0.1), W, Inches(1.1), "Tiebreaker", size=66, font=SCRIPT, align=PP_ALIGN.CENTER)
    picture(s, venue_logo(venue), *LOGO_BOX)
    rules = ("Participating teams should move so that all teams and host are in close proximity – failure to move "
             "results in a forfeit.\nTeams write down their answers for each of these years. The host adds up how many "
             "years off each answer is; whichever team's total is lowest WINS!")
    text(s, Inches(0.6), Inches(1.3), Inches(12), Inches(1.6), rules, size=15, bold=True, italic=True, spacing=6)
    lines = []
    for i, tb in enumerate(sorted(tbs, key=lambda r: int(r["order"] or 0)), 1):
        lines.append(f"{i}.  {tb['question']}")
        if with_answers:
            lines.append(f"       Answer: {tb['answer']}")
    text(s, Inches(0.6), Inches(3.2), Inches(12), Inches(4), "\n".join(lines), size=20, bold=True, spacing=6)


# =============================================================== deck assembly
def build_deck(bank, row, outdir):
    venue = bank["venues"][row["venue_id"]]
    settings = bank["settings"]
    set_id = row["set_id"]
    slots = sorted(bank["sets"][set_id], key=lambda r: (int(r["round"]), int(r["slot"])))
    deck = Deck()
    host_txt = str(row.get("host", "") or "").strip()
    key_lines = [f"ANSWER KEY — {venue['name']} — {as_date(row['date'])} — set {set_id}"
                 + (f" — host {host_txt}" if host_txt else ""), ""]
    total_secs = 0

    slide_venue_title(deck, venue)
    host = str(row.get("host", "") or "").strip()
    if host:
        slide_host(deck, venue, host)
    slide_bullets(deck, venue, "Trivia Format", [
        "3 Rounds of 15 Questions each.", "60 Seconds for each Question.", "Questions will not be repeated.",
        "Answers will be given and scored after each Round.", "Short intermission after each Round.",
        "At the end of 3 Rounds the team with the most points wins!"], size=24)
    slide_bullets(deck, venue, "Trivia Rules", [
        "No Cheating!|Please keep phones, notes, watches, etc. away during questions.",
        "6 Members Max Per Team.|1 point penalty for each member over 6.",
        "My Judgement Is Final.|The host has final say on challenges, right or wrong!",
        "Buy Food and Drinks, and Tip Your Servers!|Support small business and those working hard for you tonight.",
        "Have Fun!|Remember, this is a fun event; be respectful and enjoy yourself!"], size=22)
    slide_upcoming(deck, venue, settings)

    for rnd in (1, 2, 3):
        rslots = [r for r in slots if int(r["round"]) == rnd]
        slide_script(deck, venue, f"Round {NUM_WORDS[rnd]}", ROUND_SUB[rnd])
        for sl in rslots:
            cat = bank["cats"][sl["category_id"]]
            first = (int(sl["slot"]) - 1) * 5 + 1
            bonus = str(sl["bonus"]).upper().startswith("Y")
            slide_category(deck, cat, first, bonus)
            for i, qq in enumerate(cat_questions(bank, cat["category_id"])):
                total_secs += slide_question(deck, cat, qq, first + i, settings)
        slide_script(deck, venue, "End of Round", "Turn in your answer sheets…")
        slide_script(deck, venue, f"Round {NUM_WORDS[rnd]}", "Answers — exchange your answer sheets…")
        key_lines.append(f"ROUND {rnd}")
        for sl in rslots:
            cat = bank["cats"][sl["category_id"]]
            first = (int(sl["slot"]) - 1) * 5 + 1
            bonus = str(sl["bonus"]).upper().startswith("Y")
            slide_category(deck, cat, first, bonus)
            key_lines.append(f"  {cat['title']}{'  (BONUS – double points)' if bonus else ''}")
            for i, qq in enumerate(cat_questions(bank, cat["category_id"])):
                slide_answer(deck, cat, qq, first + i)
                key_lines.append(f"    {first + i:>2}. {qq['answer']}")
        key_lines.append("")
        slide_results(deck, venue, f"Round {NUM_WORDS[rnd]} Results")
        if rnd == 1:
            slide_script(deck, venue, "Intermission")
        elif rnd == 2:
            slide_results(deck, venue, "Standings as of Round Two")
            slide_qr(deck, "Are you having fun?\nWant to Tip your Host?", ASSETS / "brand/tip_qr.png",
                     "We’ll be back with Round 3 soon!")
        else:
            slide_results(deck, venue, "Final Standings")

    tbs = bank["tiebreakers"].get(set_id, [])
    if tbs:
        slide_tiebreaker(deck, venue, tbs, with_answers=False)
        slide_tiebreaker(deck, venue, tbs, with_answers=True)
        key_lines.append("TIEBREAKER")
        key_lines += [f"    {t['question']}  ->  {t['answer']}" for t in tbs]
    slide_qr(deck, "Did you have a good time?\nWant to Tip your Host?", ASSETS / "brand/tip_qr.png",
             "Thanks for playing, see you soon!")
    s = deck.slide()
    picture(s, ASSETS / "brand/feedback_qr.jpg", 0, 0, W, H)

    outdir.mkdir(parents=True, exist_ok=True)
    stem = f"{as_date(row['date'])} {venue['name']}" + (f" ({host_txt})" if host_txt else "")
    deck.prs.save(outdir / f"{stem}.pptx")
    (outdir / f"{stem} - answer key.txt").write_text("\n".join(key_lines) + "\n")
    return outdir / f"{stem}.pptx", total_secs


def usage_rows(bank, schedule):
    """One Usage row per question played: date, venue, set, category, question (skipping ones already logged)."""
    have = {(as_date(u["date"]), u["venue_id"], u["set_id"], u["category_id"], u.get("question_id", ""))
            for u in bank["Usage"]}
    out = []
    for r in schedule:
        d = as_date(r["date"])
        for sl in bank["sets"][r["set_id"]]:
            for l in bank["links"].get(sl["category_id"], []):
                k = (d, r["venue_id"], r["set_id"], sl["category_id"], l["question_id"])
                if k not in have:
                    have.add(k)
                    out.append(k)
    return out


def record_usage_xlsx(rows, path):
    wb = load_workbook(path)  # keep formulas
    ws = wb["Usage"]
    for k in rows:
        ws.append(list(k))
        ws.cell(ws.max_row, 1).number_format = "yyyy-mm-dd"
    wb.save(path)


def record_usage_sheet(rows, cfg):
    sheet_call(cfg, "append_usage", {"rows": [[str(k[0]), *k[1:]] for k in rows]})


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--bank", help="use this .xlsx file instead of the Google Sheet")
    ap.add_argument("--out", default=str(ROOT / "out"))
    ap.add_argument("--date", help="only this date (yyyy-mm-dd)")
    ap.add_argument("--week", help="the Mon-Sun week containing this date (yyyy-mm-dd), or 'this' / 'next'")
    ap.add_argument("--day", help="only these weekdays, e.g. wed or tue,thu")
    ap.add_argument("--from", dest="date_from", help="first date (yyyy-mm-dd)")
    ap.add_argument("--to", dest="date_to", help="last date (yyyy-mm-dd)")
    ap.add_argument("--venue", help="only these venue_ids, e.g. wine_press or pops,narrative")
    ap.add_argument("--only", help="exact nights as date:venue_id pairs, e.g. 2026-12-02:wine_press,2026-12-02:pops")
    ap.add_argument("--record", action="store_true", help="log the questions used to the Usage tab")
    ap.add_argument("--check", action="store_true", help="only run checks, build nothing")
    ap.add_argument("--strict", action="store_true", help="treat warnings as errors")
    ap.add_argument("--video", metavar="SET_ID", help="make the at-home video for this set (out/video/)")
    ap.add_argument("--video-seconds", type=int, default=30, help="video: seconds per question (default 30)")
    ap.add_argument("--video-answer-seconds", type=float, default=6, help="video: seconds per answer (default 6)")
    ap.add_argument("--video-height", type=int, default=720, choices=[720, 1080], help="video: 720 or 1080")
    ap.add_argument("--video-no-media", action="store_true", help="video: leave out music/movie categories")
    ap.add_argument("--ids", action="store_true", help="(.xlsx only) fill in missing question/category IDs, then stop")
    a = ap.parse_args()

    cfg = None if a.bank else load_config()
    if cfg:
        print("Reading the Google Sheet…")
        try:
            tables = tables_from_sheet(cfg)
        except SheetError as e:
            sys.exit(f"Google Sheet: {e}")
        if a.ids:
            sys.exit("In the Google Sheet, IDs are assigned by the TBTgen script as you type "
                     "(or TBTgen > Assign missing IDs).")
    else:
        bank_path = a.bank or str(ROOT / "trivia_bank.xlsx")
        if not a.bank:
            print(f"No {CONFIG.name} found: using {Path(bank_path).name}")
        try:
            n = assign_ids(bank_path)
        except PermissionError:
            sys.exit("Can't write to the spreadsheet - close it in Excel/Numbers and try again.")
        if n or a.ids:
            print(f"Assigned {n} new IDs (saved to the spreadsheet).")
        if a.ids:
            return
        tables = tables_from_xlsx(bank_path)
    bank = load_bank(tables)
    missing = [r for r in bank["Questions"] if not r.get("question_id") and r.get("question")]
    if missing:
        print(f"NOTE: {len(missing)} question(s) have no ID yet and can't be used until they do.")

    if a.video:
        if a.video not in bank["sets"]:
            sys.exit(f"Set '{a.video}' is not on the Sets tab.")
        errs, warns = check_sets(bank, {a.video})
        for w in warns:
            print("WARNING:", w)
        for e in errs:
            print("ERROR:  ", e)
        if errs:
            sys.exit("Stopped: fix the problems above.")
        import video
        try:
            path, secs = video.make_video(bank, a.video, a.out, seconds=a.video_seconds,
                                          answer_seconds=a.video_answer_seconds, height=a.video_height,
                                          media=not a.video_no_media)
        except RuntimeError as e:
            sys.exit(f"Video failed: {e}")
        print(f"built  {Path(path).relative_to(ROOT) if Path(path).is_relative_to(ROOT) else path}   "
              f"(length {int(secs // 60)}:{int(secs % 60):02d})")
        return
    lo = as_date(a.date_from) if a.date_from else None
    hi = as_date(a.date_to) if a.date_to else None
    if a.date:
        lo = hi = as_date(a.date)
    if a.week:
        today = dt.date.today()
        ref = {"this": today, "next": today + dt.timedelta(days=7)}.get(a.week.lower()) or as_date(a.week)
        lo = ref - dt.timedelta(days=ref.weekday())
        hi = lo + dt.timedelta(days=6)
    days = None
    if a.day:
        names = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
        try:
            days = {names.index(d.strip().lower()[:3]) for d in a.day.split(",")}
        except ValueError:
            sys.exit("--day takes weekday names like mon, tue, wed (comma-separated)")
    venues = {v.strip() for v in a.venue.split(",")} if a.venue else None
    only = None
    if a.only:
        try:
            only = {(as_date(p.split(":")[0]), p.split(":")[1].strip()) for p in a.only.split(",") if p.strip()}
        except (ValueError, IndexError):
            sys.exit("--only takes date:venue_id pairs, e.g. 2026-12-02:wine_press")
    rows = [r for r in bank["Schedule"]
            if r.get("date")
            and (only is None or (as_date(r["date"]), r["venue_id"]) in only)
            and (venues is None or r["venue_id"] in venues)
            and (lo is None or as_date(r["date"]) >= lo)
            and (hi is None or as_date(r["date"]) <= hi)
            and (days is None or as_date(r["date"]).weekday() in days)]
    if lo or hi:
        print(f"Dates: {lo or '…'} to {hi or '…'}" + (f", {a.day} only" if a.day else ""))
    schedule = []
    for r in rows:
        label = f"{as_date(r['date'])} {r.get('calendar_event') or r['venue_id']}"
        if not r["venue_id"] or r["venue_id"] not in bank["venues"]:
            print(f"SKIPPED: {label}: venue '{r['venue_id']}' is not on the Venues tab")
        elif not r["set_id"]:
            print(f"SKIPPED: {label}: no set assigned yet")
        else:
            schedule.append(r)
    if not schedule:
        sys.exit("Nothing to build: no matching Schedule rows with a venue and a set.")
    for r in schedule:
        v = bank["venues"][r["venue_id"]]
        if not (v.get("logo") and (ASSETS / str(v["logo"])).exists()):
            print(f"NOTE: {v['name']} has no logo file yet; using the Trivia By Troy logo")

    errs, warns = check(bank, schedule)
    for w in warns:
        print("WARNING:", w)
    for e in errs:
        print("ERROR:  ", e)
    if errs or (a.strict and warns):
        sys.exit("Stopped: fix the problems above (or drop --strict).")
    if a.check:
        return

    for r in schedule:
        path, secs = build_deck(bank, r, Path(a.out))
        print(f"built  {path.relative_to(ROOT) if path.is_relative_to(ROOT) else path}   "
              f"(question time ≈ {secs / 60:.0f} min)")
    if a.record:
        rows = usage_rows(bank, schedule)
        if rows:
            try:
                record_usage_sheet(rows, cfg) if cfg else record_usage_xlsx(rows, bank_path)
            except SheetError as e:
                sys.exit(f"Decks were built, but logging to the Usage tab failed: {e}")
        print(f"Usage: recorded {len(rows)} new rows")


if __name__ == "__main__":
    main()
