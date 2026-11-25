#!/usr/bin/env python3
"""
Import questions and categories from old Keynote (.key) / PowerPoint (.pptx) game decks.

  python3 import_decks.py FOLDER [FOLDER ...] [--pictures] [--clips] [--out imports/NAME]

Scans the folders (and their sub-folders) for decks, recognises the Trivia By Troy slide patterns
("Questions 1 - 5" category intros, "Question N" slides, answer slides that repeat the question with the
answer added), pairs answers with questions, and merges repeats across decks. Nothing is written to the
question bank: it produces a review package that the app's Import tab (or you) can add from.

Output (in imports/<date-time>/ unless --out is given):
  import.json   everything found, for the app's Import tab
  review.xlsx   Decks / Categories / Questions / Check-these tabs to look over
  assets/       pictures (--pictures) and song/movie clips (--clips), named by content so repeats are stored once

Without --pictures / --clips, media is only counted (so you can see how much space it would take).
"""
import argparse
import datetime as dt
import difflib
import hashlib
import json
import re
import shutil
import subprocess
import sys
from collections import Counter, defaultdict
from pathlib import Path

import keynote_read as kr

ROOT = Path(__file__).resolve().parent
CONVERT = {"tif", "tiff", "jxl", "heic", "heif", "bmp", "webp"}   # turned into .jpg/.png on import
WORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "1": 1, "2": 2, "3": 3, "4": 4, "5": 5}
MONTHS = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}

RE_INTRO = re.compile(r"^questions?\s*#?\s*(\d+)\s*(?:-|–|—|to|thru|through)\s*(\d+)$", re.I)
RE_QUESTION = re.compile(r"^question\s*#?\s*(\d+)$", re.I)
RE_ROUND = re.compile(r"^(?:round\s+(one|two|three|four|five|[1-5])|(final)\s+round)$", re.I)
LABELS = re.compile(r"^(good luck!?|tonight.?s\s+bonus\s+category|bonus\s+category|double points!?|trivia night|"
                    r"one points? (each|awarded|for (the )?(title|artist|song|each)).*|answers?|exchange your answer sheets.*)$", re.I | re.S)
NOTE_SKIP = re.compile(r"^(good luck!?|tonight.?s\s+bonus\s+category|bonus\s+category|double points!?|trivia night|"
                       r"answers?|exchange your answer sheets.*|trivia by troy)$", re.I | re.S)
STRUCTURE = re.compile(r"(end of round|intermission|results|standings|trivia (rules|format|night seasons)|tiebreaker|"
                       r"upcoming trivia|tip your host|having fun|how did we do|thanks for playing|season \d+|"
                       r"round (one|two|three|four)|good luck|hang in there|comes down to this)", re.I)


def norm(s):
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", str(s).lower().replace("’", "'"))).strip()


def similar(a, b):
    a, b = norm(a), norm(b)
    if not a or not b:
        return False
    short, long_ = sorted((a, b), key=len)
    # one contains the other AND they're of similar length (so a short answer like "Home Alone" that happens
    # to appear inside its question doesn't count as a repeat of the question)
    if short in long_ and len(short) >= 0.6 * len(long_):
        return True
    return difflib.SequenceMatcher(None, a, b).ratio() > 0.8


def qkey(question, answer):
    return hashlib.sha1(f"{norm(question)}|{norm(answer)}".encode()).hexdigest()[:12]


# ------------------------------------------------------------------------------------------------ dates
def guess_date(path):
    s = " ".join([path.stem, path.parent.name])
    pats = [
        (r"(20\d\d)[-_. ](\d{1,2})[-_. ](\d{1,2})", lambda m: (int(m[1]), int(m[2]), int(m[3]))),
        (r"(?<!\d)(\d{1,2})[-_. ](\d{1,2})[-_. ](20\d\d)", lambda m: (int(m[3]), int(m[1]), int(m[2]))),
        (r"(?<!\d)(\d{1,2})[-_.](\d{1,2})[-_.](\d\d)(?!\d)", lambda m: (2000 + int(m[3]), int(m[1]), int(m[2]))),
        (r"(20\d\d)(\d\d)(\d\d)", lambda m: (int(m[1]), int(m[2]), int(m[3]))),
        (r"([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(20\d\d)",
         lambda m: (int(m[3]), MONTHS.get(m[1][:3].lower(), 0), int(m[2]))),
    ]
    for pat, fn in pats:
        for m in re.finditer(pat, s):
            try:
                y, mo, d = fn(m)
                return str(dt.date(y, mo, d)), "file name"
            except (ValueError, TypeError):
                continue
    return str(dt.date.fromtimestamp(path.stat().st_mtime)), "file date"


# ------------------------------------------------------------------------------------------- per deck
def parse_deck(path, venues=()):
    slides = kr.read_deck(path)
    deck = {"file": str(path), "name": path.name, "slides": len(slides), "problems": [], "unknown": []}
    deck["date"], deck["date_source"] = guess_date(path)
    low = str(path).lower()
    deck["venue_guess"] = next((vid for vid, names in venues if any(n and n.lower() in low for n in names)), "")

    # pictures that appear on many slides are logos/backgrounds, not question pictures
    img_count = Counter(n for s in slides for n in {x.name for x in s.images})
    recurring = {n for n, c in img_count.items() if c >= max(4, len(slides) * 0.2)}

    rnd, answers_phase, implicit = 1, False, 0
    cat = None
    cats, qmap, counts = [], {}, Counter()
    for s in slides:
        texts = [t.strip() for t in s.texts if t.strip()]
        flat = [t.replace("\n", " ").strip() for t in texts]
        imgs = [i for i in s.images if i.name not in recurring]
        # round markers / answers section
        for t in flat:
            m = RE_ROUND.match(t)
            if m:
                rnd = WORDS.get((m.group(1) or "").lower(), rnd + 1) if m.group(1) else rnd + 1
                answers_phase = any(x.lower().startswith("answers") for x in flat)
                cat = None
                break
        intro = next((RE_INTRO.match(t) for t in flat if RE_INTRO.match(t)), None)
        qhead = next((RE_QUESTION.match(t) for t in flat if RE_QUESTION.match(t)), None)
        if intro:
            counts["intro"] += 1
            first, last = int(intro.group(1)), int(intro.group(2))
            rest = [t for t in texts if not RE_INTRO.match(t.replace("\n", " ").strip())]
            cand = [t for t in rest if not LABELS.match(t.replace("\n", " ").strip())]
            title = min((t for t in cand if "\n" not in t.strip() and len(t) <= 60), key=len, default="")
            blurb = max((t for t in cand if t != title), key=len, default="")
            blurb = re.sub(r"\s*good luck!?\s*$", "", blurb, flags=re.I).strip()
            bonus = any("bonus" in t.lower() for t in texts)
            existing = next((c for c in cats if c["round"] == rnd and c["first"] == first and norm(c["title"]) == norm(title)), None)
            if existing:            # the intro shown again before the answers
                cat = existing
                continue
            cat = {"title": title or f"Round {rnd} questions {first}-{last}", "blurb": blurb, "bonus": bonus,
                   "round": rnd, "first": first, "last": last, "questions": {}}
            cats.append(cat)
            continue
        if qhead:
            num = int(qhead.group(1))
            # each text box -> its main text; styled subtext (italic/smaller, e.g. "One Point Each!") kept aside
            split = {t.strip(): p for t, p in zip(s.texts, s.parts or [])}
            subs, orig = {}, {}
            mains = []
            for t in texts:
                main, sub = split.get(t, (t, ""))
                main = main or t
                mains.append(main)
                orig[main] = t
                if sub:
                    subs[main] = sub
            texts = mains
            rest = [t for t in texts if not RE_QUESTION.match(t.replace("\n", " ").strip())]
            body = [t for t in rest if not LABELS.match(t.replace("\n", " ").strip())]
            key = (rnd + implicit, num)
            prev = qmap.get(key)
            if prev and not answers_phase and body and not any(similar(b, prev["question"]) for b in body):
                implicit += 1           # same number again with a different question: a new round without a marker
                key, prev = (rnd + implicit, num), None
            if prev is None:
                counts["question"] += 1
                question = max(body, key=len, default="")
                extra_q = ([subs[question]] if question in subs else []) + \
                    [t for t in rest if t != question and not NOTE_SKIP.match(t.replace("\n", " ").strip())]
                q = {"number": num, "round": key[0], "slide": s.number, "question": question, "answer": "",
                     "question_note": "\n".join(extra_q), "answer_note": "",
                     "question_image": imgs[0] if imgs else None, "answer_image": None,
                     "media": s.media[0] if s.media else None}
                qmap[key] = q
                owner = cat if cat and cat["first"] <= num <= cat["last"] else None
                if owner is None:
                    owner = {"title": f"Round {key[0]} (no category intro)", "blurb": "", "bonus": False, "round": key[0],
                             "first": num, "last": num, "questions": {}}
                    cats.append(owner)
                    cat = owner
                owner["questions"][num] = q
            else:
                counts["answer"] += 1
                extra = [b for b in body if not similar(b, prev["question"])]
                if extra:
                    prev["answer"] = min(extra, key=len) if len(extra) > 1 and all(len(e) < 80 for e in extra) else extra[0]
                    prev["answer"] = orig.get(prev["answer"], prev["answer"])  # an answer box is kept whole: "Title (1987)"
                    prev["answer_slide"] = s.number
                # any other text on the answer slide (scoring notes, fun facts) that wasn't on the question slide
                qtexts = [prev["question"], prev.get("question_note", "")]
                others = [t for t in rest if t != prev["answer"] and not any(similar(t, x) for x in qtexts if x)
                          and not NOTE_SKIP.match(t.replace("\n", " ").strip())]
                others = [orig.get(t, t) for t in others if orig.get(t, t) != prev["answer"]]
                if others:
                    prev["answer_note"] = "\n".join(others)
                qnames = {prev["question_image"].name} if prev["question_image"] else set()
                new_imgs = [i for i in imgs if i.name not in qnames]
                if new_imgs:
                    prev["answer_image"] = new_imgs[0]
                if not prev["media"] and s.media:
                    prev["media"] = s.media[0]
            continue
        counts["other"] += 1
        joined = " ".join(flat)
        if joined and "?" in joined and not STRUCTURE.search(joined):
            deck["unknown"].append({"slide": s.number, "text": joined[:200]})

    # finish categories
    out = []
    for c in cats:
        qs = [c["questions"][n] for n in sorted(c["questions"])]
        if not qs:
            continue
        kinds = {Path(q["media"].name).suffix.lower() for q in qs if q["media"]}
        blob = f"{c['title']} {c['blurb']} " + " ".join(q["question"] for q in qs)
        if kinds & {".mp3", ".m4a", ".wav", ".aac", ".aif", ".aiff"} or re.search(r"name (this|these) songs?", blob, re.I):
            ctype = "music"
        elif kinds & {".mp4", ".mov", ".m4v", ".avi"} or re.search(r"what movie is this|name the movies", blob, re.I):
            ctype = "movie"
        else:
            ctype = "standard"
        for q in qs:
            if not q["answer"]:
                deck["problems"].append(f"slide {q['slide']}: Question {q['number']} (round {q['round']}) has no answer slide")
        if len(qs) != 5:
            deck["problems"].append(f"category '{c['title']}' has {len(qs)} questions")
        out.append({**c, "type": ctype, "questions": qs})
    deck["counts"] = dict(counts)
    deck["categories"] = out
    if not out:
        deck["problems"].append("no questions recognised (not a Trivia By Troy game deck?)")
    return deck


# ------------------------------------------------------------------------------------------- media
def save_media(ref, folder, kind):
    """Writes one picture/clip into folder (content-named, so repeats are stored once); returns 'kind/name'."""
    data = ref.read()
    digest = hashlib.sha1(data).hexdigest()[:16]
    ext = ref.ext or "bin"
    target_dir = folder / kind
    target_dir.mkdir(parents=True, exist_ok=True)
    if kind == "images" and ext in CONVERT:
        for done in target_dir.glob(digest + ".*"):
            return f"{kind}/{done.name}", None
        src = target_dir / f"{digest}.src.{ext}"
        src.write_bytes(data)
        try:
            from PIL import Image
            with Image.open(src) as im:
                alpha = im.mode in ("RGBA", "LA") or "transparency" in im.info
                out = target_dir / f"{digest}.{'png' if alpha else 'jpg'}"
                (im if alpha else im.convert("RGB")).save(out, quality=90)
            src.unlink()
            return f"{kind}/{out.name}", None
        except Exception:
            out = target_dir / f"{digest}.jpg"
            if shutil.which("sips"):  # macOS can convert JPEG XL / HEIC / TIFF
                r = subprocess.run(["sips", "-s", "format", "jpeg", str(src), "--out", str(out)], capture_output=True)
                if r.returncode == 0 and out.exists():
                    src.unlink()
                    return f"{kind}/{out.name}", None
            final = target_dir / f"{digest}.{ext}"
            src.rename(final)
            return f"{kind}/{final.name}", f"couldn't convert {ref.name} (.{ext}) to JPEG; kept as is"
    target = target_dir / f"{digest}.{ext}"
    if not target.exists():
        target.write_bytes(data)
    return f"{kind}/{target.name}", None


# ------------------------------------------------------------------------------------------- run
def load_bank_quietly():
    try:
        import build
        cfg = build.load_config()
        tables = build.tables_from_sheet(cfg) if cfg else build.tables_from_xlsx(ROOT / "trivia_bank.xlsx")
        return build.load_bank(tables)
    except Exception as e:
        print(f"NOTE: couldn't read the question bank ({e}); duplicates won't be checked against it")
        return None


ICLOUD_STUB = re.compile(r"^\.(.+\.(?:key|pptx))\.icloud$", re.I)


def find_decks(folders):
    """Returns (decks, cloud_only): cloud_only are decks iCloud hasn't downloaded to this Mac (older macOS shows
    them as hidden ".Name.key.icloud" placeholder files, which can't be read)."""
    found, cloud_only = [], []
    for f in folders:
        p = Path(f).expanduser()
        if p.is_file() and p.suffix.lower() in (".key", ".pptx"):
            found.append(p)
            continue
        if not p.exists():
            print(f"NOTE: {p} doesn't exist")
            continue
        for x in sorted(p.rglob("*")):
            stub = ICLOUD_STUB.match(x.name)
            if stub:
                cloud_only.append(x.with_name(stub.group(1)))
                continue
            if x.suffix.lower() in (".key", ".pptx") and not x.name.startswith(("~$", "._")):
                if any(part.endswith(".key") for part in x.relative_to(p).parts[:-1]):
                    continue  # inside a .key package folder
                found.append(x)
    return found, cloud_only


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("folders", nargs="+", help="folders (or files) to scan for .key / .pptx decks")
    ap.add_argument("--pictures", action="store_true", help="also import question/answer pictures")
    ap.add_argument("--clips", action="store_true", help="also import song and movie clips")
    ap.add_argument("--out", help="output folder (default imports/<date-time>)")
    ap.add_argument("--no-bank", action="store_true", help="don't compare with the current question bank")
    a = ap.parse_args()

    decks, cloud_only = find_decks(a.folders)
    if cloud_only:
        print(f"WARNING: {len(cloud_only)} deck(s) are in iCloud but not downloaded to this Mac, so they're skipped:")
        for c in cloud_only[:10]:
            print(f"   {c}")
        if len(cloud_only) > 10:
            print(f"   … and {len(cloud_only) - 10} more")
        print("   In Finder, right-click the folder > Download Now (or Keep Downloaded), wait, then scan again.")
    if not decks:
        sys.exit("No .key or .pptx files found." + (" (Only iCloud placeholders: download them first.)" if cloud_only else ""))
    out = Path(a.out) if a.out else ROOT / "imports" / dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    out.mkdir(parents=True, exist_ok=True)
    bank = None if a.no_bank else load_bank_quietly()
    venues = [(vid, [str(v.get("name") or ""), str(v.get("calendar_name") or "")]) for vid, v in (bank or {}).get("venues", {}).items() if vid]
    existing_q = {}
    if bank:
        for qid, q in bank["qs"].items():
            existing_q[qkey(q["question"], q["answer"])] = qid

    print(f"Found {len(decks)} deck(s). Reading…", flush=True)
    questions, categories, deck_rows = {}, {}, []
    media = {"images": [0, 0, set()], "clips": [0, 0, set()]}   # count, bytes, unique names
    for i, path in enumerate(decks, 1):
        print(f"deck: {i}/{len(decks)} {path.name}", flush=True)
        try:
            d = parse_deck(path, venues)
        except Exception as e:
            deck_rows.append({"file": str(path), "name": path.name, "error": str(e), "problems": [f"couldn't read: {e}"],
                              "unknown": [], "counts": {}, "slides": 0, "date": "", "date_source": "", "venue_guess": ""})
            continue
        appearance = {"deck": path.name, "date": d["date"], "venue": d["venue_guess"]}
        cat_keys = []
        for c in d["categories"]:
            keys = []
            for q in c["questions"]:
                k = qkey(q["question"], q["answer"])
                rec = questions.get(k)
                if rec is None:
                    rec = questions[k] = {"key": k, "question": q["question"], "answer": q["answer"],
                                          "question_note": q.get("question_note", ""), "answer_note": q.get("answer_note", ""),
                                          "question_image": "",
                                          "answer_image": "", "media": "", "media_names": [], "appearances": [],
                                          "existing_id": existing_q.get(k, "")}
                rec["appearances"].append(appearance)
                for field, kind, want in (("question_image", "images", a.pictures), ("answer_image", "images", a.pictures),
                                          ("media", "clips", a.clips)):
                    ref = q.get(field)
                    if not ref or rec[field]:
                        continue
                    if ref.name not in media[kind][2]:
                        media[kind][2].add(ref.name)
                        media[kind][0] += 1
                        media[kind][1] += ref.size
                    rec["media_names"].append(ref.name)
                    if want:
                        try:
                            rel, warn = save_media(ref, out / "assets", "images" if kind == "images" else "media")
                            rec[field] = rel
                            if warn:
                                d["problems"].append(warn)
                        except Exception as e:
                            d["problems"].append(f"couldn't copy {ref.name}: {e}")
                keys.append(k)
            ck = hashlib.sha1((norm(c["title"]) + "|" + "|".join(keys)).encode()).hexdigest()[:12]
            crec = categories.get(ck)
            if crec is None:
                crec = categories[ck] = {"key": ck, "title": c["title"], "blurb": c["blurb"], "type": c["type"],
                                         "bonus": c["bonus"], "questions": keys, "appearances": []}
            crec["appearances"].append(appearance)
            cat_keys.append(ck)
        d.pop("categories")
        d["category_keys"] = cat_keys
        deck_rows.append(d)

    for c in categories.values():  # already in the bank? (same title and all the same questions)
        c["existing"] = all(questions[k]["existing_id"] for k in c["questions"])
    result = {"created": dt.datetime.now().isoformat(timespec="seconds"), "folders": a.folders,
              "options": {"pictures": a.pictures, "clips": a.clips}, "decks": deck_rows,
              "categories": sorted(categories.values(), key=lambda c: (c["appearances"][0]["date"], c["title"])),
              "questions": questions,
              "media": {k: {"count": v[0], "bytes": v[1], "imported": (a.pictures if k == "images" else a.clips)}
                        for k, v in media.items()}}
    (out / "import.json").write_text(json.dumps(result, indent=1))
    write_review(result, out / "review.xlsx")
    nq = len(questions)
    new_q = sum(1 for q in questions.values() if not q["existing_id"])
    probs = sum(len(d["problems"]) for d in deck_rows)
    unk = sum(len(d["unknown"]) for d in deck_rows)
    mb = lambda b: f"{b / 1e6:.1f} MB"
    print(f"Read {len(decks)} deck(s): {len(categories)} categories, {nq} questions ({new_q} new, {nq - new_q} already in the bank).")
    print(f"Pictures: {media['images'][0]} ({mb(media['images'][1])}){'' if a.pictures else ' - not imported'}. "
          f"Clips: {media['clips'][0]} ({mb(media['clips'][1])}){'' if a.clips else ' - not imported'}.")
    if probs or unk:
        print(f"Check: {probs} problem(s) and {unk} slide(s) that look like questions but weren't recognised (see review.xlsx).")
    print(f"built  {out.relative_to(ROOT) if out.is_relative_to(ROOT) else out}/review.xlsx")
    print(f"IMPORT_DIR {out}")


def write_review(r, path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    head = PatternFill("solid", fgColor="1F2937")

    def tab(title, cols, rows, widths):
        ws = wb.create_sheet(title)
        ws.append(cols)
        for c in ws[1]:
            c.font, c.fill = Font(bold=True, color="FFFFFF"), head
        for row in rows:
            ws.append(row)
            for c in ws[ws.max_row]:  # deck text that starts with "=" stays text, never a formula
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.data_type = "s"
        for i, w in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = w
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = "A2"

    wb.remove(wb.active)
    tab("Decks", ["file", "date", "date from", "venue guess", "slides", "category intros", "question slides", "answer slides",
                  "other slides", "problems"],
        [[d["name"], d["date"], d["date_source"], d["venue_guess"], d["slides"], d["counts"].get("intro", 0),
          d["counts"].get("question", 0), d["counts"].get("answer", 0), d["counts"].get("other", 0), len(d["problems"])]
         for d in r["decks"]], [40, 12, 12, 14, 8, 10, 10, 10, 10, 9])
    q = r["questions"]
    tab("Categories", ["key", "title", "type", "bonus", "questions", "times used", "first used", "last used", "already in bank",
                       "blurb"],
        [[c["key"], c["title"], c["type"], "Y" if c["bonus"] else "", len(c["questions"]), len(c["appearances"]),
          min(a["date"] for a in c["appearances"]), max(a["date"] for a in c["appearances"]), "Y" if c["existing"] else "",
          c["blurb"]] for c in r["categories"]], [14, 34, 10, 7, 10, 10, 12, 12, 12, 60])
    cat_of = defaultdict(list)
    for c in r["categories"]:
        for k in c["questions"]:
            cat_of[k].append(c["title"])
    tab("Questions", ["key", "question", "answer", "question note", "answer note", "categories", "times used",
                      "first used", "last used", "in bank as", "picture", "answer picture", "clip"],
        [[x["key"], x["question"], x["answer"], x.get("question_note", ""), x.get("answer_note", ""),
          ", ".join(dict.fromkeys(cat_of[x["key"]])), len(x["appearances"]),
          min(a["date"] for a in x["appearances"]), max(a["date"] for a in x["appearances"]), x["existing_id"],
          x["question_image"], x["answer_image"], x["media"]] for x in q.values()],
        [14, 60, 30, 30, 30, 30, 10, 12, 12, 12, 24, 24, 24])
    rows = [[d["name"], p] for d in r["decks"] for p in d["problems"]]
    rows += [[d["name"], f"slide {u['slide']} looks like a question but wasn't recognised: {u['text']}"] for d in r["decks"]
             for u in d["unknown"]]
    tab("Check these", ["deck", "what to check"], rows, [40, 110])
    wb.save(path)


if __name__ == "__main__":
    main()


# ------------------------------------------------------------------------------------------- apply
def apply_import(run_dir, category_keys, bank, writer, history=False):
    """Adds the chosen categories (and their questions, pictures, clips) to the question bank.
    Questions already in the bank are reused; anything added before from this import isn't added twice."""
    import build
    run_dir = Path(run_dir)
    r = json.loads((run_dir / "import.json").read_text())
    applied_path = run_dir / "applied.json"
    applied = json.loads(applied_path.read_text()) if applied_path.exists() else {"questions": {}, "categories": {}}
    existing = {qkey(q["question"], q["answer"]): qid for qid, q in bank["qs"].items()}
    cats = [c for c in r["categories"] if c["key"] in set(category_keys)]
    if not cats:
        raise ValueError("no categories selected")

    def place(rel):
        if not rel:
            return ""
        src = run_dir / "assets" / rel
        if not src.exists():
            return ""
        kind = "images" if rel.startswith("images/") else "media"
        dst = ROOT / "assets" / kind / "imported" / src.name
        dst.parent.mkdir(parents=True, exist_ok=True)
        if not dst.exists():
            shutil.copy2(src, dst)
        return f"{kind}/imported/{src.name}"

    need = list(dict.fromkeys(k for c in cats for k in c["questions"]))
    new_keys, rows = [], []
    for k in need:
        if k in existing or k in applied["questions"]:
            continue
        q = r["questions"][k]
        first = min(q["appearances"], key=lambda a: a["date"])
        n = len(q["appearances"])
        rows.append({"question": q["question"], "answer": q["answer"], "question_note": q.get("question_note", ""),
                     "answer_note": q.get("answer_note", ""), "question_image": place(q["question_image"]),
                     "answer_image": place(q["answer_image"]), "media": place(q["media"]),
                     "notes": f"Imported from {first['deck']} ({first['date']})" + (f"; used {n} times" if n > 1 else "")})
        new_keys.append(k)
    ids = writer.append("Questions", rows) if rows else []
    applied["questions"].update(dict(zip(new_keys, ids)))
    qid_of = {**existing, **applied["questions"]}

    sigs = {(norm(c.get("title", "")), tuple(l["question_id"] for l in bank["links"].get(cid, []))): cid
            for cid, c in bank["cats"].items()}
    new_cats, cat_rows = [], []
    for c in cats:
        qids = tuple(qid_of[k] for k in c["questions"] if k in qid_of)
        if c["key"] in applied["categories"] or (norm(c["title"]), qids) in sigs:
            applied["categories"].setdefault(c["key"], sigs.get((norm(c["title"]), qids)))
            continue
        dates = sorted(a["date"] for a in c["appearances"])
        cat_rows.append({"title": c["title"], "blurb": c["blurb"], "type": c["type"],
                         "notes": f"Imported; used {len(dates)} time{'s' if len(dates) != 1 else ''} ({dates[0]}"
                                  + (f" to {dates[-1]})" if dates[-1] != dates[0] else ")")})
        new_cats.append((c, qids))
    cat_ids = writer.append("Categories", cat_rows) if cat_rows else []
    links = []
    for (c, qids), cid in zip(new_cats, cat_ids):
        applied["categories"][c["key"]] = cid
        links += [{"category_id": cid, "position": i + 1, "question_id": qid} for i, qid in enumerate(qids)]
    if links:
        writer.append("CategoryQuestions", links)

    usage = []
    if history:
        have = {(str(build.as_date(u["date"])), u["venue_id"], u["category_id"], u.get("question_id", ""))
                for u in bank["Usage"] if u.get("date")}
        for c in cats:
            cid = applied["categories"].get(c["key"])
            if not cid:
                continue
            for a in c["appearances"]:
                if not a.get("venue") or not a.get("date"):
                    continue
                for k in c["questions"]:
                    qid = qid_of.get(k)
                    key = (a["date"], a["venue"], cid, qid)
                    if qid and key not in have:
                        have.add(key)
                        usage.append((build.as_date(a["date"]), a["venue"], "imported", cid, qid))
        if usage:
            cfg = build.load_config()
            if cfg:
                build.record_usage_sheet(usage, cfg)
            else:
                build.record_usage_xlsx(usage, ROOT / "trivia_bank.xlsx")
    applied_path.write_text(json.dumps(applied, indent=1))
    return {"questions_added": len(ids), "questions_reused": len(need) - len(ids), "categories_added": len(cat_ids),
            "categories_reused": len(cats) - len(cat_ids), "history_rows": len(usage)}
