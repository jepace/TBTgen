#!/usr/bin/env python3
"""
TBTgen app: a local web page for everything TBTgen does.

  python3 tbtgen_app.py        (or double-click TBTgen.command on a Mac)

Tabs: Nights (build decks for a week), Sets (what's in each set, problems, history), Questions (search the
bank), Video (at-home videos, with a player), Files (decks, answer keys, videos), Status (setup checks).

It runs build.py underneath, so the command line still works the same. Only this computer can reach the page
(it listens on 127.0.0.1, and every request must carry a per-launch key). Close the Terminal window, or click
Quit, to stop it.
"""
import datetime as dt
import json
import mimetypes
import os
import re
import secrets
import shutil
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from collections import defaultdict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "out"
ASSETS = ROOT / "assets"
TOKEN = secrets.token_urlsafe(16)
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

MISSING = []
for mod, pkg in (("pptx", "python-pptx"), ("openpyxl", "openpyxl"), ("PIL", "pillow"), ("lxml", "lxml")):
    try:
        __import__(mod)
    except ImportError:
        MISSING.append(pkg)
if not MISSING:
    sys.path.insert(0, str(ROOT))
    import build  # noqa: E402


# ------------------------------------------------------------------------------------------ data
DATA = {"bank": None, "source": None, "loaded_at": None, "error": None, "sheet_url": None}
DATA_LOCK = threading.Lock()


def get_bank(force=False):
    """The question bank, loaded once and kept until Reload (the Google Sheet can take a few seconds)."""
    with DATA_LOCK:
        if DATA["bank"] is not None and not force:
            return DATA["bank"]
        cfg = build.load_config()
        try:
            if cfg:
                tables = build.tables_from_sheet(cfg)
                source, url = "Google Sheet", build.SHEET_URL
            else:
                tables = build.tables_from_xlsx(ROOT / "trivia_bank.xlsx")
                source, url = "trivia_bank.xlsx", None
            DATA.update(bank=build.load_bank(tables), source=source, sheet_url=url, error=None,
                        loaded_at=time.strftime("%H:%M:%S"))
        except Exception as e:
            DATA.update(error=str(e))
            raise
        return DATA["bank"]


def usage_index(bank):
    by_q, by_c = defaultdict(list), defaultdict(list)
    for u in bank["Usage"]:
        try:
            d = str(build.as_date(u["date"]))
        except (ValueError, TypeError):
            continue
        if u.get("question_id"):
            by_q[u["question_id"]].append((d, u["venue_id"]))
        by_c[u["category_id"]].append((d, u["venue_id"]))
    return by_q, by_c


def set_mix(bank, set_id):
    """How many of a set's questions each generation is likely to know (from the Questions 'generations' tags)."""
    counts = {g: 0 for g in build.GENS}
    diff = {d: 0 for d in build.DIFFS}
    total = tagged = 0
    for sl in bank["sets"].get(set_id, []):
        for l in bank["links"].get(sl["category_id"], []):
            q = bank["qs"].get(l["question_id"])
            if not q:
                continue
            total += 1
            gs = build.parse_gens(q.get("generations"))
            if gs:
                tagged += 1
            for g in gs:
                counts[g] += 1
            d = build.parse_diff(q.get("difficulty"))
            if d in diff:
                diff[d] += 1
    return {"counts": counts, "tagged": tagged, "total": total, "diff": diff}


def audience_fit(bank, venue, set_id):
    """Generations in the venue's audience that the set serves poorly (under a quarter of tagged questions)."""
    aud = build.parse_gens((venue or {}).get("audience"))
    if not aud or not set_id or set_id not in bank["sets"]:
        return {"audience": aud, "low": []}
    m = set_mix(bank, set_id)
    if m["tagged"] < max(1, m["total"] // 2):
        return {"audience": aud, "low": [], "untagged": True}
    return {"audience": aud, "low": [g for g in aud if m["counts"][g] < 0.25 * m["tagged"]]}


def venue_name(bank, vid):
    v = bank["venues"].get(vid)
    return v["name"] if v else vid


def week_bounds(spec):
    today = dt.date.today()
    ref = {"this": today, "next": today + dt.timedelta(days=7)}.get(spec) or dt.date.fromisoformat(spec)
    start = ref - dt.timedelta(days=ref.weekday())
    return start, start + dt.timedelta(days=6)


def api_week(spec):
    bank = get_bank()
    start, end = week_bounds(spec)
    nights = []
    for r in bank["Schedule"]:
        if not r.get("date"):
            continue
        d = build.as_date(r["date"])
        if not (start <= d <= end):
            continue
        venue = bank["venues"].get(r["venue_id"])
        status = "ok"
        if not venue:
            status = "no_venue"
        elif not r.get("set_id"):
            status = "no_set"
        elif r["set_id"] not in bank["sets"]:
            status = "bad_set"
        nights.append({"date": str(d), "day": DAYS[d.weekday()], "venue_id": r["venue_id"],
                       "venue": venue["name"] if venue else (r.get("calendar_event") or r["venue_id"] or "?"),
                       "host": str(r.get("host") or ""), "set_id": str(r.get("set_id") or ""), "status": status,
                       "fit": audience_fit(bank, venue, r.get("set_id"))})
    nights.sort(key=lambda n: (n["date"], n["venue"]))
    return {"start": str(start), "end": str(end), "nights": nights, "sets": sorted(bank["sets"], reverse=True)}


def video_info(set_id):
    meta = OUT / "video" / f"{set_id} at-home.json"
    if meta.exists():
        info = json.loads(meta.read_text())
        info["file"] = f"video/{set_id} at-home.mp4"
        info["poster"] = f"video/{set_id} at-home.png"
        return info
    return None


def cat_gens(bank, cid):
    c = {g: 0 for g in build.GENS}
    for l in bank["links"].get(cid, []):
        for g in build.parse_gens(bank["qs"].get(l["question_id"], {}).get("generations")):
            c[g] += 1
    return c


def api_sets():
    bank = get_bank()
    _, by_c = usage_index(bank)
    nights = defaultdict(list)
    for r in bank["Schedule"]:
        if r.get("set_id") and r.get("date"):
            nights[r["set_id"]].append({"date": str(build.as_date(r["date"])), "venue": venue_name(bank, r["venue_id"])})
    out = []
    for set_id in sorted(bank["sets"], reverse=True):
        errs, warns = build.check_sets(bank, {set_id})
        rounds = [[], [], []]
        for sl in sorted(bank["sets"][set_id], key=lambda r: (int(r["round"] or 0), int(r["slot"] or 0))):
            cid = sl["category_id"]
            cat = bank["cats"].get(cid, {})
            hist = sorted(by_c.get(cid, []), reverse=True)
            rnd = int(sl["round"] or 1) - 1
            if 0 <= rnd < 3:
                rounds[rnd].append({
                    "category_id": cid, "title": cat.get("title", "?"), "type": cat.get("type", ""),
                    "bonus": build.is_yes(sl["bonus"]), "questions": len(bank["links"].get(cid, [])),
                    "gens": cat_gens(bank, cid),
                    "last_played": hist[0][0] if hist else "",
                    "played_at": sorted({venue_name(bank, v) for _, v in hist})[:6]})
        tbs = sorted(bank["tiebreakers"].get(set_id, []), key=lambda t: int(t.get("order") or 0))
        out.append({"set_id": set_id, "rounds": rounds, "errors": errs, "warnings": warns, "mix": set_mix(bank, set_id),
                    "tiebreakers": [{"question": str(t["question"]), "answer": str(t["answer"])} for t in tbs],
                    "nights": sorted(nights.get(set_id, []), key=lambda n: n["date"]), "video": video_info(set_id)})
    return {"sets": out}


def api_questions(q, cat, limit=300, gen="", diff=""):
    bank = get_bank()
    by_q, _ = usage_index(bank)
    cats_of = defaultdict(list)
    for cid, links in bank["links"].items():
        for l in links:
            cats_of[l["question_id"]].append(cid)
    words = [w for w in q.lower().split() if w]
    res = []
    for qid, qq in bank["qs"].items():
        if cat and cat not in cats_of.get(qid, []):
            continue
        qgens = build.parse_gens(qq.get("generations"))
        if gen == "untagged" and qgens or gen and gen != "untagged" and gen not in qgens:
            continue
        qdiff = build.parse_diff(qq.get("difficulty"))
        if diff == "unrated" and qdiff in build.DIFFS or diff and diff != "unrated" and qdiff != diff:
            continue
        hay = build.plain(f"{qid} {qq['question']} {qq['answer']} {qq.get('question_note', '')} {qq.get('answer_note', '')} "
                              f"{qq.get('notes', '')}").lower()
        if any(w not in hay for w in words):
            continue
        hist = sorted(by_q.get(qid, []), reverse=True)
        res.append({"id": qid, "question": str(qq["question"]), "answer": str(qq["answer"]),
                    "question_image": str(qq.get("question_image") or ""), "answer_image": str(qq.get("answer_image") or ""),
                    "media": str(qq.get("media") or ""), "retired": build.is_yes(qq.get("retired")),
                    "notes": str(qq.get("notes") or ""), "clip_start": str(qq.get("clip_start") or ""),
                    "question_note": str(qq.get("question_note") or ""), "answer_note": str(qq.get("answer_note") or ""),
                    "clip_end": str(qq.get("clip_end") or ""), "seconds": str(qq.get("seconds") or ""),
                    "generations": qgens, "difficulty": qdiff if qdiff in build.DIFFS else "",
                    "categories": [{"id": c, "title": bank["cats"].get(c, {}).get("title", c)} for c in cats_of.get(qid, [])],
                    "times": len(hist), "last": hist[0][0] if hist else "",
                    "history": [{"date": d, "venue": venue_name(bank, v)} for d, v in hist[:12]]})
    res.sort(key=lambda r: r["id"])
    categories = sorted(({"id": c, "title": v.get("title", c), "retired": build.is_yes(v.get("retired"))}
                         for c, v in bank["cats"].items()), key=lambda c: c["title"].lower())
    return {"total": len(bank["qs"]), "matches": len(res), "results": res[:limit], "categories": categories}


def api_categories():
    bank = get_bank()
    _, by_c = usage_index(bank)
    lookback = int(bank["settings"].get("repeat_lookback_days", 365) or 365)
    out = []
    for cid, c in bank["cats"].items():
        qs = [{"position": int(l["position"] or 0), "id": l["question_id"],
               "question": str(bank["qs"].get(l["question_id"], {}).get("question", "(missing)")),
               "answer": str(bank["qs"].get(l["question_id"], {}).get("answer", "")),
               "retired": build.is_yes(bank["qs"].get(l["question_id"], {}).get("retired")),
               "generations": build.parse_gens(bank["qs"].get(l["question_id"], {}).get("generations")),
               "difficulty": build.parse_diff(bank["qs"].get(l["question_id"], {}).get("difficulty"))}
              for l in bank["links"].get(cid, [])]
        hist = sorted(by_c.get(cid, []), reverse=True)
        in_sets = sorted(sid for sid, slots in bank["sets"].items() if any(sl["category_id"] == cid for sl in slots))
        out.append({"id": cid, "title": str(c.get("title", "")), "blurb": str(c.get("blurb", "")),
                    "type": str(c.get("type", "") or "standard"), "season_tag": str(c.get("season_tag", "")),
                    "retired": build.is_yes(c.get("retired")), "notes": str(c.get("notes", "")),
                    "questions": sorted(qs, key=lambda q: q["position"]), "last_played": hist[0][0] if hist else "",
                    "played_at": sorted({venue_name(bank, v) for _, v in hist})[:8], "times": len({d for d, _ in hist}),
                    "in_sets": in_sets, "gens": cat_gens(bank, cid)})
    out.sort(key=lambda c: c["title"].lower())
    return {"categories": out, "lookback_days": lookback, "today": str(dt.date.today())}


ASSET_FOLDERS = {"images": "images", "media": "media", "venues": "venues"}


def api_assets(kind):
    folder = ASSETS / ASSET_FOLDERS.get(kind, "images")
    if not folder.exists():
        return {"files": []}
    return {"files": sorted(p.relative_to(ASSETS).as_posix() for p in folder.rglob("*")
                            if p.is_file() and not p.name.startswith("."))}


def save_upload(kind, name, data_url):
    import base64
    folder = ASSETS / ASSET_FOLDERS.get(kind, "images")
    folder.mkdir(parents=True, exist_ok=True)
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(name).stem).strip("._") or "file"
    ext = re.sub(r"[^A-Za-z0-9]", "", Path(name).suffix.lower())[:5]
    pictures = {"jpg", "jpeg", "png", "gif", "webp"}
    allowed = {"images": pictures, "venues": pictures, "media": {"mp3", "m4a", "wav", "aac", "mp4", "mov", "m4v"}}
    if ext not in allowed[kind if kind in allowed else "images"]:
        raise ValueError(f"file type .{ext} isn't allowed here")
    raw = base64.b64decode(data_url.split(",", 1)[-1])
    target = folder / f"{stem}.{ext}"
    n = 2
    while target.exists():
        target = folder / f"{stem}-{n}.{ext}"
        n += 1
    target.write_bytes(raw)
    return target.relative_to(ASSETS).as_posix()


def api_venues():
    bank = get_bank()
    _, by_c = usage_index(bank)
    today = str(dt.date.today())
    nights = defaultdict(list)
    for r in bank["Schedule"]:
        if r.get("date") and r.get("venue_id"):
            nights[r["venue_id"]].append(str(build.as_date(r["date"])))
    played = defaultdict(set)
    for u in bank["Usage"]:
        try:
            played[u["venue_id"]].add(str(build.as_date(u["date"])))
        except (ValueError, TypeError):
            pass
    out = []
    for vid, v in bank["venues"].items():
        if not vid:
            continue
        upcoming = sorted(d for d in nights.get(vid, []) if d >= today)
        logo = str(v.get("logo") or "")
        out.append({"venue_id": vid, "name": str(v.get("name", "")), "logo": logo,
                    "logo_ok": bool(logo) and (ASSETS / logo).is_file(), "region": str(v.get("region", "")),
                    "night": str(v.get("night", "")), "upcoming_nights": str(v.get("upcoming_nights", "")),
                    "active": str(v.get("active", "")).strip().upper() not in ("N", "NO", "FALSE", "0"),
                    "calendar_name": str(v.get("calendar_name", "")), "next_night": upcoming[0] if upcoming else "",
                    "audience": build.parse_gens(v.get("audience")),
                    "scheduled": len(upcoming), "last_played": max(played.get(vid, {""})),
                    "games_played": len(played.get(vid, ()))})
    order = {d: i for i, d in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"])}
    out.sort(key=lambda v: (not v["active"], order.get(v["night"][:3].title(), 9), v["name"].lower()))
    return {"venues": out}


def api_settings():
    bank = get_bank()
    return {"settings": [{"key": str(r.get("key", "")), "value": "" if r.get("value") is None else str(r.get("value")),
                          "meaning": str(r.get("meaning", "") or "")} for r in bank["Settings"] if r.get("key")]}


IMPORTS = ROOT / "imports"


def api_import_runs():
    runs = []
    if IMPORTS.exists():
        for d in sorted(IMPORTS.iterdir(), reverse=True):
            f = d / "import.json"
            if not f.exists():
                continue
            try:
                r = json.loads(f.read_text())
            except ValueError:
                continue
            runs.append({"name": d.name, "created": r.get("created", ""), "decks": len(r.get("decks", [])),
                         "categories": len(r.get("categories", [])), "questions": len(r.get("questions", {})),
                         "folders": r.get("folders", [])})
    return {"runs": runs}


def api_import_run(name):
    import import_decks as imp
    d = IMPORTS / Path(name).name
    r = json.loads((d / "import.json").read_text())
    applied = json.loads((d / "applied.json").read_text()) if (d / "applied.json").exists() else {"questions": {}, "categories": {}}
    try:
        bank = get_bank()
        existing = {imp.qkey(q["question"], q["answer"]): qid for qid, q in bank["qs"].items()}
    except Exception:
        existing = {}
    for k, q in r["questions"].items():
        q["in_bank"] = existing.get(k) or applied["questions"].get(k) or ""
    for c in r["categories"]:
        c["added_as"] = applied["categories"].get(c["key"]) or ""
        c["new_questions"] = sum(1 for k in c["questions"] if not r["questions"][k]["in_bank"])
    r["name"] = d.name
    r["has_review"] = (d / "review.xlsx").exists()
    return r


def pick_folder():
    if sys.platform != "darwin":
        raise ValueError("folder picking needs a Mac: type or paste the folder path instead")
    p = subprocess.run(["osascript", "-e", 'POSIX path of (choose folder with prompt "Pick the folder with old decks")'],
                       capture_output=True, text=True, timeout=600)
    if p.returncode != 0:
        raise ValueError("no folder picked")
    return p.stdout.strip()


def api_files():
    files = []
    if OUT.exists():
        for p in sorted(OUT.rglob("*")):
            if p.is_file() and p.suffix.lower() in (".pptx", ".txt", ".mp4", ".key"):
                kind = {".pptx": "deck", ".key": "deck", ".txt": "answer key", ".mp4": "video"}[p.suffix.lower()]
                st = p.stat()
                files.append({"name": p.name, "rel": p.relative_to(OUT).as_posix(), "kind": kind, "size": st.st_size,
                              "modified": time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime))})
    files.sort(key=lambda f: f["modified"], reverse=True)
    return {"files": files, "folder": str(OUT)}


def api_status():
    st = {"missing": MISSING, "ffmpeg": bool(shutil.which("ffmpeg") and shutil.which("ffprobe")),
          "config": (ROOT / "tbtgen_config.json").exists(), "out": str(OUT), "python": sys.version.split()[0],
          "root": str(ROOT)}
    if not MISSING:
        if DATA["bank"] is None:
            try:
                get_bank()
            except Exception:
                pass  # the error is shown from DATA["error"]
        st.update(source=DATA["source"] or ("Google Sheet" if build.load_config() else "trivia_bank.xlsx"),
                  loaded_at=DATA["loaded_at"], sheet_url=DATA["sheet_url"], error=DATA["error"])
        try:
            import video
            st["fonts"] = video.fonts_report()
        except Exception as e:
            st["fonts"] = {"error": str(e)}
        bank = DATA["bank"]
        if bank:
            st["counts"] = {"questions": len(bank["qs"]), "categories": len(bank["cats"]), "sets": len(bank["sets"]),
                            "venues": len(bank["venues"]), "schedule rows": len(bank["Schedule"]),
                            "usage rows": len(bank["Usage"])}
            st["settings"] = {k: str(v) for k, v in bank["settings"].items() if k != "calendar_id"}
            st["assets_missing"] = sorted({qq[col] for qq in bank["qs"].values()
                                           for col in ("question_image", "answer_image", "media")
                                           if qq.get(col) and not (ASSETS / qq[col]).exists()})[:30]
    return st


# ------------------------------------------------------------------------------------------ jobs
JOB = {"id": 0, "kind": "", "title": "", "lines": [], "done": True, "rc": None, "files": [], "progress": None}
JOB_LOCK = threading.Lock()


def start_job(kind, title, args, script="build.py"):
    with JOB_LOCK:
        if not JOB["done"]:
            return None
        JOB.update(id=JOB["id"] + 1, kind=kind, title=title, lines=[], done=False, rc=None, files=[], progress=None,
                   import_dir=None)
        job_id = JOB["id"]

    def run():
        env = dict(os.environ, PYTHONUNBUFFERED="1")
        p = subprocess.Popen([sys.executable, str(ROOT / script), *args], cwd=ROOT, stdout=subprocess.PIPE,
                             stderr=subprocess.STDOUT, text=True, env=env)
        for line in p.stdout:
            line = line.rstrip("\n")
            with JOB_LOCK:
                m = re.match(r"(?:video|deck): (\d+)/(\d+)", line)
                if m:
                    JOB["progress"] = [int(m.group(1)), int(m.group(2))]
                    continue  # don't flood the log
                if line.startswith("IMPORT_DIR "):
                    JOB["import_dir"] = Path(line[len("IMPORT_DIR "):]).name
                    continue
                JOB["lines"].append(line)
                if line.startswith("built  "):
                    rel = line[len("built  "):].split("   (")[0]
                    rel = rel[len("out/"):] if rel.startswith("out/") else Path(rel).name
                    JOB["files"].append(rel)
        rc = p.wait()
        with JOB_LOCK:
            JOB.update(done=True, rc=rc)
            if rc == 0:
                DATA["bank"] = DATA["bank"] if kind != "build" or "--record" not in args else None

    threading.Thread(target=run, daemon=True).start()
    return job_id


def open_path(path):
    path = str(path)
    if sys.platform == "darwin":
        subprocess.Popen(["open", path])
    elif sys.platform.startswith("win"):
        os.startfile(path)  # noqa: S606
    else:
        subprocess.Popen(["xdg-open", path])


def reveal_path(path):
    if sys.platform == "darwin":
        subprocess.Popen(["open", "-R", str(path)])
    else:
        open_path(Path(path).parent)


def safe_under(base, rel):
    p = (base / rel).resolve()
    return p if p.is_file() and base.resolve() in p.parents else None


# ------------------------------------------------------------------------------------------ http
class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *a):
        pass

    def send_json(self, obj, code=200):
        b = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(b)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(b)

    def send_file(self, path, download=False):
        size = path.stat().st_size
        ctype = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        if path.suffix.lower() == ".txt":
            ctype = "text/plain; charset=utf-8"
        start, end, code = 0, size - 1, 200
        rng = self.headers.get("Range")
        if rng and rng.startswith("bytes=") and size:
            a, _, b = rng[6:].split(",")[0].partition("-")
            try:
                if a:
                    start, end = int(a), (int(b) if b else size - 1)
                else:
                    start, end = max(0, size - int(b)), size - 1
                end = min(end, size - 1)
                code = 206 if start <= end else 416
            except ValueError:
                pass
        if code == 416:
            self.send_response(416)
            self.send_header("Content-Range", f"bytes */{size}")
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Accept-Ranges", "bytes")
        self.send_header("Content-Length", str(end - start + 1))
        if code == 206:
            self.send_header("Content-Range", f"bytes {start}-{end}/{size}")
        if download:
            self.send_header("Content-Disposition", f'attachment; filename="{path.name}"')
        self.end_headers()
        with open(path, "rb") as f:
            f.seek(start)
            left = end - start + 1
            while left > 0:
                chunk = f.read(min(1 << 20, left))
                if not chunk:
                    break
                try:
                    self.wfile.write(chunk)
                except (BrokenPipeError, ConnectionResetError):
                    return
                left -= len(chunk)

    def host_ok(self):
        """Refuse requests whose Host isn't this app's own address. Stops a web page that points its own
        domain name at 127.0.0.1 ("DNS rebinding") from loading the page, reading the token and using the API."""
        port = self.server.server_address[1]
        if self.headers.get("Host", "").lower() in (f"127.0.0.1:{port}", f"localhost:{port}"):
            return True
        self.send_json({"error": "forbidden"}, 403)
        return False

    def do_GET(self):
        if not self.host_ok():
            return
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        if u.path == "/":
            b = PAGE.replace("__TOKEN__", TOKEN).encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(b)))
            self.end_headers()
            self.wfile.write(b)
            return
        if u.path == "/favicon.ico":
            self.send_response(204)
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        if q.get("t") != TOKEN:
            return self.send_json({"error": "forbidden"}, 403)
        try:
            if u.path == "/api/status":
                return self.send_json(api_status())
            if MISSING:
                return self.send_json({"error": "Python packages missing: " + ", ".join(MISSING)}, 500)
            if u.path == "/api/week":
                return self.send_json(api_week(q.get("week", "next")))
            if u.path == "/api/sets":
                return self.send_json(api_sets())
            if u.path == "/api/questions":
                return self.send_json(api_questions(q.get("q", ""), q.get("cat", ""), gen=q.get("gen", ""),
                                                    diff=q.get("diff", "")))
            if u.path == "/api/files":
                return self.send_json(api_files())
            if u.path == "/api/categories":
                return self.send_json(api_categories())
            if u.path == "/api/ai/status":
                import ai
                try:
                    settings = get_bank()["settings"]
                except Exception:
                    settings = {}
                return self.send_json(ai.status(settings))
            if u.path == "/api/import/runs":
                return self.send_json(api_import_runs())
            if u.path == "/api/import/run":
                return self.send_json(api_import_run(q.get("name", "")))
            if u.path == "/api/import/asset":
                p = safe_under(IMPORTS / Path(q.get("run", "")).name / "assets", q.get("name", ""))
                if not p:
                    return self.send_json({"error": "not found"}, 404)
                return self.send_file(p)
            if u.path == "/api/venues":
                return self.send_json(api_venues())
            if u.path == "/api/settings":
                return self.send_json(api_settings())
            if u.path == "/api/assets":
                return self.send_json(api_assets(q.get("kind", "images")))
            if u.path == "/api/job":
                with JOB_LOCK:
                    return self.send_json(dict(JOB))
            if u.path in ("/api/file", "/api/asset"):
                base = OUT if u.path == "/api/file" else ASSETS
                p = safe_under(base, q.get("name", ""))
                if not p:
                    return self.send_json({"error": "not found"}, 404)
                return self.send_file(p, download=q.get("download") == "1")
        except Exception as e:  # show the reason on the page
            return self.send_json({"error": str(e)}, 500)
        self.send_json({"error": "not found"}, 404)

    def do_POST(self):
        if not self.host_ok():
            return
        if self.headers.get("X-TBTgen-Token") != TOKEN:
            return self.send_json({"error": "forbidden"}, 403)
        try:
            body = json.loads(self.rfile.read(int(self.headers.get("Content-Length") or 0)) or b"{}")
        except ValueError:
            return self.send_json({"error": "bad request"}, 400)
        path = urlparse(self.path).path
        if path == "/api/reload":
            try:
                get_bank(force=True)
                return self.send_json({"ok": True, "loaded_at": DATA["loaded_at"], "source": DATA["source"]})
            except Exception as e:
                return self.send_json({"error": str(e)}, 500)
        if path == "/api/edit":
            import bank_write
            try:
                w = bank_write.writer(ROOT)
                if body.get("op") == "replace":
                    w.replace(body["tab"], body.get("match") or {}, body.get("rows") or [])
                    new_id = None
                else:
                    new_id = w.upsert(body["tab"], body.get("key") or None, body.get("values") or {})
            except (bank_write.WriteError, KeyError, ValueError) as e:
                return self.send_json({"error": str(e)}, 400)
            except Exception as e:
                return self.send_json({"error": f"couldn't save: {e}"}, 500)
            with DATA_LOCK:
                DATA["bank"] = None  # re-read on next request
            return self.send_json({"ok": True, "id": new_id})
        if path.startswith("/api/ai/"):
            import ai
            try:
                if path == "/api/ai/keys":
                    ai.save_keys({k: body.get(k) for k in ("anthropic_api_key", "gemini_api_key")})
                    return self.send_json({"ok": True})
                bank = get_bank()
                if path == "/api/ai/models":
                    return self.send_json({"models": ai.list_models(bank["settings"], body.get("provider") or None)})
                if path == "/api/ai/check":
                    qids = body.get("question_ids") or []
                    cid = body.get("category_id")
                    if cid and not qids:
                        qids = [l["question_id"] for l in bank["links"].get(cid, [])]
                    ctx = body.get("context") or (bank["cats"].get(cid, {}).get("title", "") if cid else "")
                    return self.send_json(ai.check_questions(bank, qids, ctx))
                if path == "/api/ai/generate":
                    return self.send_json(ai.generate(bank, body.get("topic", ""), body.get("count", 5),
                                                      body.get("difficulty", "mixed"), body.get("notes", ""),
                                                      body.get("type", "standard"), body.get("generations")))
                if path == "/api/ai/tag":
                    qids = body.get("question_ids") or []
                    if body.get("category_id") and not qids:
                        qids = [l["question_id"] for l in bank["links"].get(body["category_id"], [])]
                    if body.get("set_id") and not qids:
                        qids = [l["question_id"] for sl in bank["sets"].get(body["set_id"], [])
                                for l in bank["links"].get(sl["category_id"], [])]
                    return self.send_json(ai.tag_generations(bank, list(dict.fromkeys(qids))))
                if path == "/api/ai/ideas":
                    return self.send_json(ai.category_ideas(bank, body.get("theme", ""), body.get("count", 10),
                                                            body.get("generations"), body.get("notes", "")))
            except ai.AIError as e:
                return self.send_json({"error": str(e)}, 400)
            except Exception as e:
                return self.send_json({"error": f"AI request failed: {e}"}, 500)
            return self.send_json({"error": "not found"}, 404)
        if path == "/api/pick-folder":
            try:
                return self.send_json({"path": pick_folder()})
            except (ValueError, OSError, subprocess.SubprocessError) as e:
                return self.send_json({"error": str(e)}, 400)
        if path == "/api/import/scan":
            folders = [f for f in body.get("folders", []) if str(f).strip()]
            if not folders:
                return self.send_json({"error": "give a folder to scan"}, 400)
            args = []
            if body.get("pictures"):
                args.append("--pictures")
            if body.get("clips"):
                args.append("--clips")
            args += ["--", *folders]  # a folder name starting with "-" stays a folder
            job = start_job("import", f"Reading decks in {', '.join(Path(f).name or f for f in folders)}", args,
                            script="import_decks.py")
            return self.send_json({"job": job} if job else {"error": "something is already running"}, 200 if job else 409)
        if path == "/api/import/apply":
            import bank_write
            import import_decks as imp
            try:
                res = imp.apply_import(IMPORTS / Path(body.get("run", "")).name, body.get("category_keys") or [],
                                       get_bank(), bank_write.writer(ROOT), history=bool(body.get("history")))
            except (bank_write.WriteError, ValueError, FileNotFoundError) as e:
                return self.send_json({"error": str(e)}, 400)
            except Exception as e:
                return self.send_json({"error": f"import failed: {e}"}, 500)
            with DATA_LOCK:
                DATA["bank"] = None
            return self.send_json({"ok": True, **res})
        if path == "/api/import/open-review":
            p = IMPORTS / Path(body.get("run", "")).name / "review.xlsx"
            if p.exists():
                open_path(p)
            return self.send_json({"ok": p.exists()})
        if path == "/api/upload":
            try:
                rel = save_upload(body.get("kind", "images"), body.get("name", "file"), body.get("data", ""))
            except (ValueError, OSError) as e:
                return self.send_json({"error": str(e)}, 400)
            return self.send_json({"ok": True, "path": rel})
        if path == "/api/run":
            mode = body.get("mode", "build")
            if mode == "video":
                set_id = str(body.get("set_id", ""))
                if not set_id:
                    return self.send_json({"error": "pick a set"}, 400)
                try:
                    secs, ans_secs = int(body.get("seconds", 30)), float(body.get("answer_seconds", 6))
                except (TypeError, ValueError):
                    return self.send_json({"error": "seconds must be numbers"}, 400)
                args = ["--video", set_id, "--video-seconds", str(secs),
                        "--video-answer-seconds", str(ans_secs),
                        "--video-height", "1080" if str(body.get("height")) == "1080" else "720"]
                if not body.get("media", True):
                    args.append("--video-no-media")
                title = f"Video: {set_id}"
            else:
                nights = [n for n in body.get("nights", []) if ":" in n]
                if not nights:
                    return self.send_json({"error": "no nights selected"}, 400)
                args = ["--only", ",".join(nights)]
                if mode == "check":
                    args.append("--check")
                elif body.get("record"):
                    args.append("--record")
                title = ("Checking " if mode == "check" else "Building ") + f"{len(nights)} night(s)"
            job = start_job(mode, title, args)
            return self.send_json({"job": job} if job else {"error": "something is already running"},
                                  200 if job else 409)
        if path == "/api/open":
            if body.get("folder"):
                OUT.mkdir(exist_ok=True)
                open_path(OUT)
                return self.send_json({"ok": True})
            p = safe_under(OUT, body.get("file", ""))
            if not p:
                return self.send_json({"error": "file not found"}, 404)
            (reveal_path if body.get("reveal") else open_path)(p)
            return self.send_json({"ok": True})
        if path == "/api/open-sheet":
            target = DATA["sheet_url"] if DATA["sheet_url"] else str(ROOT / "trivia_bank.xlsx")
            webbrowser.open(target) if target.startswith("http") else open_path(target)
            return self.send_json({"ok": True})
        if path == "/api/quit":
            self.send_json({"ok": True})
            threading.Thread(target=self.server.shutdown, daemon=True).start()
            return
        self.send_json({"error": "not found"}, 404)


def free_port(preferred=8642):
    for port in [preferred, *range(preferred + 1, preferred + 50)]:
        with socket.socket() as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return 0


def main():
    port = free_port()
    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{srv.server_address[1]}/"
    print(f"TBTgen is running at {url}\nLeave this window open while you use it; close it (or click Quit) to stop.")
    if "--no-browser" not in sys.argv:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    print("TBTgen stopped.")


PAGE = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>TBTgen</title>
<style>
:root{--bg:#f6f5f2;--card:#fff;--ink:#1d1d1f;--muted:#6b6b70;--line:#e3e1dc;--accent:#c2185b;--accent-ink:#fff;
--ok:#1b7f4b;--warn:#a15c00;--bad:#b3261e;--chip:#efede8;--log:#16161a;--logink:#e8e6e3;--gold:#b8860b;--hover:#faf9f6}
@media (prefers-color-scheme:dark){:root{--bg:#141416;--card:#1e1e22;--ink:#ececef;--muted:#9a9aa2;--line:#33333a;
--accent:#ff4f9a;--accent-ink:#1a0710;--ok:#5fd39a;--warn:#f0b35a;--bad:#ff8a80;--chip:#2a2a30;--log:#0d0d0f;--logink:#d6d4d0;--gold:#f5c242;--hover:#24242a}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",system-ui,sans-serif}
main{max-width:1100px;margin:0 auto;padding:18px 16px 60px}
header{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:10px}
h1{font-size:22px;margin:0;letter-spacing:-.01em}h1 b{color:var(--accent)}h2{font-size:17px;margin:0 0 10px}
nav{display:flex;gap:2px;border-bottom:1px solid var(--line);margin-bottom:16px;overflow-x:auto}
nav a{padding:9px 14px;color:var(--muted);text-decoration:none;border-bottom:2px solid transparent;white-space:nowrap;font-weight:500}
nav a.on{color:var(--ink);border-bottom-color:var(--accent)}nav a:hover{color:var(--ink)}
.pill{font-size:12px;padding:3px 10px;border-radius:99px;background:var(--chip);color:var(--muted);white-space:nowrap}
.pill.busy{background:var(--accent);color:var(--accent-ink);cursor:pointer}
.spacer{flex:1}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:14px}
.row{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.label{font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin-right:4px}
button{font:inherit;border:1px solid var(--line);background:var(--card);color:var(--ink);padding:7px 14px;border-radius:8px;cursor:pointer}
button:hover{border-color:var(--muted)}button:disabled{opacity:.45;cursor:default}
button.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-ink);font-weight:600}
button.sm{padding:3px 10px;font-size:13px}button.link{border:none;background:none;color:var(--accent);padding:0;font-size:13px}
.seg{display:flex}.seg button{border-radius:0;margin-left:-1px}.seg button:first-child{border-radius:8px 0 0 8px;margin-left:0}
.seg button:last-child{border-radius:0 8px 8px 0}.seg button.on,.chip.on{background:var(--ink);color:var(--bg);border-color:var(--ink)}
.chip{padding:5px 11px;border-radius:99px;font-size:13px}
input[type=date],input[type=search],input[type=number],select{font:inherit;padding:6px 9px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink)}
input[type=search]{min-width:260px;flex:1}input[type=number]{width:80px}
table{width:100%;border-collapse:collapse;font-size:14px}th,td{text-align:left;padding:8px 6px;border-bottom:1px solid var(--line);vertical-align:top}
th{font-size:12px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);font-weight:600}
tr.dim td{color:var(--muted)}td.c{width:30px}tr.click{cursor:pointer}tr.click:hover td{background:var(--hover)}
.tag{font-size:12px;padding:2px 8px;border-radius:99px;background:var(--chip);white-space:nowrap;display:inline-block}
.tag.ok{color:var(--ok)}.tag.warn{color:var(--warn)}.tag.bad{color:var(--bad)}.tag.gold{color:var(--gold)}
.empty{color:var(--muted);padding:20px 0;text-align:center}
.muted{color:var(--muted)}.small{font-size:13px}.mono{font-family:ui-monospace,Menlo,monospace;font-size:12.5px}
pre{background:var(--log);color:var(--logink);border-radius:10px;padding:12px;margin:0;max-height:300px;overflow:auto;font:12.5px/1.5 ui-monospace,Menlo,monospace;white-space:pre-wrap}
.note{border-left:3px solid var(--warn);padding:8px 12px;background:var(--chip);border-radius:6px;margin-bottom:12px}
.note.bad{border-color:var(--bad)}.note code{font-size:12.5px}
.ln-err{color:#ff8a80}.ln-warn{color:#f0b35a}.ln-ok{color:#5fd39a}.ln-skip{color:#9a9aa2}
.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}
.rnd{border:1px solid var(--line);border-radius:10px;padding:10px}.rnd h4{margin:0 0 6px;font-size:13px;color:var(--muted);text-transform:uppercase;letter-spacing:.05em}
.cat{padding:6px 0;border-top:1px dashed var(--line)}.cat:first-of-type{border-top:none}
.bar{height:10px;background:var(--chip);border-radius:99px;overflow:hidden}.bar i{display:block;height:100%;background:var(--accent);width:0;transition:width .4s}
video{width:100%;max-height:62vh;background:#000;border-radius:10px}
.thumb{width:84px;height:56px;object-fit:cover;border-radius:6px;background:var(--chip)}
.detail{background:var(--hover)}.detail td{padding:12px}.detail img{max-height:160px;border-radius:8px;margin-right:10px}
.kv{display:grid;grid-template-columns:max-content 1fr;gap:4px 16px;font-size:14px}.kv div:nth-child(odd){color:var(--muted)}
.ok-i{color:var(--ok)}.bad-i{color:var(--bad)}
.hidden{display:none!important}
dialog{border:1px solid var(--line);border-radius:14px;background:var(--card);color:var(--ink);padding:0;width:min(760px,94vw);max-height:90vh}
dialog::backdrop{background:rgba(0,0,0,.45)}
dialog .dh{display:flex;align-items:center;gap:10px;padding:14px 18px;border-bottom:1px solid var(--line);position:sticky;top:0;background:var(--card);z-index:1}
dialog .db{padding:16px 18px;overflow:auto;max-height:calc(90vh - 130px)}dialog .df{display:flex;gap:10px;padding:12px 18px;border-top:1px solid var(--line);position:sticky;bottom:0;background:var(--card)}
.f{display:grid;grid-template-columns:150px 1fr;gap:10px 14px;align-items:start}.f>label{padding-top:7px;color:var(--muted);font-size:14px}
.f input[type=text],.f textarea,.f select{width:100%;font:inherit;padding:7px 9px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink)}
.f textarea{min-height:76px;resize:vertical}
.pickrow input[type=text],dialog input[type=text]{font:inherit;padding:7px 9px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink)}.f .hint{font-size:12px;color:var(--muted);margin-top:3px}
.pickrow{display:flex;gap:6px;align-items:center;flex-wrap:wrap}.pickrow input[type=text]{flex:1;min-width:160px}.pickrow img{height:46px;border-radius:6px}
.slot{display:flex;gap:8px;align-items:flex-start;padding:7px 0;border-bottom:1px dashed var(--line)}.slot .n{width:22px;color:var(--muted);font-weight:600;padding-top:2px}
.slot .t{flex:1}.slot button{padding:2px 8px;font-size:13px}
.res{max-height:220px;overflow:auto;border:1px solid var(--line);border-radius:8px;margin-top:6px}.res div{padding:6px 10px;border-bottom:1px solid var(--line);display:flex;gap:8px;align-items:center}
.res div:last-child{border-bottom:none}.res .t{flex:1;font-size:14px}
.sgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}.sgrid .rnd select{width:100%;margin-bottom:4px}.sgrid label.b{font-size:12px;color:var(--muted);display:flex;gap:4px;align-items:center;margin-bottom:8px}
.warnlist div{font-size:13px;margin:2px 0}
#toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:var(--ink);color:var(--bg);padding:10px 18px;border-radius:10px;font-size:14px;z-index:99;opacity:0;transition:opacity .25s;pointer-events:none}
#toast.show{opacity:1}#toast.err{background:var(--bad);color:#fff}
select.setpick{font-size:13px;padding:3px 6px}
.airesult{border:1px solid var(--line);border-radius:10px;padding:12px;margin-bottom:10px}
.airesult .q{font-weight:600}.airesult ul{margin:6px 0 6px 18px;padding:0}.airesult li{font-size:14px}
.sugg{background:var(--hover);border-radius:8px;padding:8px 10px;margin:6px 0;font-size:14px}
.ai-busy{display:flex;gap:10px;align-items:center;padding:18px 0;color:var(--muted)}
.spin{width:18px;height:18px;border:2px solid var(--line);border-top-color:var(--accent);border-radius:50%;animation:sp 0.8s linear infinite}
@keyframes sp{to{transform:rotate(360deg)}}
.gen{display:flex;gap:10px;align-items:flex-start;border-bottom:1px dashed var(--line);padding:9px 0}.gen input{margin-top:4px}
.gchips{display:inline-flex;gap:4px;flex-wrap:wrap}.gchip{font-size:11px;padding:1px 7px;border-radius:99px;background:var(--chip);color:var(--muted);white-space:nowrap}
.gpick{display:flex;gap:6px;flex-wrap:wrap}.gpick label{display:flex;gap:4px;align-items:center;font-size:13px;border:1px solid var(--line);border-radius:99px;padding:3px 10px;cursor:pointer}
.gpick label:has(input:checked){background:var(--ink);color:var(--bg);border-color:var(--ink)}.gpick input{display:none}
.mix{display:grid;grid-template-columns:max-content 1fr max-content;gap:3px 10px;align-items:center;font-size:12.5px;max-width:420px}
.mix .b{height:8px;background:var(--chip);border-radius:99px;overflow:hidden}.mix .b i{display:block;height:100%;background:var(--accent)}
.mix .lo{color:var(--warn)}
.dtag{font-size:11px;padding:1px 7px;border-radius:99px;white-space:nowrap;border:1px solid var(--line)}
.dtag.easy{color:var(--ok)}.dtag.medium{color:var(--warn)}.dtag.hard{color:var(--bad)}
@media (max-width:760px){.f{grid-template-columns:1fr}.f>label{padding-top:0}.sgrid{grid-template-columns:1fr}}
@media (max-width:760px){.grid3{grid-template-columns:1fr}.hide-sm{display:none}input[type=search]{min-width:0}}
.mk-red{color:#d32f2f}.mk-gold{color:#b8860b}
.slideprev{background:#000;color:#fff;border-radius:8px;padding:12px 14px;font-family:"Trebuchet MS",sans-serif;line-height:1.3}
.slideprev .mk-red{color:#e53935}.slideprev .mk-gold{color:#f5c242}.slideprev .pn .mi{font-style:normal}.slideprev .pq{font-weight:bold;white-space:pre-wrap}
.slideprev .pn{font-style:italic;font-size:.88em;margin-top:2px;white-space:pre-wrap}.slideprev .pa{color:#f5c242;font-weight:bold;font-size:1.2em;margin-top:10px}
</style></head><body><main>
<header><h1><b>TBT</b>gen</h1><span class="pill" id="source">…</span><span class="pill busy hidden" id="busy" title="Show progress"></span>
<span class="spacer"></span><button id="reloadData" title="Read the spreadsheet again">Reload data</button>
<button id="openSheet" class="hidden">Open spreadsheet</button><button id="quit" title="Stop TBTgen">Quit</button></header>
<nav id="tabs"><a data-tab="nights">Nights</a><a data-tab="sets">Sets</a><a data-tab="categories">Categories</a><a data-tab="questions">Questions</a><a data-tab="venues">Venues</a>
<a data-tab="video">Video</a><a data-tab="files">Files</a><a data-tab="import">Import</a><a data-tab="settings">Settings</a><a data-tab="status">Status</a></nav>
<div id="notes"></div>

<!-- ============================== NIGHTS -->
<section data-pane="nights">
<div class="card">
  <div class="row" style="margin-bottom:12px"><span class="label">Week</span>
    <div class="seg" id="weekSeg"><button data-week="this">This week</button><button data-week="next" class="on">Next week</button></div>
    <input type="date" id="pick" title="Any date in the week"></div>
  <div class="row"><span class="label">Days</span><span id="days" class="row" style="gap:6px"></span></div>
</div>
<div class="card">
  <div class="row" style="margin-bottom:8px"><strong id="range">Loading…</strong><span class="spacer"></span>
    <button id="all" class="sm">Select all</button><button id="none" class="sm">Select none</button></div>
  <table><thead><tr><th class="c"></th><th>Day</th><th class="hide-sm">Date</th><th>Venue</th><th>Host</th><th>Set</th></tr></thead>
  <tbody id="rows"></tbody></table>
  <div class="row" style="margin-top:14px">
    <button id="check">Check</button><button id="build" class="primary">Build</button>
    <label class="row" style="gap:6px" title="Writes what was played to the Usage tab, so repeat checks know about it">
      <input type="checkbox" id="record"> Log as played (Usage)</label>
    <span class="spacer"></span><button id="openout">Open decks folder</button></div>
</div>
</section>

<!-- ============================== SETS -->
<section data-pane="sets" class="hidden">
<div class="card"><div class="row"><span class="muted small">Each set is 3 rounds × 3 categories, with one bonus category.</span><span class="spacer"></span>
  <button class="primary" id="newSet">New set</button></div></div>
<div id="setsList"><div class="empty">Loading…</div></div></section>

<!-- ============================== CATEGORIES -->
<section data-pane="categories" class="hidden">
<div class="card"><div class="row"><input type="search" id="cSearch" placeholder="Search categories…">
  <label class="row small" style="gap:5px"><input type="checkbox" id="cRetired"> show retired</label>
  <span class="muted small" id="cCount"></span><span class="spacer"></span><button id="ideasBtn">Category ideas with AI…</button><button class="primary" id="newCat">New category</button></div></div>
<div class="card" style="padding:6px 16px"><table><thead><tr><th>Category</th><th class="hide-sm">Type</th><th>Questions</th>
  <th class="hide-sm">Last played</th><th class="hide-sm">In sets</th></tr></thead><tbody id="cRows"></tbody></table></div>
</section>

<!-- ============================== QUESTIONS -->
<section data-pane="questions" class="hidden">
<div class="card"><div class="row"><input type="search" id="qSearch" placeholder="Search questions, answers, IDs…">
  <select id="qCat"><option value="">All categories</option></select>
  <select id="qGen"><option value="">All ages</option><option>Boomer</option><option>Gen X</option><option>Millennial</option><option>Gen Z</option><option>Gen Alpha</option><option value="untagged">Not tagged yet</option></select>
  <select id="qDiff"><option value="">Any difficulty</option><option>easy</option><option>medium</option><option>hard</option><option value="unrated">Not rated yet</option></select>
  <span class="muted small" id="qCount"></span>
  <span class="spacer"></span><button class="primary" id="newQ">New question</button></div>
  <p class="muted small" style="margin:10px 0 0">Click a question for details and Edit. Changes are saved straight to the spreadsheet.</p></div>
<div class="card" style="padding:6px 16px"><table><thead><tr><th class="hide-sm"></th><th>ID</th><th>Question</th><th>Answer</th>
  <th class="hide-sm">Categories</th><th class="hide-sm">Asked</th></tr></thead><tbody id="qRows"></tbody></table></div>
</section>

<!-- ============================== VENUES -->
<section data-pane="venues" class="hidden">
<div class="card"><div class="row"><input type="search" id="vnSearch" placeholder="Search venues…">
  <label class="row small" style="gap:5px"><input type="checkbox" id="vnInactive"> show inactive</label>
  <span class="muted small" id="vnCount"></span><span class="spacer"></span><button class="primary" id="newVenue">New venue</button></div>
  <p class="muted small" style="margin:10px 0 0">The calendar name must match how the venue is written in calendar events (“<i>Calendar name</i> - Host”).
  Venues without a logo use the Trivia By Troy logo.</p></div>
<div class="card" style="padding:6px 16px"><table><thead><tr><th></th><th>Venue</th><th>Night</th><th class="hide-sm">Calendar name</th>
  <th class="hide-sm">Region</th><th class="hide-sm">Next night</th></tr></thead><tbody id="vnRows"></tbody></table></div>
</section>

<!-- ============================== IMPORT -->
<section data-pane="import" class="hidden">
<div class="card"><h2>Import old decks</h2>
  <p class="muted small" style="margin-top:0">Reads Keynote (.key) and PowerPoint (.pptx) game decks in a folder and its sub-folders, finds the
  categories, questions and answers, and merges repeats. Nothing is added to the bank until you pick categories and click Add.</p>
  <div class="pickrow" style="margin-bottom:10px"><input type="text" id="imFolder" placeholder="/Users/troy/Documents/Trivia decks">
    <button id="imPick">Choose folder…</button></div>
  <div class="row" style="margin-bottom:6px"><label class="row small" style="gap:6px"><input type="checkbox" id="imPics" checked> Import pictures</label>
    <label class="row small" style="gap:6px"><input type="checkbox" id="imClips"> Import song &amp; movie clips</label></div>
  <div class="hint muted small" style="margin-bottom:12px">Pictures and clips are optional: unticked, they're only counted, so you can see how much space they'd take first.
    Clips (especially movies) are the big ones.</div>
  <div class="row"><button class="primary" id="imScan">Scan decks</button><span class="muted small" id="imRunsHint"></span></div></div>
<div class="card hidden" id="imProgress"><div class="row" style="margin-bottom:8px"><strong id="imPTitle">Reading decks…</strong><span class="spacer"></span>
  <span class="muted small" id="imPct"></span></div><div class="bar"><i id="imBar"></i></div></div>
<div class="card hidden" id="imSummary"></div>
<div class="card hidden" id="imCats"></div>
<div class="card"><h2>Earlier scans</h2><div id="imRuns" class="muted small">None yet.</div></div>
</section>

<!-- ============================== SETTINGS -->
<section data-pane="settings" class="hidden">
<div class="card"><h2>Settings</h2><p class="muted small" style="margin-top:0">These live on the spreadsheet's Settings tab and apply to every deck and video.</p>
  <div class="f" id="stForm"><div class="empty">Loading…</div></div>
  <div class="row" style="margin-top:14px"><button class="primary" id="stSave" disabled>Save changes</button><span class="muted small" id="stDirty"></span></div></div>
<div class="card" id="aiCard"><h2>AI assistant</h2>
  <p class="muted small" style="margin-top:0">Used only when you click <b>Check with AI</b> or <b>Generate with AI</b>. It suggests; you decide what to keep.
  Keys are stored only on this computer (tbtgen_config.json), never in the spreadsheet or GitHub.</p>
  <div class="f">
    <label>Provider</label><div><select id="aiProvider"><option value="anthropic">Claude (Anthropic)</option><option value="gemini">Gemini (Google)</option></select></div>
    <label>Model</label><div><div class="pickrow"><input type="text" id="aiModel" list="aiModelList" placeholder="default"><datalist id="aiModelList"></datalist>
      <button class="sm" id="aiList">List models</button></div><div class="hint" id="aiModelHint"></div></div>
    <label>Effort</label><div><select id="aiEffort"><option value="low">low (cheapest)</option><option value="medium">medium</option><option value="high">high (most careful)</option></select>
      <div class="hint">How hard the AI thinks. Higher costs more. Applies to Claude.</div></div>
    <label>Fact-check online</label><div><label class="row small" style="gap:6px"><input type="checkbox" id="aiWeb"> let the AI search the web to verify facts (recommended; small extra cost per search)</label></div>
    <label>Claude API key</label><div><input type="password" id="aiKeyA" autocomplete="off" style="width:100%;font:inherit;padding:7px 9px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink)">
      <div class="hint">From console.anthropic.com → API keys. Leave blank to keep the saved one.</div></div>
    <label>Gemini API key</label><div><input type="password" id="aiKeyG" autocomplete="off" style="width:100%;font:inherit;padding:7px 9px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink)">
      <div class="hint">From aistudio.google.com → Get API key. Leave blank to keep the saved one.</div></div>
  </div>
  <div class="row" style="margin-top:14px"><button class="primary" id="aiSave">Save AI settings</button><span class="small muted" id="aiUsage"></span></div></div>
<div class="card"><h2>Add a setting</h2><div class="pickrow"><input type="text" id="stNewKey" placeholder="key (e.g. video_media_extra_seconds)">
  <input type="text" id="stNewVal" placeholder="value" style="max-width:160px"><input type="text" id="stNewMeaning" placeholder="what it's for (optional)">
  <button id="stAdd">Add</button></div><div class="hint muted small" style="margin-top:6px">Known extras: <span class="mono">video_media_extra_seconds</span> (seconds after a clip in videos, default 10).</div></div>
</section>

<!-- ============================== VIDEO -->
<section data-pane="video" class="hidden">
<div class="card"><h2>At-home video</h2>
  <p class="muted small" style="margin-top:0">One video per set: title, how to play, three rounds with a countdown on every question,
  answers after each round, and a thanks-for-playing card. Music and movie clips play on their questions.</p>
  <div class="row" style="margin-bottom:10px"><span class="label">Set</span><select id="vSet"></select>
    <span class="label" style="margin-left:10px">Size</span><div class="seg" id="vSize"><button data-h="720" class="on">720p</button><button data-h="1080">1080p</button></div></div>
  <div class="row" style="margin-bottom:10px">
    <label class="row" style="gap:6px">Seconds per question <input type="number" id="vSec" value="30" min="5" max="120"></label>
    <label class="row" style="gap:6px">Seconds per answer <input type="number" id="vAns" value="6" min="2" max="30" step="0.5"></label></div>
  <div class="row" style="margin-bottom:12px"><label class="row" style="gap:6px"><input type="checkbox" id="vMedia" checked>
    Include music &amp; movie categories</label>
    <span class="muted small">Song and movie clips may be copyrighted: fine to watch at home, risky to post publicly.</span></div>
  <div class="row"><button class="primary" id="vBuild">Make video</button><span class="muted small" id="vEta"></span></div>
</div>
<div class="card hidden" id="vProgress"><div class="row" style="margin-bottom:8px"><strong id="vTitle">Making video…</strong>
  <span class="spacer"></span><span class="muted small" id="vPct"></span></div><div class="bar"><i id="vBar"></i></div></div>
<div class="card hidden" id="vPlayerCard"><div class="row" style="margin-bottom:8px"><strong id="vPlayerTitle"></strong><span class="spacer"></span>
  <button class="sm" id="vDownload">Download</button><button class="sm" id="vOpen">Open in player</button><button class="sm" id="vReveal">Show in folder</button></div>
  <video id="vPlayer" controls preload="metadata"></video><div class="muted small" id="vMeta" style="margin-top:6px"></div></div>
<div class="card"><h2>Made so far</h2><div id="vList" class="muted">None yet.</div></div>
</section>

<!-- ============================== FILES -->
<section data-pane="files" class="hidden">
<div class="card"><div class="row"><div class="seg" id="fKind"><button data-k="" class="on">All</button><button data-k="deck">Decks</button>
  <button data-k="answer key">Answer keys</button><button data-k="video">Videos</button></div><input type="search" id="fSearch" placeholder="Filter by name…">
  <span class="spacer"></span><button id="fOpenFolder">Open folder</button></div></div>
<div class="card" style="padding:6px 16px"><table><thead><tr><th>Name</th><th class="hide-sm">Kind</th><th class="hide-sm">Size</th><th class="hide-sm">Made</th><th></th></tr></thead>
  <tbody id="fRows"></tbody></table></div>
<div class="card hidden" id="fView"><div class="row" style="margin-bottom:8px"><strong id="fViewTitle"></strong><span class="spacer"></span>
  <button class="sm" id="fViewClose">Close</button></div><pre id="fViewBody" style="max-height:480px"></pre></div>
</section>

<!-- ============================== STATUS -->
<section data-pane="status" class="hidden"><div id="statusBody"><div class="empty">Loading…</div></div></section>

<!-- ============================== DIALOGS -->
<dialog id="dlg"><div class="dh"><strong id="dlgTitle"></strong><span class="spacer"></span><button class="sm" data-close>✕</button></div>
  <div class="db" id="dlgBody"></div><div class="df" id="dlgFoot"></div></dialog>
<dialog id="dlg2"><div class="dh"><strong id="dlg2Title"></strong><span class="spacer"></span><button class="sm" data-close>✕</button></div>
  <div class="db" id="dlg2Body"></div><div class="df" id="dlg2Foot"></div></dialog>
<div id="toast"></div>

<!-- ============================== RUN LOG (shared) -->
<section class="card hidden" id="runcard">
  <div class="row" style="margin-bottom:8px"><strong id="runtitle">Running…</strong><span class="spacer"></span><button class="sm" id="runHide">Hide</button></div>
  <pre id="log"></pre><div id="runfiles" class="row" style="margin-top:10px"></div>
</section>
</main>
<script>
const T="__TOKEN__",$=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)];
const DAYS=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"];
const S={week:"next",nights:[],dayOn:new Set(DAYS),polling:null,files:[],fKind:"",vHeight:"720",sets:[],qTimer:null,loaded:{}};
const store={get(k,d){try{const v=localStorage.getItem("tbt_"+k);return v===null?d:JSON.parse(v)}catch(e){return d}},set(k,v){try{localStorage.setItem("tbt_"+k,JSON.stringify(v))}catch(e){}}};
const get=p=>fetch(p+(p.includes("?")?"&":"?")+"t="+T).then(r=>r.json());
const post=(p,b)=>fetch(p,{method:"POST",headers:{"Content-Type":"application/json","X-TBTgen-Token":T},body:JSON.stringify(b||{})}).then(r=>r.json());
// slide-text markup, same rules as build.markup_runs: *italic* **bold** [red]..[/red] [gold]..[/gold]
function md(s){let t=esc(s).replace(/\\\*/g,"\x07").replace(/\\\[/g,"\x08"),p;
  do{p=t;t=t.replace(/(?<![\w*])\*\*\*(?=[^\s*])([\s\S]+?)(?<=[^\s*])\*\*\*(?![\w*])/g,"<b><i>$1</i></b>")
    .replace(/(?<![\w*])\*\*(?=[^\s*])([\s\S]+?)(?<=[^\s*])\*\*(?![\w*])/g,"<b>$1</b>")
    .replace(/(?<![\w*])\*(?=[^\s*])([\s\S]+?)(?<=[^\s*])\*(?![\w*])/g,'<i class="mi">$1</i>')
    .replace(/\[(red|gold)\]([\s\S]*?)\[\/(?:\1)?\]/gi,(m,c,x)=>`<span class="mk-${c.toLowerCase()}">${x}</span>`)}while(p!==t);
  return t.replace(/\x07/g,"*").replace(/\x08/g,"[")}
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fileUrl=(rel,dl)=>`/api/file?name=${encodeURIComponent(rel)}&t=${T}${dl?"&download=1":""}`;
const assetUrl=rel=>`/api/asset?name=${encodeURIComponent(rel)}&t=${T}`;
const fmtDate=s=>s?new Date(s+"T12:00:00").toLocaleDateString(undefined,{month:"short",day:"numeric"}):"";
const fmtLong=s=>s?new Date(s+"T12:00:00").toLocaleDateString(undefined,{weekday:"short",month:"short",day:"numeric",year:"numeric"}):"";
const fmtSize=b=>b>1e9?(b/1e9).toFixed(1)+" GB":b>1e6?(b/1e6).toFixed(1)+" MB":Math.max(1,Math.round(b/1e3))+" KB";
const fmtLen=s=>`${Math.floor(s/60)}:${String(Math.round(s%60)).padStart(2,"0")}`;

/* ---------- tabs ---------- */
function showTab(t){if(!$(`[data-pane="${t}"]`))t="nights";store.set("tab",t);
  $$("nav a").forEach(a=>a.classList.toggle("on",a.dataset.tab===t));$$("[data-pane]").forEach(p=>p.classList.toggle("hidden",p.dataset.pane!==t));
  if(location.hash!=="#"+t)history.replaceState(null,"","#"+t);
  ({nights:()=>S.loaded.nights||loadWeek(),sets:loadSets,categories:loadCats,questions:()=>S.loaded.q||loadQuestions(),venues:loadVenues,video:loadVideoTab,files:loadFiles,import:loadImportTab,settings:loadSettings,status:loadStatus}[t]||(()=>{}))()}
$$("nav a").forEach(a=>a.onclick=()=>showTab(a.dataset.tab));
window.addEventListener("hashchange",()=>{const t=location.hash.slice(1);if(t&&t!==store.get("tab",""))showTab(t)});

/* ---------- nights ---------- */
function chips(){$("#days").innerHTML=DAYS.map(d=>`<button class="chip${S.dayOn.has(d)?" on":""}" data-d="${d}">${d}</button>`).join("");
  $$("#days .chip").forEach(b=>b.onclick=()=>{const d=b.dataset.d;S.dayOn.has(d)?S.dayOn.delete(d):S.dayOn.add(d);store.set("days",[...S.dayOn]);chips();renderNights()})}
function renderNights(){const shown=S.nights.filter(n=>S.dayOn.has(n.day));
  if(S.rangeTxt)$("#range").textContent=S.rangeTxt+(shown.length<S.nights.length?` · ${shown.length} shown`:"");
  $("#rows").innerHTML=shown.length?shown.map(n=>{const ok=n.status==="ok";
    const tag=n.status==="no_venue"?`<span class="tag bad">venue not on Venues tab</span>`:
      `<select class="setpick" data-date="${n.date}" data-venue="${esc(n.venue_id)}" title="Pick the set for this night (saved to the spreadsheet)">
        <option value="">— no set yet —</option>${(S.setIds||[]).map(id=>`<option${id===n.set_id?" selected":""}>${esc(id)}</option>`).join("")}
        ${n.status==="bad_set"?`<option selected>${esc(n.set_id)} (not on Sets tab)</option>`:""}</select>`;
    return `<tr class="${ok?"":"dim"}"><td class="c"><input type="checkbox" data-k="${n.date}:${esc(n.venue_id)}" ${ok&&n.sel!==false?"checked":""} ${ok?"":"disabled"}></td>
      <td>${n.day}</td><td class="hide-sm">${fmtDate(n.date)}</td><td>${esc(n.venue)}${n.fit&&n.fit.audience.length?`<div>${gchips(n.fit.audience)}</div>`:""}</td><td>${esc(n.host)}</td>
      <td>${tag}${n.fit&&n.fit.low&&n.fit.low.length?`<div class="small" style="color:var(--warn);margin-top:3px" title="Under a quarter of this set's tagged questions suit them">⚠ light on ${esc(n.fit.low.join(", "))}</div>`:""}</td></tr>`}).join("")
    :`<tr><td colspan="6" class="empty">No nights on the schedule for these days.</td></tr>`;
  $$("#rows input").forEach(c=>c.onchange=()=>{const n=S.nights.find(x=>`${x.date}:${x.venue_id}`===c.dataset.k);if(n)n.sel=c.checked;countSel()});
  $$("#rows select.setpick").forEach(sel=>sel.onchange=async()=>{sel.disabled=true;
    const r=await edit({op:"upsert",tab:"Schedule",key:{date:sel.dataset.date,venue_id:sel.dataset.venue},values:{set_id:sel.value}});
    if(r)toast(sel.value?`Set ${sel.value} saved for ${sel.dataset.venue}`:"Set cleared");S.loaded={};loadWeek()});countSel()}
const selected=()=>$$("#rows input:checked").map(c=>c.dataset.k);
function countSel(){const k=selected().length;$("#build").textContent=k?`Build ${k} deck${k>1?"s":""}`:"Build";$("#build").disabled=$("#check").disabled=!k||!!S.polling}
async function loadWeek(){S.loaded.nights=true;$("#range").textContent="Loading schedule…";$("#rows").innerHTML="";
  const r=await get("/api/week?week="+encodeURIComponent(S.week)).catch(e=>({error:String(e)}));
  if(r.error){$("#range").textContent="Couldn't load the schedule";$("#rows").innerHTML=`<tr><td colspan="6" class="empty">${esc(r.error)}</td></tr>`;return}
  S.nights=r.nights;S.setIds=r.sets;S.rangeTxt=`${fmtDate(r.start)} – ${fmtDate(r.end)} · ${S.nights.length} night${S.nights.length==1?"":"s"}`;renderNights();refreshSource()}
$$("#weekSeg button").forEach(b=>b.onclick=()=>{$$("#weekSeg button").forEach(x=>x.classList.remove("on"));b.classList.add("on");$("#pick").value="";S.week=b.dataset.week;loadWeek()});
$("#pick").onchange=e=>{if(!e.target.value)return;$$("#weekSeg button").forEach(x=>x.classList.remove("on"));S.week=e.target.value;loadWeek()};
$("#all").onclick=()=>{S.nights.forEach(n=>n.sel=true);renderNights()};$("#none").onclick=()=>{S.nights.forEach(n=>n.sel=false);renderNights()};
$("#check").onclick=()=>run({mode:"check",nights:selected()});
$("#build").onclick=()=>run({mode:"build",nights:selected(),record:$("#record").checked});
$("#openout").onclick=()=>post("/api/open",{folder:true});

/* ---------- sets ---------- */
async function loadSets(){const r=await get("/api/sets").catch(e=>({error:String(e)}));
  if(r.error){$("#setsList").innerHTML=`<div class="card note bad">${esc(r.error)}</div>`;return}S.sets=r.sets;
  $("#setsList").innerHTML=r.sets.length?r.sets.map(s=>{
    const probs=[...s.errors.map(e=>`<div class="small" style="color:var(--bad)">✗ ${esc(e)}</div>`),...s.warnings.map(w=>`<div class="small" style="color:var(--warn)">! ${esc(w)}</div>`)].join("");
    const rounds=s.rounds.map((cats,i)=>`<div class="rnd"><h4>Round ${i+1}</h4>${cats.map(c=>`<div class="cat"><strong>${esc(c.title)}</strong>
      ${c.bonus?'<span class="tag gold">bonus ×2</span>':""} ${c.type!=="standard"?`<span class="tag">${esc(c.type)}</span>`:""}
      ${c.questions!==5?`<span class="tag bad">${c.questions} questions</span>`:""}
      <div class="muted small">${esc(c.category_id)} · ${c.last_played?`last played ${fmtDate(c.last_played)} at ${esc(c.played_at.join(", "))}`:"never played"}</div></div>`).join("")||'<div class="muted small">empty</div>'}</div>`).join("");
    const nights=s.nights.length?s.nights.map(n=>`${fmtDate(n.date)} ${esc(n.venue)}`).join(" · "):"not assigned to any night";
    return `<div class="card"><div class="row" style="margin-bottom:10px"><h2 style="margin:0">${esc(s.set_id)}</h2>
      ${s.errors.length?'<span class="tag bad">problems</span>':s.warnings.length?'<span class="tag warn">warnings</span>':'<span class="tag ok">ready</span>'}
      ${s.video?`<span class="tag">video ${fmtLen(s.video.length_seconds)}</span>`:""}<span class="spacer"></span>
      <button class="sm" data-editset="${esc(s.set_id)}">Edit set</button><button class="sm" data-aiset="${esc(s.set_id)}">Check with AI</button><button class="sm" data-tagset="${esc(s.set_id)}">Tag with AI</button>
      <button class="sm" data-video="${esc(s.set_id)}">${s.video?"Watch / remake video":"Make video"}</button></div>
      <div class="grid3">${rounds}</div><div class="small muted" style="margin-top:10px">Nights: ${nights}</div>
      <details style="margin-top:8px"><summary class="small">Age &amp; difficulty mix</summary><div style="margin-top:6px">${mixHtml(s.mix)}</div></details>
      ${probs?`<div style="margin-top:8px">${probs}</div>`:""}</div>`}).join(""):`<div class="card empty">No sets on the Sets tab yet.</div>`;
  $$("[data-video]").forEach(b=>b.onclick=()=>{store.set("vSet",b.dataset.video);showTab("video")});
  $$("[data-editset]").forEach(b=>b.onclick=()=>editSet(S.sets.find(x=>x.set_id===b.dataset.editset)));
  $$("[data-aiset]").forEach(b=>b.onclick=()=>aiCheckSet(S.sets.find(x=>x.set_id===b.dataset.aiset)));
  $$("[data-tagset]").forEach(b=>b.onclick=()=>aiTag({set_id:b.dataset.tagset},`Tag: set ${b.dataset.tagset}`,1,()=>loadSets()))}

/* ---------- questions ---------- */
async function loadQuestions(){S.loaded.q=true;const q=$("#qSearch").value,cat=$("#qCat").value,gen=$("#qGen").value,diff=$("#qDiff").value;
  const r=await get(`/api/questions?q=${encodeURIComponent(q)}&cat=${encodeURIComponent(cat)}&gen=${encodeURIComponent(gen)}&diff=${encodeURIComponent(diff)}`).catch(e=>({error:String(e)}));
  if(r.error){$("#qRows").innerHTML=`<tr><td colspan="6" class="empty">${esc(r.error)}</td></tr>`;return}
  if($("#qCat").options.length<=1)$("#qCat").innerHTML='<option value="">All categories</option>'+r.categories.map(c=>`<option value="${esc(c.id)}">${esc(c.title)}${c.retired?" (retired)":""}</option>`).join("");
  $("#qCount").textContent=r.matches===r.total?`${r.total} questions`:`${r.matches} of ${r.total}`+(r.results.length<r.matches?` (first ${r.results.length} shown)`:"");
  $("#qRows").innerHTML=r.results.length?r.results.map((x,i)=>{const img=x.question_image||x.answer_image;
    return `<tr class="click${x.retired?" dim":""}" data-i="${i}"><td class="hide-sm">${img?`<img class="thumb" loading="lazy" src="${assetUrl(img)}" alt="">`:x.media?'<span class="tag">'+(/\.(mp3|m4a|wav|aac)$/i.test(x.media)?"♪ audio":"▶ video")+'</span>':""}</td>
      <td class="mono">${esc(x.id)}</td><td>${md(x.question)}${x.retired?' <span class="tag">retired</span>':""}<div class="row" style="gap:4px">${dtag(x.difficulty)}${gchips(x.generations)}</div></td><td><strong>${md(x.answer)}</strong></td>
      <td class="hide-sm small">${x.categories.map(c=>esc(c.title)).join("<br>")||'<span class="muted">none</span>'}</td>
      <td class="hide-sm small" style="white-space:nowrap">${x.times?`${x.times}× · ${fmtDate(x.last)}`:'<span class="muted">never</span>'}</td></tr>`}).join("")
    :`<tr><td colspan="6" class="empty">No questions match.</td></tr>`;
  $$("#qRows tr.click").forEach(tr=>tr.onclick=()=>{const nx=tr.nextElementSibling;if(nx&&nx.classList.contains("detail")){nx.remove();return}
    const x=r.results[+tr.dataset.i];const d=document.createElement("tr");d.className="detail";
    d.innerHTML=`<td colspan="6"><div class="row" style="align-items:flex-start">
      ${x.question_image?`<div><div class="small muted">Question picture</div><img src="${assetUrl(x.question_image)}" alt=""></div>`:""}
      ${x.answer_image?`<div><div class="small muted">Answer picture</div><img src="${assetUrl(x.answer_image)}" alt=""></div>`:""}
      ${x.media?`<div><div class="small muted">Clip: ${esc(x.media)}</div>${/\.(mp3|m4a|wav|aac)$/i.test(x.media)?`<audio controls preload="none" src="${assetUrl(x.media)}"></audio>`:`<video controls preload="none" style="max-width:320px" src="${assetUrl(x.media)}"></video>`}</div>`:""}
      <div style="flex:1;min-width:220px"><div style="white-space:pre-wrap">${md(x.question)}</div>${x.question_note?`<div class="small"><i>${md(x.question_note)}</i></div>`:""}
      <div style="margin-top:6px">Answer: <strong>${md(x.answer)}</strong>${x.answer_note?` <span class="small"><i>${md(x.answer_note)}</i></span>`:""}</div>
      ${x.notes?`<div class="small muted">Notes: ${esc(x.notes)}</div>`:""}
      <div class="small" style="margin-top:6px">In: ${x.categories.map(c=>`${esc(c.title)} <span class="muted mono">${esc(c.id)}</span>`).join(", ")||"no category"}</div>
      <div class="small muted" style="margin-top:4px">${x.history.length?"Asked: "+x.history.map(h=>`${fmtDate(h.date)} ${esc(h.venue)}`).join(" · "):"Never asked yet"}</div>
      <div style="margin-top:10px" class="row"><button class="sm primary" data-editq>Edit question</button><button class="sm" data-aiq>Check with AI</button></div></div></div></td>`;
    d.querySelector("[data-editq]").onclick=()=>editQuestion(x,()=>loadQuestions());
    d.querySelector("[data-aiq]").onclick=()=>aiCheck({question_ids:[x.id]},`Check: ${x.id}`,1,()=>loadQuestions());
    tr.after(d)})}
$("#qSearch").oninput=()=>{clearTimeout(S.qTimer);S.qTimer=setTimeout(loadQuestions,250)};$("#qCat").onchange=loadQuestions;$("#qGen").onchange=loadQuestions;$("#qDiff").onchange=loadQuestions;

/* ---------- video ---------- */
async function loadVideoTab(){const r=await get("/api/sets").catch(e=>({error:String(e)}));if(r.error){$("#vList").textContent=r.error;return}
  S.sets=r.sets;const want=store.get("vSet",""),cur=$("#vSet").value||want;
  $("#vSet").innerHTML=r.sets.map(s=>`<option value="${esc(s.set_id)}">${esc(s.set_id)}${s.errors.length?" (has problems)":""}</option>`).join("");
  if(cur&&r.sets.some(s=>s.set_id===cur))$("#vSet").value=cur;
  const made=r.sets.filter(s=>s.video);
  $("#vList").innerHTML=made.length?`<table><tbody>${made.map(s=>`<tr class="click" data-play="${esc(s.set_id)}"><td style="width:96px"><img class="thumb" src="${fileUrl(s.video.poster)}" alt=""></td>
    <td><strong>${esc(s.set_id)}</strong><div class="small muted">${fmtLen(s.video.length_seconds)} · ${s.video.height}p · ${s.video.seconds}s per question${s.video.media?"":" · no music/movies"}</div></td>
    <td style="text-align:right"><button class="sm">Play</button></td></tr>`).join("")}</tbody></table>`:"None yet.";
  $$("[data-play]").forEach(tr=>tr.onclick=()=>playVideo(S.sets.find(s=>s.set_id===tr.dataset.play)));vEta()}
function playVideo(s){if(!s||!s.video)return;const v=s.video;$("#vPlayerCard").classList.remove("hidden");
  $("#vPlayerTitle").textContent=`${s.set_id} · at-home video`;$("#vPlayer").poster=fileUrl(v.poster);$("#vPlayer").src=fileUrl(v.file);
  $("#vMeta").textContent=`${fmtLen(v.length_seconds)} long · ${v.height}p · ${v.seconds}s per question · ${v.answer_seconds}s per answer`+(v.skipped&&v.skipped.length?` · left out: ${v.skipped.join(", ")}`:"");
  $("#vDownload").onclick=()=>{location.href=fileUrl(v.file,true)};$("#vOpen").onclick=()=>post("/api/open",{file:v.file});
  $("#vReveal").onclick=()=>post("/api/open",{file:v.file,reveal:true});$("#vPlayerCard").scrollIntoView({behavior:"smooth",block:"nearest"})}
function vEta(){const sec=+$("#vSec").value||30,ans=+$("#vAns").value||6;const est=45*(sec+2.5+ans)+60+60;
  $("#vEta").textContent=`About ${Math.round(est/60)} minutes long (plus clip time). Making it takes a few minutes.`}
["#vSec","#vAns"].forEach(id=>$(id).oninput=()=>{store.set(id,$(id).value);vEta()});$("#vMedia").onchange=()=>store.set("vMedia",$("#vMedia").checked);
$("#vSet").onchange=()=>store.set("vSet",$("#vSet").value);
$$("#vSize button").forEach(b=>b.onclick=()=>{$$("#vSize button").forEach(x=>x.classList.remove("on"));b.classList.add("on");S.vHeight=b.dataset.h;store.set("vHeight",S.vHeight)});
$("#vBuild").onclick=()=>run({mode:"video",set_id:$("#vSet").value,seconds:+$("#vSec").value||30,answer_seconds:+$("#vAns").value||6,height:S.vHeight,media:$("#vMedia").checked});

/* ---------- files ---------- */
async function loadFiles(){const r=await get("/api/files");S.files=r.files||[];renderFiles()}
function renderFiles(){const q=$("#fSearch").value.toLowerCase(),fs=S.files.filter(f=>(!S.fKind||f.kind===S.fKind)&&f.name.toLowerCase().includes(q));
  $("#fRows").innerHTML=fs.length?fs.map(f=>`<tr><td>${esc(f.name)}${f.rel.includes("/")?` <span class="muted small">${esc(f.rel.split("/")[0])}/</span>`:""}</td>
    <td class="hide-sm"><span class="tag">${f.kind}</span></td><td class="hide-sm small">${fmtSize(f.size)}</td><td class="hide-sm small">${esc(f.modified)}</td>
    <td style="text-align:right;white-space:nowrap">${f.kind!=="deck"?`<button class="sm" data-view="${esc(f.rel)}" data-kind="${f.kind}">View</button> `:""}<button class="sm" data-open="${esc(f.rel)}">Open</button>
    <button class="sm" data-dl="${esc(f.rel)}">Download</button></td></tr>`).join(""):`<tr><td colspan="5" class="empty">No files yet. Build some decks or a video.</td></tr>`;
  $$("[data-open]").forEach(b=>b.onclick=()=>post("/api/open",{file:b.dataset.open}));
  $$("[data-dl]").forEach(b=>b.onclick=()=>{location.href=fileUrl(b.dataset.dl,true)});
  $$("[data-view]").forEach(b=>b.onclick=async()=>{const rel=b.dataset.view;
    if(b.dataset.kind==="video"){const s=S.sets.find(x=>x.video&&x.video.file===rel);if(s){showTab("video");setTimeout(()=>playVideo(s),300);return}
      $("#fView").classList.remove("hidden");$("#fViewTitle").textContent=rel;$("#fViewBody").outerHTML=`<video id="fViewBody" controls src="${fileUrl(rel)}"></video>`;return}
    const txt=await fetch(fileUrl(rel)).then(r=>r.text());$("#fView").classList.remove("hidden");$("#fViewTitle").textContent=rel;
    if($("#fViewBody").tagName!=="PRE")$("#fViewBody").outerHTML='<pre id="fViewBody" style="max-height:480px"></pre>';$("#fViewBody").textContent=txt;
    $("#fView").scrollIntoView({behavior:"smooth",block:"nearest"})})}
$$("#fKind button").forEach(b=>b.onclick=()=>{$$("#fKind button").forEach(x=>x.classList.remove("on"));b.classList.add("on");S.fKind=b.dataset.k;renderFiles()});
$("#fSearch").oninput=renderFiles;$("#fOpenFolder").onclick=()=>post("/api/open",{folder:true});$("#fViewClose").onclick=()=>$("#fView").classList.add("hidden");

/* ---------- status ---------- */
async function loadStatus(){const s=await get("/api/status");const ok=b=>b?'<span class="ok-i">✓</span>':'<span class="bad-i">✗</span>';
  const rows=[["Python",`${ok(true)} ${esc(s.python)}`],["Python packages",s.missing.length?`${ok(false)} missing: ${esc(s.missing.join(", "))}`:`${ok(true)} all installed`],
    ["ffmpeg",s.ffmpeg?`${ok(true)} installed`:`${ok(false)} not installed: needed for music/movie clips and videos`],
    ["Data source",`${esc(s.source||"?")}${s.loaded_at?` <span class="muted">(read at ${esc(s.loaded_at)})</span>`:""}`],
    ["Google Sheet link",s.config?`${ok(true)} tbtgen_config.json found`:`<span class="muted">not set up: using trivia_bank.xlsx (see README: “Connect the generator to the Google Sheet”)</span>`],
    ["TBTgen folder",`<span class="mono">${esc(s.root)}</span>`],["Decks & videos",`<span class="mono">${esc(s.out)}</span>`]];
  let html=`<div class="card"><h2>Setup</h2><div class="kv">${rows.map(([k,v])=>`<div>${k}</div><div>${v}</div>`).join("")}</div>
    <div class="row" style="margin-top:12px"><button id="testConn">${s.config?"Test Google Sheet connection":"Re-read spreadsheet"}</button><span id="testRes" class="small"></span></div></div>`;
  if(s.error)html+=`<div class="card note bad">Last load failed: ${esc(s.error)}</div>`;
  const ai=await get("/api/ai/status").catch(()=>null);
  if(ai)html+=`<div class="card"><h2>AI assistant <button class="sm" onclick="showTab('settings')" style="margin-left:8px">Set up</button></h2><div class="kv">
    <div>Provider</div><div>${ai.provider==="anthropic"?"Claude (Anthropic)":"Gemini (Google)"} · ${esc(ai.model||"no model picked")}</div>
    <div>Python package</div><div>${ok(ai.sdk[ai.provider])} ${ai.sdk[ai.provider]?"installed":`missing: <span class="mono">python3 -m pip install --user --break-system-packages ${ai.provider==="anthropic"?"anthropic":"google-genai"}</span>`}</div>
    <div>API key</div><div>${ok(ai.keys[ai.provider])} ${ai.keys[ai.provider]?"saved":"not set"}</div>
    <div>This month</div><div>${ai.usage.calls} request${ai.usage.calls==1?"":"s"} · about $${ai.usage.cost_usd.toFixed(2)}${ai.usage.unpriced_calls?` (+${ai.usage.unpriced_calls} Gemini, not priced)`:""}${ai.usage.web_searches?` · ${ai.usage.web_searches} web searches`:""}</div></div></div>`;
  if(s.counts)html+=`<div class="card"><h2>In the spreadsheet</h2><div class="kv">${Object.entries(s.counts).map(([k,v])=>`<div>${esc(k)}</div><div>${v}</div>`).join("")}</div></div>`;
  if(s.assets_missing&&s.assets_missing.length)html+=`<div class="card note">Files referenced by questions but not in the assets folder:<br><span class="mono">${s.assets_missing.map(esc).join("<br>")}</span></div>`;
  if(s.fonts)html+=`<div class="card"><h2>Video fonts</h2><p class="muted small" style="margin-top:0">Decks use Trebuchet MS, Savoye LET and Brush Script MT by name. Videos draw text themselves with these font files:</p>
    <div class="kv">${Object.entries(s.fonts).map(([k,v])=>`<div>${esc(k)}</div><div class="mono">${esc(v)}</div>`).join("")}</div></div>`;
  if(s.settings)html+=`<div class="card"><h2>Settings tab <button class="sm" onclick="showTab('settings')" style="margin-left:8px">Edit</button></h2><div class="kv">${Object.entries(s.settings).map(([k,v])=>`<div class="mono">${esc(k)}</div><div>${esc(v)}</div>`).join("")}</div></div>`;
  $("#statusBody").innerHTML=html;
  $("#testConn").onclick=async()=>{$("#testRes").textContent="Reading…";const r=await post("/api/reload");
    const msg=r.error?`<span class="bad-i">✗ ${esc(r.error)}</span>`:`<span class="ok-i">✓ read ${esc(r.source)} at ${esc(r.loaded_at)}</span>`;
    if(!r.error){refreshSource();await loadStatus()}$("#testRes").innerHTML=msg}}

/* ---------- generations ---------- */
const GENS=["Boomer","Gen X","Millennial","Gen Z","Gen Alpha"],GEN_YEARS={"Boomer":"1946–64","Gen X":"1965–80","Millennial":"1981–96","Gen Z":"1997–2012","Gen Alpha":"2013+"};
const gchips=gs=>gs&&gs.length?`<span class="gchips">${gs.map(g=>`<span class="gchip">${esc(g)}</span>`).join("")}</span>`:"";
const gpick=(name,sel)=>`<div class="gpick" data-gpick="${name}">${GENS.map(g=>`<label title="born ${GEN_YEARS[g]}"><input type="checkbox" value="${g}" ${(sel||[]).includes(g)?"checked":""}>${g}</label>`).join("")}</div>`;
const gread=(root,name)=>[...root.querySelectorAll(`[data-gpick="${name}"] input:checked`)].map(i=>i.value);
const DIFFS=["easy","medium","hard"];
const dtag=d=>d?`<span class="dtag ${esc(d)}">${esc(d)}</span>`:"";
function diffHtml(diff,total){if(!diff)return"";const rated=DIFFS.reduce((a,d)=>a+(diff[d]||0),0);if(!rated)return`<div class="small muted">No difficulty ratings yet</div>`;
  return `<div class="row small" style="gap:6px;margin-top:6px"><span class="muted">Difficulty:</span>${DIFFS.map(d=>`${dtag(d)} ${diff[d]||0}`).join(" ")}<span class="muted">· ${rated} of ${total} rated</span></div>`}
function mixHtml(mix,audience){if(!mix||!mix.total)return"";const t=Math.max(mix.tagged,1);
  return `<div class="mix">${GENS.map(g=>{const n=mix.counts[g],pct=Math.round(100*n/t),lo=(audience||[]).includes(g)&&n<0.25*mix.tagged;
    return `<span class="${lo?"lo":""}">${g}</span><span class="b"><i style="width:${pct}%"></i></span><span class="muted">${n}</span>`}).join("")}</div>
    <div class="small muted" style="margin-top:4px">Questions each generation is likely to know · ${mix.tagged} of ${mix.total} tagged${mix.tagged<mix.total?` (use Tag with AI, or tick ages in the question editor)`:""}</div>${diffHtml(mix.diff,mix.total)}`}

/* ---------- editing ---------- */
function toast(msg,err){const t=$("#toast");t.textContent=msg;t.className="show"+(err?" err":"");clearTimeout(S.toastT);S.toastT=setTimeout(()=>t.className="",err?6000:2600)}
async function edit(body){const r=await post("/api/edit",body).catch(e=>({error:String(e)}));if(r.error){toast("Couldn't save: "+r.error,true);return null}S.loaded={};return r}
function dialog(which,title,bodyHtml,buttons){const d=$(which==2?"#dlg2":"#dlg"),p=which==2?"#dlg2":"#dlg";
  $(p+"Title").textContent=title;$(p+"Body").innerHTML=bodyHtml;$(p+"Foot").innerHTML="";
  buttons.forEach(([label,cls,fn])=>{const b=document.createElement("button");b.textContent=label;if(cls)b.className=cls;
    if(label==="spacer"){const sp=document.createElement("span");sp.className="spacer";$(p+"Foot").appendChild(sp);return}
    b.onclick=async()=>{b.disabled=true;try{if(fn)await fn(d)}finally{b.disabled=false}};$(p+"Foot").appendChild(b)});
  d.querySelectorAll("[data-close]").forEach(x=>x.onclick=()=>d.close());if(!d.open)d.showModal();return d}
const val=(root,sel)=>root.querySelector(sel).value.trim();
async function assetPicker(root,name,kind,current){const box=root.querySelector(`[data-asset="${name}"]`);
  const list=(await get("/api/assets?kind="+kind)).files||[];const ph={media:"media/clip.mp3",venues:"venues/logo.png"}[kind]||"images/picture.jpg";const dl=`dl_${name}_${Math.random().toString(36).slice(2,7)}`;
  box.innerHTML=`<div class="pickrow"><input type="text" list="${dl}" value="${esc(current)}" placeholder="${ph}">
    <datalist id="${dl}">${list.map(f=>`<option value="${esc(f)}">`).join("")}</datalist>
    <label class="btn" style="border:1px solid var(--line);padding:6px 12px;border-radius:8px;cursor:pointer;font-size:14px">Upload…<input type="file" hidden accept="${kind==="media"?"audio/*,video/*":"image/*"}"></label>
    <button type="button" class="sm" data-clear>Clear</button><span class="pv"></span></div>`;
  const inp=box.querySelector("input[type=text]"),pv=box.querySelector(".pv");
  const show=()=>{const v=inp.value.trim();pv.innerHTML=!v?"":kind==="media"?`<span class="tag">${/\.(mp3|m4a|wav|aac)$/i.test(v)?"♪":"▶"} ${esc(v.split("/").pop())}</span>`:`<img src="${assetUrl(v)}" alt="" onerror="this.replaceWith(document.createTextNode('(not found)'))">`};
  inp.oninput=show;box.querySelector("[data-clear]").onclick=()=>{inp.value="";show()};show();
  box.querySelector("input[type=file]").onchange=async e=>{const f=e.target.files[0];if(!f)return;pv.textContent="Uploading…";
    const data=await new Promise(res=>{const r=new FileReader();r.onload=()=>res(r.result);r.readAsDataURL(f)});
    const r=await post("/api/upload",{kind,name:f.name,data});
    if(r.error){toast(r.error,true);show();return}inp.value=r.path;show();toast("Uploaded "+r.path)};
  return ()=>inp.value.trim()}
async function editQuestion(q,after,which){q=q||{};const isNew=!q.id;
  let getQ=()=>q.question_image||"",getA=()=>q.answer_image||"",getM=()=>q.media||"";  // until the pickers load
  const d=dialog(which||1,isNew?"New question":`Edit ${q.id}`,`<div class="f">
    <label>Question</label><div><textarea data-f="question">${esc(q.question)}</textarea>
      <div class="hint">Formatting for the question, answer and notes: <span class="mono">*italic*</span> · <span class="mono">**bold**</span> · <span class="mono">[red]red[/red]</span> · <span class="mono">[gold]gold[/gold]</span></div></div>
    <label>Answer</label><div><input type="text" data-f="answer" value="${esc(q.answer)}"></div>
    <label>Note under question</label><div><input type="text" data-f="question_note" value="${esc(q.question_note)}" placeholder="optional, e.g. One point each!">
      <div class="hint">Extra text on the question slide, shown smaller under the question</div></div>
    <label>Note with answer</label><div><input type="text" data-f="answer_note" value="${esc(q.answer_note)}" placeholder="optional, e.g. One point for title · one for artist">
      <div class="hint">Extra text revealed together with the answer (scoring note, fun fact…)</div></div>
    <label>Slide preview</label><div><div class="slideprev" id="qPrev"></div></div>
    <label>Question picture</label><div data-asset="qimg"></div>
    <label>Answer picture</label><div data-asset="aimg"><div class="hint">optional: shown when the answer is revealed</div></div>
    <label>Song / movie clip</label><div><div data-asset="media"></div><div class="hint">for Music and Movie categories</div></div>
    <label>Clip start / end</label><div class="pickrow"><input type="text" data-f="clip_start" value="${esc(q.clip_start)}" placeholder="start (s)" style="max-width:110px">
      <input type="text" data-f="clip_end" value="${esc(q.clip_end)}" placeholder="end (s)" style="max-width:110px"><span class="hint">seconds into the clip; blank = whole clip</span></div>
    <label>Seconds on screen</label><div><input type="text" data-f="seconds" value="${esc(q.seconds)}" placeholder="blank = default (60)" style="max-width:180px"></div>
    <label>Host notes</label><div><input type="text" data-f="notes" value="${esc(q.notes)}"><div class="hint">For you only: never shown on a slide</div></div>
    <label>Difficulty</label><div><select data-f="difficulty" style="max-width:200px"><option value="">not rated</option>${DIFFS.map(d=>`<option${d===q.difficulty?" selected":""}>${d}</option>`).join("")}</select>
      <div class="hint">For a typical bar team: easy = most teams get it, hard = few do</div></div>
    <label>Who'll know it</label><div>${gpick("qg",q.generations)}<div class="hint">Generations likely to know the answer: used to balance sets for each venue's crowd</div></div>
    <label>Retired</label><div><label class="row small" style="gap:6px"><input type="checkbox" data-f="retired" ${q.retired?"checked":""}> don't use this question any more</label></div>
    ${isNew?"":`<label>Used in</label><div class="small" style="padding-top:7px">${(q.categories||[]).map(c=>esc(c.title)).join(", ")||"no category yet"}</div>`}</div>`,
    [["spacer"],["Cancel","",dd=>dd.close()],[isNew?"Add question":"Save","primary",async dd=>{
      const v={question:val(dd,'[data-f=question]'),answer:val(dd,'[data-f=answer]'),question_note:val(dd,'[data-f=question_note]'),
        answer_note:val(dd,'[data-f=answer_note]'),question_image:getQ(),answer_image:getA(),media:getM(),
        clip_start:val(dd,'[data-f=clip_start]'),clip_end:val(dd,'[data-f=clip_end]'),seconds:val(dd,'[data-f=seconds]'),
        notes:val(dd,'[data-f=notes]'),retired:dd.querySelector('[data-f=retired]').checked?"Y":"",generations:gread(dd,"qg").join(", "),difficulty:dd.querySelector('[data-f=difficulty]').value};
      if(!v.question||!v.answer){toast("A question needs both the question and the answer",true);return}
      const r=await edit({op:"upsert",tab:"Questions",key:isNew?null:{question_id:q.id},values:v});if(!r)return;
      toast(isNew?`Added ${r.id}`:`Saved ${q.id}`);dd.close();if(after)after(r.id||q.id,v)}]]);
  const prev=()=>{const g=f=>val(d,`[data-f=${f}]`);d.querySelector("#qPrev").innerHTML=
    `<div class="pq">${md(g("question"))||'<span style="opacity:.5">Question text</span>'}</div>${g("question_note")?`<div class="pn">${md(g("question_note"))}</div>`:""}
    <div class="pa">${md(g("answer"))||'<span style="opacity:.5">Answer</span>'}</div>${g("answer_note")?`<div class="pn">${md(g("answer_note"))}</div>`:""}`};
  ["question","answer","question_note","answer_note"].forEach(f=>d.querySelector(`[data-f=${f}]`).addEventListener("input",prev));prev();
  [getQ,getA,getM]=await Promise.all([assetPicker(d,"qimg","images",q.question_image||""),assetPicker(d,"aimg","images",q.answer_image||""),assetPicker(d,"media","media",q.media||"")])}

/* categories */
async function loadCats(){const r=await get("/api/categories").catch(e=>({error:String(e)}));
  if(r.error){$("#cRows").innerHTML=`<tr><td colspan="5" class="empty">${esc(r.error)}</td></tr>`;return}S.cats=r;renderCats()}
function daysAgo(d){return d?Math.round((Date.parse(S.cats?.today||new Date().toISOString().slice(0,10))-Date.parse(d))/864e5):null}
function renderCats(){if(!S.cats)return;const q=$("#cSearch").value.toLowerCase(),showR=$("#cRetired").checked;
  const cs=S.cats.categories.filter(c=>(showR||!c.retired)&&(`${c.id} ${c.title} ${c.blurb} ${c.season_tag}`.toLowerCase().includes(q)));
  $("#cCount").textContent=`${cs.length} categor${cs.length==1?"y":"ies"}`;
  $("#cRows").innerHTML=cs.length?cs.map(c=>{const n=c.questions.length,ago=daysAgo(c.last_played);
    return `<tr class="click${c.retired?" dim":""}" data-cid="${esc(c.id)}"><td><strong>${esc(c.title)}</strong> <span class="mono muted">${esc(c.id)}</span>${c.retired?' <span class="tag">retired</span>':""}
      <div class="small muted">${esc(c.blurb).slice(0,110)}${c.blurb.length>110?"…":""}</div><div class="row" style="gap:4px">${(()=>{const n=Object.fromEntries(DIFFS.map(x=>[x,c.questions.filter(q=>q.difficulty===x).length]));
        return DIFFS.filter(x=>n[x]).map(x=>`${dtag(x)}<span class="small muted">${n[x]}</span>`).join(" ")})()}${gchips(GENS.filter(g=>c.gens[g]>=3))}</div></td><td class="hide-sm">${c.type!=="standard"?`<span class="tag">${esc(c.type)}</span>`:'<span class="muted small">standard</span>'}${c.season_tag?`<div class="small muted">${esc(c.season_tag)}</div>`:""}</td>
      <td>${n===5?`<span class="tag ok">5</span>`:`<span class="tag ${n?"warn":"bad"}">${n}</span>`}</td>
      <td class="hide-sm small">${c.last_played?`${fmtDate(c.last_played)} <span class="muted">(${ago}d ago)</span><div class="muted">${esc(c.played_at.join(", "))}</div>`:'<span class="muted">never</span>'}</td>
      <td class="hide-sm small" style="white-space:nowrap">${c.in_sets.map(esc).join("<br>")||'<span class="muted">—</span>'}</td></tr>`}).join("")
    :`<tr><td colspan="5" class="empty">No categories match.</td></tr>`;
  $$("#cRows tr.click").forEach(tr=>tr.onclick=()=>editCategory(S.cats.categories.find(c=>c.id===tr.dataset.cid)))}
$("#cSearch").oninput=renderCats;$("#cRetired").onchange=renderCats;$("#newCat").onclick=()=>editCategory(null);$("#ideasBtn").onclick=aiIdeas;
async function editCategory(c){c=c||{questions:[],type:"standard"};const isNew=!c.id;let slots=c.questions.map(q=>({id:q.id,question:q.question,answer:q.answer,generations:q.generations||[],difficulty:q.difficulty||""}));
  const d=dialog(1,isNew?"New category":`Edit ${c.title} (${c.id})`,`<div class="f">
    <label>Title</label><div><input type="text" data-f="title" value="${esc(c.title)}" placeholder="e.g. Turkey"></div>
    <label>Intro blurb</label><div><textarea data-f="blurb" placeholder="Shown on the category intro slide">${esc(c.blurb)}</textarea></div>
    <label>Type</label><div><select data-f="type">${["standard","music","movie"].map(t=>`<option${t===c.type?" selected":""}>${t}</option>`).join("")}</select></div>
    <label>Season</label><div><input type="text" data-f="season_tag" value="${esc(c.season_tag)}" placeholder="e.g. November (optional)"></div>
    <label>Notes</label><div><input type="text" data-f="notes" value="${esc(c.notes)}"></div>
    <label>Retired</label><div><label class="row small" style="gap:6px"><input type="checkbox" data-f="retired" ${c.retired?"checked":""}> don't use this category any more</label></div>
    <label>Questions</label><div><div id="slots"></div>
      <div class="pickrow" style="margin-top:8px"><input type="text" id="qFind" placeholder="Find a question to add (text, answer or ID)…"><button class="sm" id="qNewHere">New question…</button></div>
      <div class="res hidden" id="qFound"></div><div class="hint">A category needs exactly 5 questions. A question can be in more than one category.</div>
      <div class="row" style="margin-top:10px"><button class="sm" id="aiGenHere">Generate with AI…</button><button class="sm" id="aiCheckHere">Check these with AI</button>
        <button class="sm" id="aiTagHere">Tag with AI</button></div></div></div>`,
    [["spacer"],["Cancel","",dd=>dd.close()],[isNew?"Create category":"Save","primary",async dd=>{
      const v={title:val(dd,'[data-f=title]'),blurb:val(dd,'[data-f=blurb]'),type:val(dd,'[data-f=type]'),season_tag:val(dd,'[data-f=season_tag]'),
        notes:val(dd,'[data-f=notes]'),retired:dd.querySelector('[data-f=retired]').checked?"Y":""};
      if(!v.title){toast("Give the category a title",true);return}
      if(slots.length!==5&&!confirm(`This category has ${slots.length} question(s); a game needs 5. Save anyway?`))return;
      const r=await edit({op:"upsert",tab:"Categories",key:isNew?null:{category_id:c.id},values:v});if(!r)return;
      const cid=isNew?r.id:c.id;
      const r2=await edit({op:"replace",tab:"CategoryQuestions",match:{category_id:cid},rows:slots.map((q,i)=>({position:i+1,question_id:q.id}))});if(!r2)return;
      toast(isNew?`Created ${cid}`:`Saved ${v.title}`);dd.close();loadCats()}]]);
  const draw=()=>{d.querySelector("#slots").innerHTML=slots.length?slots.map((q,i)=>`<div class="slot"><span class="n">${i+1}</span><div class="t">${md(q.question)}<div class="small"><strong>${md(q.answer)}</strong> <span class="mono muted">${esc(q.id)}</span> ${dtag(q.difficulty)} ${gchips(q.generations)}</div></div>
      <button data-up="${i}" ${i?"":"disabled"}>↑</button><button data-dn="${i}" ${i<slots.length-1?"":"disabled"}>↓</button><button data-rm="${i}">✕</button></div>`).join("")
      :'<div class="muted small" style="padding:6px 0">No questions yet.</div>';
    d.querySelectorAll("[data-up]").forEach(b=>b.onclick=()=>{const i=+b.dataset.up;[slots[i-1],slots[i]]=[slots[i],slots[i-1]];draw()});
    d.querySelectorAll("[data-dn]").forEach(b=>b.onclick=()=>{const i=+b.dataset.dn;[slots[i+1],slots[i]]=[slots[i],slots[i+1]];draw()});
    d.querySelectorAll("[data-rm]").forEach(b=>b.onclick=()=>{slots.splice(+b.dataset.rm,1);draw();find()})};
  let ft=null;const find=async()=>{const q=d.querySelector("#qFind").value.trim(),box=d.querySelector("#qFound");if(!q){box.classList.add("hidden");return}
    const r=await get("/api/questions?q="+encodeURIComponent(q));const rs=(r.results||[]).filter(x=>!x.retired).slice(0,12);box.classList.remove("hidden");
    box.innerHTML=rs.length?rs.map(x=>`<div><span class="t">${md(x.question)}<br><span class="small"><strong>${md(x.answer)}</strong> <span class="mono muted">${esc(x.id)}</span>
      ${x.categories.length?`<span class="muted"> · in ${esc(x.categories.map(c=>c.title).join(", "))}</span>`:""}</span></span>
      <button class="sm" data-add="${esc(x.id)}" ${slots.length>=5||slots.some(s=>s.id===x.id)?"disabled":""}>${slots.some(s=>s.id===x.id)?"Added":"Add"}</button></div>`).join(""):'<div class="muted small">No matches</div>';
    box.querySelectorAll("[data-add]").forEach(b=>b.onclick=()=>{const x=rs.find(y=>y.id===b.dataset.add);slots.push({id:x.id,question:x.question,answer:x.answer});draw();find()})};
  d.querySelector("#qFind").oninput=()=>{clearTimeout(ft);ft=setTimeout(find,250)};
  d.querySelector("#qNewHere").onclick=()=>editQuestion(null,(id,v)=>{if(slots.length<5){slots.push({id,question:v.question,answer:v.answer});draw()}},2);
  d.querySelector("#aiTagHere").onclick=()=>{if(!slots.length){toast("Add some questions first",true);return}
    aiTag({question_ids:slots.map(q=>q.id)},`Tag: ${val(d,'[data-f=title]')||"category"}`,2,tags=>{slots.forEach(q=>{if(tags[q.id]){q.generations=tags[q.id].generations;q.difficulty=tags[q.id].difficulty}});draw()})};
  d.querySelector("#aiCheckHere").onclick=()=>{if(!slots.length){toast("Add some questions first",true);return}
    aiCheck({question_ids:slots.map(q=>q.id),context:val(d,'[data-f=title]')},`Check: ${val(d,'[data-f=title]')||"category"}`,2,()=>{})};
  d.querySelector("#aiGenHere").onclick=()=>aiGenerate({title:val(d,'[data-f=title]'),type:val(d,'[data-f=type]'),room:5-slots.length,gens:c.generations||[],
    blurbEmpty:!val(d,'[data-f=blurb]'),onAdd:(added,blurb)=>{added.forEach(q=>{if(slots.length<5)slots.push(q)});draw();
      if(blurb)d.querySelector('[data-f=blurb]').value=blurb;if(!val(d,'[data-f=title]')&&added.topic)d.querySelector('[data-f=title]').value=added.topic}});draw()}

/* sets */
function nextSetId(){const ids=(S.sets||[]).map(s=>s.set_id);const now=new Date(Date.now()+7*864e5);
  const d=new Date(Date.UTC(now.getFullYear(),now.getMonth(),now.getDate()));const day=d.getUTCDay()||7;d.setUTCDate(d.getUTCDate()+4-day);
  const y=d.getUTCFullYear(),w=Math.ceil(((d-Date.UTC(y,0,1))/864e5+1)/7);
  for(const L of "ABCDEFGH"){const id=`${y}-W${String(w).padStart(2,"0")}-${L}`;if(!ids.includes(id))return id}return ""}
$("#newSet").onclick=()=>editSet(null);
async function editSet(set){const isNew=!set;const cr=await get("/api/categories");const cats=cr.categories,look=cr.lookback_days;S.cats=cr;
  const cur=Array.from({length:9},(_,i)=>{const c=set&&set.rounds[Math.floor(i/3)][i%3];return c?{id:c.category_id,bonus:c.bonus}:{id:"",bonus:false}});
  const tbs=Array.from({length:3},(_,i)=>(set&&set.tiebreakers[i])||{question:"",answer:""});
  const opts=sel=>`<option value="">— pick a category —</option>`+cats.filter(c=>!c.retired||c.id===sel).map(c=>{const ago=daysAgo(c.last_played);
    return `<option value="${esc(c.id)}"${c.id===sel?" selected":""}>${esc(c.title)}${c.questions.length!==5?` (${c.questions.length}q)`:""}${c.type!=="standard"?` · ${c.type}`:""}${ago!==null?` · ${ago}d ago`:""}</option>`}).join("");
  const d=dialog(1,isNew?"New set":`Edit set ${set.set_id}`,`<div class="f" style="margin-bottom:12px"><label>Set ID</label><div>
      <input type="text" id="setId" value="${esc(isNew?nextSetId():set.set_id)}" ${isNew?"":"readonly"}><div class="hint">${isNew?"Usually year-week-letter, e.g. 2026-W49-A":"Set IDs can't be changed (nights and history refer to them)"}</div></div></div>
    <div class="sgrid">${[0,1,2].map(r=>`<div class="rnd"><h4>Round ${r+1}</h4>${[0,1,2].map(sl=>{const i=r*3+sl;return `<div class="small muted">Questions ${sl*5+1}–${sl*5+5}</div>
      <select data-slot="${i}">${opts(cur[i].id)}</select><label class="b"><input type="radio" name="bonus" value="${i}" ${cur[i].bonus?"checked":""}> bonus (double points)</label>`}).join("")}</div>`).join("")}</div>
    <div class="warnlist" id="setWarn" style="margin-top:10px"></div><div id="setMix" style="margin-top:10px"></div>
    <h4 style="margin:14px 0 6px">Tiebreakers <span class="muted small" style="font-weight:400">(year questions: closest total wins)</span></h4>
    ${tbs.map((t,i)=>`<div class="pickrow" style="margin-bottom:6px"><input type="text" data-tbq="${i}" value="${esc(t.question)}" placeholder="Tiebreaker ${i+1} question"><input type="text" data-tba="${i}" value="${esc(t.answer)}" placeholder="answer" style="max-width:110px;flex:0 0 110px"></div>`).join("")}`,
    [["spacer"],["Cancel","",dd=>dd.close()],[isNew?"Create set":"Save set","primary",async dd=>{
      const id=dd.querySelector("#setId").value.trim();if(!id){toast("Give the set an ID",true);return}
      if(isNew&&(S.sets||[]).some(s=>s.set_id===id)){toast(`There's already a set called ${id}`,true);return}
      const picks=[...dd.querySelectorAll("[data-slot]")].map(s=>s.value),b=dd.querySelector("input[name=bonus]:checked");
      if(picks.some(p=>!p)&&!confirm("Some slots are empty. Save anyway?"))return;
      const rows=picks.map((cid,i)=>({round:Math.floor(i/3)+1,slot:i%3+1,category_id:cid,bonus:b&&+b.value===i?"Y":""})).filter(r=>r.category_id);
      if(!await edit({op:"replace",tab:"Sets",match:{set_id:id},rows}))return;
      const tb=[0,1,2].map(i=>({order:i+1,question:dd.querySelector(`[data-tbq="${i}"]`).value.trim(),answer:dd.querySelector(`[data-tba="${i}"]`).value.trim()})).filter(t=>t.question);
      if(!await edit({op:"replace",tab:"Tiebreakers",match:{set_id:id},rows:tb}))return;
      toast(isNew?`Created set ${id}`:`Saved set ${id}`);dd.close();loadSets()}]]);
  const warn=()=>{const picks=[...d.querySelectorAll("[data-slot]")].map(s=>s.value),w=[];const seen={};
    picks.forEach((cid,i)=>{if(!cid)return;const c=cats.find(x=>x.id===cid);const where=`Round ${Math.floor(i/3)+1}`;
      if(seen[cid])w.push(["bad",`${c.title} is picked twice`]);seen[cid]=1;
      if(c.questions.length!==5)w.push(["bad",`${c.title} has ${c.questions.length} questions (needs 5)`]);
      const ago=daysAgo(c.last_played);if(ago!==null&&ago<=look)w.push(["warn",`${c.title} (${where}) was played ${ago} days ago at ${c.played_at.join(", ")}`]);
      if(c.questions.some(q=>q.retired))w.push(["bad",`${c.title} contains a retired question`])});
    const qs={};picks.forEach(cid=>{const c=cats.find(x=>x.id===cid);if(c)c.questions.forEach(q=>{if(qs[q.id]&&qs[q.id]!==c.title)w.push(["bad",`Question ${q.id} is in both ${qs[q.id]} and ${c.title}`]);qs[q.id]=c.title})});
    if(!d.querySelector("input[name=bonus]:checked"))w.push(["warn","No bonus category picked"]);
    const mix={counts:Object.fromEntries(GENS.map(g=>[g,0])),tagged:0,total:0,diff:Object.fromEntries(DIFFS.map(x=>[x,0]))};
    picks.forEach(cid=>{const c=cats.find(x=>x.id===cid);if(c)c.questions.forEach(q=>{mix.total++;if(q.generations&&q.generations.length){mix.tagged++;q.generations.forEach(g=>mix.counts[g]++)}
      if(mix.diff[q.difficulty]!==undefined)mix.diff[q.difficulty]++})});
    d.querySelector("#setMix").innerHTML=mix.total?`<h4 style="margin:0 0 6px">Age &amp; difficulty mix</h4>${mixHtml(mix)}`:"";
    d.querySelector("#setWarn").innerHTML=w.length?w.map(([k,t])=>`<div style="color:var(--${k})">${k==="bad"?"✗":"!"} ${esc(t)}</div>`).join(""):'<div style="color:var(--ok)">✓ Looks good</div>'};
  d.querySelectorAll("[data-slot],input[name=bonus]").forEach(x=>x.onchange=warn);warn()}

$("#newQ").onclick=()=>editQuestion(null,()=>loadQuestions());

/* venues */
async function loadVenues(){const r=await get("/api/venues").catch(e=>({error:String(e)}));
  if(r.error){$("#vnRows").innerHTML=`<tr><td colspan="6" class="empty">${esc(r.error)}</td></tr>`;return}S.venues=r.venues;renderVenues()}
function renderVenues(){const q=$("#vnSearch").value.toLowerCase(),all=$("#vnInactive").checked;
  const vs=(S.venues||[]).filter(v=>(all||v.active)&&`${v.venue_id} ${v.name} ${v.calendar_name} ${v.region}`.toLowerCase().includes(q));
  $("#vnCount").textContent=`${vs.length} venue${vs.length==1?"":"s"}`;
  $("#vnRows").innerHTML=vs.length?vs.map(v=>`<tr class="click${v.active?"":" dim"}" data-vid="${esc(v.venue_id)}">
    <td style="width:70px">${v.logo_ok?`<img class="thumb" style="width:56px;height:56px;object-fit:contain;background:#000" src="${assetUrl(v.logo)}" alt="">`:'<span class="tag warn">no logo</span>'}</td>
    <td><strong>${esc(v.name)}</strong> <span class="mono muted">${esc(v.venue_id)}</span>${v.active?"":' <span class="tag">inactive</span>'}
      <div class="small muted">${v.games_played?`${v.games_played} game${v.games_played>1?"s":""} logged · last ${fmtDate(v.last_played)}`:"no games logged yet"}</div></td>
    <td>${esc(v.night)||'<span class="muted">—</span>'}<div>${gchips(v.audience)}</div></td><td class="hide-sm">${esc(v.calendar_name)||'<span class="tag bad">missing</span>'}</td>
    <td class="hide-sm">${esc(v.region)||'<span class="muted">—</span>'}</td>
    <td class="hide-sm small">${v.next_night?`${fmtDate(v.next_night)}${v.scheduled>1?` <span class="muted">(+${v.scheduled-1} more)</span>`:""}`:'<span class="muted">none scheduled</span>'}</td></tr>`).join("")
    :`<tr><td colspan="6" class="empty">No venues match.</td></tr>`;
  $$("#vnRows tr.click").forEach(tr=>tr.onclick=()=>editVenue(S.venues.find(v=>v.venue_id===tr.dataset.vid)))}
$("#vnSearch").oninput=renderVenues;$("#vnInactive").onchange=renderVenues;$("#newVenue").onclick=()=>editVenue(null);
const slug=s=>s.toLowerCase().normalize("NFKD").replace(/[̀-ͯ]/g,"").replace(/&/g," and ").replace(/['’]/g,"").replace(/[^a-z0-9]+/g,"_").replace(/^_+|_+$/g,"").slice(0,30);
async function editVenue(v){const isNew=!v;v=v||{active:true,night:""};let idTouched=false;
  const days=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"];
  const d=dialog(1,isNew?"New venue":`Edit ${v.name}`,`<div class="f">
    <label>Name</label><div><input type="text" data-f="name" value="${esc(v.name)}" placeholder="e.g. Pop's Public House"><div class="hint">Shown on the decks</div></div>
    <label>Venue ID</label><div><input type="text" data-f="venue_id" value="${esc(v.venue_id)}" ${isNew?"":"readonly"} placeholder="e.g. pops">
      <div class="hint">${isNew?"Short code used by the schedule and history (letters, numbers, _). Can't be changed later.":"Can't be changed: the schedule and history refer to it."}</div></div>
    <label>Calendar name</label><div><input type="text" data-f="calendar_name" value="${esc(v.calendar_name)}" placeholder="e.g. Pop's">
      <div class="hint">Exactly how the venue starts its calendar events: “<b>${esc(v.calendar_name||"Pop's")}</b> - Host”</div></div>
    <label>Logo</label><div data-asset="vlogo"></div>
    <label>Night</label><div><select data-f="night"><option value="">—</option>${days.map(x=>`<option${String(v.night).slice(0,3).toLowerCase()===x.toLowerCase()?" selected":""}>${x}</option>`).join("")}</select></div>
    <label>Region</label><div><input type="text" data-f="region" value="${esc(v.region)}" placeholder="e.g. North (optional)">
      <div class="hint">Venues in the same region get a warning if they share questions on back-to-back nights</div></div>
    <label>Upcoming nights</label><div><textarea data-f="upcoming_nights" placeholder="One per line, e.g.&#10;December 2nd&#10;December 9th">${esc(v.upcoming_nights)}</textarea>
      <div class="hint">Shown on the deck's “Upcoming Trivia Nights” slide</div></div>
    <label>Audience</label><div>${gpick("va",v.audience)}<div class="hint">Who usually plays here. Nights whose set misses this crowd get flagged on the Nights tab.</div></div>
    <label>Active</label><div><label class="row small" style="gap:6px"><input type="checkbox" data-f="active" ${v.active?"checked":""}> currently hosting trivia</label></div></div>`,
    [["spacer"],["Cancel","",dd=>dd.close()],[isNew?"Add venue":"Save","primary",async dd=>{
      const vals={name:val(dd,'[data-f=name]'),calendar_name:val(dd,'[data-f=calendar_name]'),logo:getLogo(),night:val(dd,'[data-f=night]'),
        region:val(dd,'[data-f=region]'),upcoming_nights:dd.querySelector('[data-f=upcoming_nights]').value.replace(/\s+$/,""),active:dd.querySelector('[data-f=active]').checked?"Y":"N",
        audience:gread(dd,"va").join(", ")};
      const id=val(dd,'[data-f=venue_id]');
      if(!vals.name||!id){toast("A venue needs a name and an ID",true);return}
      if(isNew&&!/^[a-z0-9_]+$/.test(id)){toast("Use only lower-case letters, numbers and _ in the venue ID",true);return}
      if(isNew&&(S.venues||[]).some(x=>x.venue_id===id)){toast(`There's already a venue with ID ${id}`,true);return}
      const clash=(S.venues||[]).find(x=>x.venue_id!==id&&vals.calendar_name&&x.calendar_name.toLowerCase()===vals.calendar_name.toLowerCase());
      if(clash){toast(`${clash.name} already uses the calendar name “${vals.calendar_name}”`,true);return}
      if(!await edit({op:"upsert",tab:"Venues",key:{venue_id:id},values:vals}))return;
      toast(isNew?`Added ${vals.name}`:`Saved ${vals.name}`);dd.close();loadVenues()}]]);
  let getLogo=()=>v.logo||"";
  if(isNew){const nm=d.querySelector('[data-f=name]'),idf=d.querySelector('[data-f=venue_id]'),cal=d.querySelector('[data-f=calendar_name]');let calTouched=false;
    idf.oninput=()=>idTouched=true;cal.oninput=()=>calTouched=true;
    nm.oninput=()=>{if(!idTouched)idf.value=slug(nm.value);if(!calTouched)cal.value=nm.value}}
  getLogo=await assetPicker(d,"vlogo","venues",v.logo||"")}

/* settings */
const KNOWN_SETTINGS={
  default_seconds:{label:"Seconds per question",type:"int",min:10,max:600,help:"How long each question slide stays up in the decks (and the logo countdown)."},
  media_extra_seconds:{label:"Extra seconds after a clip",type:"int",min:0,max:300,help:"Music/movie question slides last the clip length plus this, unless a question sets its own seconds."},
  repeat_lookback_days:{label:"Repeat warning window (days)",type:"int",min:1,max:3650,help:"Warn when a venue saw a category or question within this many days."},
  video_media_extra_seconds:{label:"Video: extra seconds after a clip",type:"int",min:0,max:300,help:"In at-home videos, time to answer after a song or movie clip ends."},
  company_name:{label:"Company name",type:"text",help:""},
  instagram:{label:"Instagram handle",type:"text",help:"Shown on the Upcoming Trivia Nights slide and the video's end card."},
  facebook:{label:"Facebook page",type:"text",help:"Shown on the Upcoming Trivia Nights slide and the video's end card."},
  calendar_id:{label:"Google Calendar ID",type:"text",help:"Used by TBTgen → Pull schedule from calendar in the sheet. Ends in @group.calendar.google.com."}};
async function loadSettings(){const r=await get("/api/settings").catch(e=>({error:String(e)}));
  if(r.error){$("#stForm").innerHTML=`<div class="note bad">${esc(r.error)}</div>`;return}S.settings=r.settings;
  $("#stForm").innerHTML=r.settings.filter(st=>!st.key.startsWith("ai_")).map(st=>{const k=KNOWN_SETTINGS[st.key]||{label:st.key,type:"text",help:st.meaning};
    return `<label title="${esc(st.key)}">${esc(k.label)}</label><div><input type="text" data-key="${esc(st.key)}" data-type="${k.type}" value="${esc(st.value)}"
      ${k.type==="int"?`inputmode="numeric" style="max-width:140px"`:""}><div class="hint"><span class="mono">${esc(st.key)}</span>${k.help||st.meaning?" · "+esc(k.help||st.meaning):""}</div></div>`}).join("")||'<div class="muted">The Settings tab is empty.</div>';
  $$("#stForm input").forEach(i=>i.oninput=markDirty);markDirty();loadAiCard()}
function changedSettings(){return $$("#stForm input").filter(i=>{const st=S.settings.find(x=>x.key===i.dataset.key);return st&&i.value.trim()!==st.value})}
function markDirty(){const ch=changedSettings();$("#stSave").disabled=!ch.length;$("#stDirty").textContent=ch.length?`${ch.length} unsaved change${ch.length>1?"s":""}`:""}
$("#stSave").onclick=async()=>{const ch=changedSettings();
  for(const i of ch){const k=KNOWN_SETTINGS[i.dataset.key];if(k&&k.type==="int"){const n=Number(i.value.trim());
    if(!Number.isInteger(n)||n<k.min||n>k.max){toast(`${k.label} must be a whole number from ${k.min} to ${k.max}`,true);i.focus();return}}}
  $("#stSave").disabled=true;
  for(const i of ch){const k=KNOWN_SETTINGS[i.dataset.key];const v=k&&k.type==="int"?Number(i.value.trim()):i.value.trim();
    if(!await edit({op:"upsert",tab:"Settings",key:{key:i.dataset.key},values:{value:v}})){markDirty();return}}
  toast(`Saved ${ch.length} setting${ch.length>1?"s":""}`);loadSettings()};
$("#stAdd").onclick=async()=>{const key=$("#stNewKey").value.trim(),v=$("#stNewVal").value.trim(),m=$("#stNewMeaning").value.trim();
  if(!/^[a-z0-9_]+$/.test(key)){toast("Setting keys use lower-case letters, numbers and _",true);return}
  if((S.settings||[]).some(x=>x.key===key)){toast(`${key} is already there: change it above`,true);return}
  const r=await edit({op:"upsert",tab:"Settings",key:{key},values:{value:/^\d+$/.test(v)?Number(v):v,meaning:m}});if(!r)return;
  ["#stNewKey","#stNewVal","#stNewMeaning"].forEach(x=>$(x).value="");toast(`Added ${key}`);loadSettings()};

/* ---------- AI assistant ---------- */
const VERDICT={ok:["ok","Looks good"],revise:["warn","Suggest changes"],wrong:["bad","Answer looks wrong"],unchecked:["","Not checked"]};
const busy=msg=>`<div class="ai-busy"><div class="spin"></div><div>${esc(msg)}</div></div>`;
async function aiReady(){const st=await get("/api/ai/status");const p=st.provider;
  if(!st.sdk[p]){toast(`Install the ${p==="anthropic"?"anthropic":"google-genai"} Python package first (see Settings → AI assistant)`,true);return null}
  if(!st.keys[p]){toast(`Add your ${p==="anthropic"?"Claude":"Gemini"} API key on the Settings tab first`,true);showTab("settings");return null}
  if(p==="gemini"&&!st.model){toast("Pick a Gemini model on the Settings tab first",true);showTab("settings");return null}return st}
function aiMeta(r){return `${esc(r.provider)} · ${esc(r.model)}${r.web?` · ${r.web_searches} web search${r.web_searches==1?"":"es"}`:""}${r.cost_usd!=null?` · about $${r.cost_usd.toFixed(3)}`:""}`}
function renderAiResults(list){return list.map((r,i)=>{const [cls,label]=VERDICT[r.verdict]||VERDICT.unchecked;const sug=r.suggested_question||r.suggested_answer;
  return `<div class="airesult"><div class="row" style="margin-bottom:4px"><span class="tag ${cls}">${label}</span><span class="mono muted small">${esc(r.id)}</span>
    ${r.verdict!=="unchecked"?`<span class="muted small">confidence: ${esc(r.confidence)}</span>`:""}<span class="spacer"></span>
    ${r.verdict!=="ok"&&r.verdict!=="unchecked"||r.accept_also.length?`<button class="sm primary" data-apply="${i}">Use suggestion…</button>`:""}</div>
    <div class="q">${md(r.question)}</div><div class="small">Answer: <strong>${md(r.answer)}</strong></div>
    ${r.issues.length?`<ul>${r.issues.map(x=>`<li>${esc(x)}</li>`).join("")}</ul>`:""}
    ${sug?`<div class="sugg">${r.suggested_question?`<div><span class="muted">Suggested question:</span> ${esc(r.suggested_question)}</div>`:""}${r.suggested_answer?`<div><span class="muted">Suggested answer:</span> <strong>${esc(r.suggested_answer)}</strong></div>`:""}</div>`:""}
    ${r.accept_also.length?`<div class="small"><span class="muted">Also accept:</span> ${esc(r.accept_also.join(", "))}</div>`:""}
    ${r.fun_fact?`<div class="small"><span class="muted">Fun fact:</span> ${esc(r.fun_fact)}</div>`:""}
    ${r.sources.length?`<div class="small muted">Sources: ${r.sources.map(u=>`<a href="${esc(u)}" target="_blank" rel="noopener noreferrer">${esc(u.replace(/^https?:\/\/(www\.)?/,"").split("/")[0])}</a>`).join(" · ")}</div>`:""}</div>`}).join("")}
function noteWith(notes,r){const add=[];if(r.accept_also.length)add.push("Accept also: "+r.accept_also.join(", "));if(r.fun_fact)add.push("Fun fact: "+r.fun_fact);
  return [notes,...add].filter(Boolean).join(" | ").slice(0,500)}
async function aiCheck(body,title,which,after){if(!await aiReady())return;
  const d=dialog(which,title,busy("Checking… this can take up to a minute, longer when it searches the web."),[["spacer"],["Close","",dd=>dd.close()]]);
  const r=await post("/api/ai/check",body).catch(e=>({error:String(e)}));if(!d.open)return;
  const body_=d.querySelector(".db");if(r.error){body_.innerHTML=`<div class="note bad">${esc(r.error)}</div>`;return}
  showAiResults(d,r.results,aiMeta(r),after)}
function showAiResults(d,results,meta,after){const body_=d.querySelector(".db");const bad=results.filter(x=>x.verdict!=="ok").length;
  body_.innerHTML=`<div class="row small muted" style="margin-bottom:10px">${bad?`${bad} of ${results.length} need a look`:`All ${results.length} look good`} · ${meta}</div>${renderAiResults(results)}
    <p class="muted small">AI can be wrong: check anything surprising before you change a question.</p>`;
  body_.querySelectorAll("[data-apply]").forEach(b=>b.onclick=async()=>{const r=results[+b.dataset.apply];
    const qs=await get("/api/questions?q="+encodeURIComponent(r.id));const q=(qs.results||[]).find(x=>x.id===r.id);if(!q){toast("Couldn't find "+r.id,true);return}
    const pre={...q,question:r.suggested_question||q.question,answer:r.suggested_answer||q.answer,notes:noteWith(q.notes,r)};
    editQuestion(pre,()=>{b.textContent="Saved ✓";b.disabled=true;if(after)after()},d.id==="dlg2"?1:2)})}
async function aiCheckSet(set){if(!set||!await aiReady())return;const cats=set.rounds.flat();
  const d=dialog(1,`Check set ${set.set_id} with AI`,busy(`Checking ${cats.length} categories, one at a time…`),[["spacer"],["Close","",dd=>dd.close()]]);
  const all=[];let meta="",cost=0,searches=0;
  for(const [i,c] of cats.entries()){if(!d.open)return;d.querySelector(".ai-busy div:last-child").textContent=`Checking ${c.title} (${i+1} of ${cats.length})…`;
    const r=await post("/api/ai/check",{category_id:c.category_id}).catch(e=>({error:String(e)}));
    if(r.error){d.querySelector(".db").innerHTML=`<div class="note bad">${esc(c.title)}: ${esc(r.error)}</div>`;return}
    r.results.forEach(x=>x.id=x.id);all.push(...r.results);cost+=r.cost_usd||0;searches+=r.web_searches||0;meta=`${esc(r.provider)} · ${esc(r.model)}`}
  showAiResults(d,all,`${meta}${searches?` · ${searches} web searches`:""}${cost?` · about $${cost.toFixed(2)}`:""}`,()=>{})}
async function aiGenerate(opts){if(!await aiReady())return;const st=S.aiStatus;
  const d=dialog(2,"Generate questions with AI",`<div class="f">
    <label>Topic</label><div><input type="text" id="gTopic" value="${esc(opts.title||"")}" placeholder="e.g. Famous Thanksgiving TV episodes"></div>
    <label>How many</label><div><input type="text" id="gCount" value="${Math.max(1,Math.min(10,opts.room>0?opts.room:5))}" style="max-width:80px"><div class="hint">You'll pick which ones to keep</div></div>
    <label>Difficulty</label><div><select id="gDiff"><option value="mixed">a mix</option><option>easy</option><option>medium</option><option>hard</option></select></div>
    <label>Aim at</label><div>${gpick("gg",opts.gens)}<div class="hint">Optional: pick generations to write for. None = a general crowd.</div></div>
    <label>Notes</label><div><input type="text" id="gNotes" placeholder="optional: angle, things to include or avoid"></div></div><div id="gOut"></div>`,
    [["spacer"],["Close","",dd=>dd.close()],["Generate","primary",async dd=>{
      const topic=dd.querySelector("#gTopic").value.trim();if(!topic){toast("Give a topic",true);return}
      const out=dd.querySelector("#gOut");out.innerHTML=busy("Writing questions… this can take a minute.");
      const r=await post("/api/ai/generate",{topic,count:+dd.querySelector("#gCount").value||5,difficulty:dd.querySelector("#gDiff").value,
        notes:dd.querySelector("#gNotes").value,type:opts.type||"standard",generations:gread(dd,"gg")}).catch(e=>({error:String(e)}));
      if(r.error){out.innerHTML=`<div class="note bad">${esc(r.error)}</div>`;return}
      out.innerHTML=`<h4 style="margin:14px 0 4px">Pick the ones to keep</h4><div class="small muted">${aiMeta(r)}</div>
        ${r.questions.map((q,i)=>`<label class="gen"><input type="checkbox" data-g="${i}" ${i<Math.max(opts.room,0)?"checked":""}>
          <div><div>${esc(q.question)}</div><div class="small"><strong>${esc(q.answer)}</strong>${q.accept_also.length?` <span class="muted">(also: ${esc(q.accept_also.join(", "))})</span>`:""}
          ${q.difficulty?` <span class="tag">${esc(q.difficulty)}</span>`:""} ${gchips(q.generations)}</div>
          ${q.fun_fact?`<div class="small muted">Fun fact: ${esc(q.fun_fact)}</div>`:""}${q.picture_idea?`<div class="small muted">Picture idea: ${esc(q.picture_idea)}</div>`:""}
          ${q.sources.length?`<div class="small muted">Sources: ${q.sources.map(u=>`<a href="${esc(u)}" target="_blank" rel="noopener noreferrer">${esc(u.replace(/^https?:\/\/(www\.)?/,"").split("/")[0])}</a>`).join(" · ")}</div>`:""}</div></label>`).join("")}
        ${r.blurb?`<label class="gen"><input type="checkbox" id="gBlurb" ${opts.blurbEmpty?"checked":""}><div><div class="small muted">Category intro:</div><div>${esc(r.blurb)}</div></div></label>`:""}
        <div class="row" style="margin-top:10px"><button class="primary" id="gAdd">Add selected to the bank</button><span class="small muted">Room for ${Math.max(opts.room,0)} more in this category</span></div>`;
      out.querySelector("#gAdd").onclick=async()=>{const pick=[...out.querySelectorAll("[data-g]:checked")].map(c=>r.questions[+c.dataset.g]);
        if(!pick.length){toast("Tick at least one question",true);return}
        if(pick.length>Math.max(opts.room,0)&&!confirm(`Only ${Math.max(opts.room,0)} fit in this category. Add all ${pick.length} to the bank anyway (the extras won't be in the category)?`))return;
        out.querySelector("#gAdd").disabled=true;const added=[];
        for(const q of pick){const notes=[`AI-generated (${r.model})`,q.accept_also.length?"Accept also: "+q.accept_also.join(", "):"",q.fun_fact?"Fun fact: "+q.fun_fact:"",q.picture_idea?"Picture idea: "+q.picture_idea:""].filter(Boolean).join(" | ").slice(0,500);
          const res=await edit({op:"upsert",tab:"Questions",key:null,values:{question:q.question,answer:q.answer,notes,generations:(q.generations||[]).join(", "),difficulty:q.difficulty||""}});if(!res)break;
          added.push({id:res.id,question:q.question,answer:q.answer,generations:q.generations||[],difficulty:q.difficulty||""})}
        added.topic=topic;toast(`Added ${added.length} question${added.length==1?"":"s"} to the bank`);
        opts.onAdd(added,(out.querySelector("#gBlurb")||{}).checked?r.blurb:"");dd.close()}}]])}

async function aiTag(body,title,which,after){if(!await aiReady())return;
  const d=dialog(which,title,busy("Rating difficulty and who'll know each question… (quick and cheap: no web search)"),[["spacer"],["Close","",dd=>dd.close()]]);
  const r=await post("/api/ai/tag",body).catch(e=>({error:String(e)}));if(!d.open)return;const b=d.querySelector(".db");
  if(r.error){b.innerHTML=`<div class="note bad">${esc(r.error)}</div>`;return}
  b.innerHTML=`<div class="small muted" style="margin-bottom:10px">${esc(r.provider)} · ${esc(r.model)}${r.cost_usd!=null?` · about $${r.cost_usd.toFixed(3)}`:""} · adjust the ticks, then Save tags</div>
    ${r.results.map((x,i)=>`<div class="airesult"><div class="q">${md(x.question)}</div><div class="small">Answer: <strong>${md(x.answer)}</strong> <span class="mono muted">${esc(x.id)}</span></div>
      <div class="row" style="margin-top:6px"><select data-td="${i}" style="font-size:13px;padding:3px 6px"><option value="">not rated</option>${DIFFS.map(dd=>`<option${dd===(x.missing?x.current_difficulty:x.suggested_difficulty||x.current_difficulty)?" selected":""}>${dd}</option>`).join("")}</select>
        ${gpick("t"+i,x.missing?x.current:x.suggested)}</div>
      <div class="small muted" style="margin-top:4px">${x.missing?"The AI skipped this one.":esc(x.why)}${x.current.length||x.current_difficulty?` · currently: ${esc([x.current_difficulty,...x.current].filter(Boolean).join(", "))}`:" · not tagged yet"}</div></div>`).join("")}`;
  const foot=d.querySelector(".df");const save=document.createElement("button");save.className="primary";save.textContent="Save tags";foot.appendChild(save);
  save.onclick=async()=>{save.disabled=true;const tags={};let n=0;
    for(const [i,x] of r.results.entries()){const g=gread(b,"t"+i),dv=b.querySelector(`[data-td="${i}"]`).value;tags[x.id]={generations:g,difficulty:dv};
      const vals={};if(g.join(",")!==x.current.join(","))vals.generations=g.join(", ");if(dv!==x.current_difficulty)vals.difficulty=dv;if(!Object.keys(vals).length)continue;
      if(!await edit({op:"upsert",tab:"Questions",key:{question_id:x.id},values:vals})){save.disabled=false;return}n++}
    toast(n?`Saved tags for ${n} question${n==1?"":"s"}`:"No changes to save");d.close();if(after)after(tags)}}
async function aiIdeas(){if(!await aiReady())return;
  const d=dialog(1,"Category ideas with AI",`<p class="muted small" style="margin-top:0">Ideas only: titles, blurbs and angles to spark thinking. No questions are written, and nothing is saved unless you click Create.</p>
    <div class="f"><label>Theme</label><div><input type="text" id="iTheme" placeholder="optional: e.g. December, 90s nostalgia, food, a younger crowd"></div>
    <label>How many</label><div><input type="text" id="iCount" value="10" style="max-width:80px"></div>
    <label>Aim at</label><div>${gpick("ig",[])}<div class="hint">Optional: generations the ideas should land with</div></div>
    <label>Notes</label><div><input type="text" id="iNotes" placeholder="optional: anything to include or avoid"></div></div><div id="iOut"></div>`,
    [["spacer"],["Close","",dd=>dd.close()],["Get ideas","primary",async dd=>{const out=dd.querySelector("#iOut");out.innerHTML=busy("Brainstorming…");
      const r=await post("/api/ai/ideas",{theme:dd.querySelector("#iTheme").value,count:+dd.querySelector("#iCount").value||10,generations:gread(dd,"ig"),
        notes:dd.querySelector("#iNotes").value}).catch(e=>({error:String(e)}));
      if(r.error){out.innerHTML=`<div class="note bad">${esc(r.error)}</div>`;return}
      out.innerHTML=`<div class="small muted" style="margin:12px 0 6px">${esc(r.provider)} · ${esc(r.model)}${r.cost_usd!=null?` · about $${r.cost_usd.toFixed(3)}`:""}</div>
        ${r.ideas.map((x,i)=>`<div class="airesult"><div class="row"><strong>${esc(x.title)}</strong>${x.type!=="standard"?` <span class="tag">${esc(x.type)}</span>`:""}
          ${x.season?` <span class="tag">${esc(x.season)}</span>`:""}<span class="spacer"></span><button class="sm primary" data-create="${i}">Create category…</button></div>
          <div style="margin-top:4px">${esc(x.blurb)}</div><div class="small muted" style="margin-top:4px">${esc(x.angle)}</div><div style="margin-top:4px">${gchips(x.generations)}</div></div>`).join("")}`;
      out.querySelectorAll("[data-create]").forEach(b=>b.onclick=()=>{const x=r.ideas[+b.dataset.create];dd.close();
        editCategory({questions:[],type:x.type,title:x.title,blurb:x.blurb,season_tag:x.season,notes:`Idea: ${x.angle}`.slice(0,300),generations:x.generations})})}]])}

/* ---------- import ---------- */
const MB=b=>b>=1e9?(b/1e9).toFixed(1)+" GB":(b/1e6).toFixed(1)+" MB";
async function loadImportTab(){$("#imFolder").value=$("#imFolder").value||store.get("imFolder","");const r=await get("/api/import/runs");
  $("#imRuns").innerHTML=r.runs.length?`<table><tbody>${r.runs.map(x=>`<tr class="click" data-run="${esc(x.name)}"><td><strong>${esc(x.created.replace("T"," "))}</strong>
    <div class="small muted">${esc((x.folders||[]).join(", "))}</div></td><td class="small">${x.decks} decks · ${x.categories} categories · ${x.questions} questions</td>
    <td style="text-align:right"><button class="sm">Open</button></td></tr>`).join("")}</tbody></table>`:"None yet.";
  $$("#imRuns [data-run]").forEach(tr=>tr.onclick=()=>showImport(tr.dataset.run));
  if(!S.imRun&&r.runs.length)showImport(r.runs[0].name)}
$("#imPick").onclick=async()=>{const r=await post("/api/pick-folder");if(r.error){toast(r.error,true);return}$("#imFolder").value=r.path;store.set("imFolder",r.path)};
$("#imScan").onclick=async()=>{const f=$("#imFolder").value.trim();if(!f){toast("Choose or type a folder first",true);return}store.set("imFolder",f);
  const r=await post("/api/import/scan",{folders:[f],pictures:$("#imPics").checked,clips:$("#imClips").checked});if(r.error){toast(r.error,true);return}
  $("#imProgress").classList.remove("hidden");$("#imBar").style.width="0";$("#imPTitle").textContent="Reading decks…";$("#imPct").textContent="";
  $("#runcard").classList.remove("hidden");$("#log").innerHTML="";$("#runfiles").innerHTML="";S.polling=setInterval(poll,800);poll();setBusy(true)};
async function showImport(name){S.imRun=name;const r=await get("/api/import/run?name="+encodeURIComponent(name)).catch(e=>({error:String(e)}));
  if(r.error){toast(r.error,true);return}S.imData=r;const qs=r.questions;
  const nq=Object.keys(qs).length,newq=Object.values(qs).filter(q=>!q.in_bank).length;
  const probs=r.decks.flatMap(d=>d.problems.map(p=>[d.name,p])),unk=r.decks.flatMap(d=>d.unknown.map(u=>[d.name,`slide ${u.slide}: ${u.text}`]));
  const err=r.decks.filter(d=>d.error).length;
  const m=r.media;
  $("#imSummary").classList.remove("hidden");
  $("#imSummary").innerHTML=`<div class="row" style="margin-bottom:8px"><h2 style="margin:0">Scan of ${esc(r.created.replace("T"," "))}</h2><span class="spacer"></span>
      ${r.has_review?`<button class="sm" id="imReview">Open review spreadsheet</button>`:""}</div>
    <div class="kv"><div>Decks</div><div>${r.decks.length}${err?` <span class="tag bad">${err} couldn't be read</span>`:""}</div>
      <div>Categories</div><div>${r.categories.length}</div><div>Questions</div><div>${nq} · <strong>${newq} new</strong> · ${nq-newq} already in the bank</div>
      <div>Pictures</div><div>${m.images.count} (${MB(m.images.bytes)}) ${m.images.imported?'<span class="tag ok">imported</span>':'<span class="tag">counted only</span>'}</div>
      <div>Song &amp; movie clips</div><div>${m.clips.count} (${MB(m.clips.bytes)}) ${m.clips.imported?'<span class="tag ok">imported</span>':'<span class="tag">counted only</span>'}</div></div>
    ${probs.length||unk.length?`<details style="margin-top:10px"><summary class="small">Check these: ${probs.length} problem${probs.length==1?"":"s"}${unk.length?`, ${unk.length} slide${unk.length==1?"":"s"} that look like questions but weren't recognised`:""}</summary>
      <div class="small" style="margin-top:6px;max-height:260px;overflow:auto">${[...unk.map(x=>`<div style="color:var(--warn)">? ${esc(x[0])}: ${esc(x[1])}</div>`),...probs.map(x=>`<div class="muted">• ${esc(x[0])}: ${esc(x[1])}</div>`)].join("")}</div></details>`:""}
    <details style="margin-top:6px"><summary class="small">Decks</summary><table style="margin-top:6px"><tbody>${r.decks.map(d=>`<tr><td class="small">${esc(d.name)}</td>
      <td class="small">${esc(d.date)} <span class="muted">(${esc(d.date_source)})</span></td><td class="small">${esc(d.venue_guess||"")}</td>
      <td class="small muted">${d.error?`<span style="color:var(--bad)">${esc(d.error)}</span>`:`${d.counts.question||0} questions · ${d.counts.answer||0} answers`}</td></tr>`).join("")}</tbody></table></details>`;
  const rv=$("#imReview");if(rv)rv.onclick=()=>post("/api/import/open-review",{run:name});
  S.imSel=new Set(r.categories.filter(c=>!c.added_as&&c.new_questions>0).map(c=>c.key));renderImportCats()}
function renderImportCats(){const r=S.imData;if(!r)return;const qs=r.questions,q=($("#imFilter")||{value:""}).value?.toLowerCase()||"",onlyNew=($("#imOnlyNew")||{checked:false}).checked;
  const cats=r.categories.filter(c=>(!onlyNew||(!c.added_as&&c.new_questions>0))&&`${c.title} ${c.blurb}`.toLowerCase().includes(q));
  const img=rel=>rel?`<img class="thumb" loading="lazy" style="width:54px;height:40px" src="/api/import/asset?run=${encodeURIComponent(r.name)}&name=${encodeURIComponent(rel)}&t=${T}" alt="">`:"";
  $("#imCats").classList.remove("hidden");
  $("#imCats").innerHTML=`<div class="row" style="margin-bottom:8px"><h2 style="margin:0">Categories found</h2><span class="spacer"></span>
      <input type="search" id="imFilter" placeholder="Filter…" value="${esc(q)}" style="min-width:160px"><label class="row small" style="gap:5px"><input type="checkbox" id="imOnlyNew" ${onlyNew?"checked":""}> only new</label>
      <button class="sm" id="imAll">Select all new</button><button class="sm" id="imNone">Select none</button></div>
    <table><tbody>${cats.map(c=>{const st=c.added_as?`<span class="tag ok">added as ${esc(c.added_as)}</span>`:c.new_questions?`<span class="tag warn">${c.new_questions} new question${c.new_questions==1?"":"s"}</span>`:'<span class="tag">all in bank</span>';
      const dates=c.appearances.map(a=>a.date).sort();
      return `<tr><td class="c"><input type="checkbox" data-ck="${esc(c.key)}" ${S.imSel.has(c.key)?"checked":""} ${c.added_as?"disabled":""}></td>
        <td><strong>${esc(c.title)}</strong>${c.type!=="standard"?` <span class="tag">${esc(c.type)}</span>`:""} ${st}
          <div class="small muted">${esc(c.blurb).slice(0,120)}</div>
          <details><summary class="small">${c.questions.length} question${c.questions.length==1?"":"s"}</summary>${c.questions.map(k=>{const x=qs[k];
            return `<div class="row small" style="margin:4px 0;align-items:flex-start">${img(x.question_image)}${img(x.answer_image)}<div style="flex:1">${md(x.question)}${x.question_note?` <i class="muted">${md(x.question_note)}</i>`:""}<br><strong>${md(x.answer||"(no answer found)")}</strong>${x.answer_note?` <i class="muted">${md(x.answer_note)}</i>`:""}
              ${x.media?' <span class="tag">♪/▶ clip</span>':x.media_names&&x.media_names.length&&!r.media.clips.imported?' <span class="tag">clip not imported</span>':""}${x.in_bank?` <span class="tag">${esc(x.in_bank)}</span>`:""}</div></div>`}).join("")}</details></td>
        <td class="small hide-sm">used ${c.appearances.length}× <div class="muted">${esc(dates[0])}${dates.length>1?` – ${esc(dates[dates.length-1])}`:""}</div></td></tr>`}).join("")||'<tr><td class="empty">Nothing matches.</td></tr>'}</tbody></table>
    <div class="row" style="margin-top:12px"><button class="primary" id="imAdd"></button>
      <label class="row small" style="gap:6px" title="For decks whose file or folder name includes a venue name"><input type="checkbox" id="imHistory"> also record play history (venue + date from file names)</label></div>`;
  const upd=()=>{const n=S.imSel.size;$("#imAdd").textContent=n?`Add ${n} categor${n==1?"y":"ies"} to the bank`:"Add to the bank";$("#imAdd").disabled=!n};
  $$("#imCats [data-ck]").forEach(cb=>cb.onchange=()=>{cb.checked?S.imSel.add(cb.dataset.ck):S.imSel.delete(cb.dataset.ck);upd()});upd();
  $("#imFilter").oninput=()=>{clearTimeout(S.imT);S.imT=setTimeout(()=>{renderImportCats();const f=$("#imFilter");f.focus();f.setSelectionRange(f.value.length,f.value.length)},250)};
  $("#imOnlyNew").onchange=renderImportCats;
  $("#imAll").onclick=()=>{r.categories.filter(c=>!c.added_as&&c.new_questions>0).forEach(c=>S.imSel.add(c.key));renderImportCats()};
  $("#imNone").onclick=()=>{S.imSel.clear();renderImportCats()};
  $("#imAdd").onclick=async()=>{const n=S.imSel.size;if(!n)return;if(!confirm(`Add ${n} categor${n==1?"y":"ies"} and their new questions to the question bank?`))return;
    $("#imAdd").disabled=true;$("#imAdd").textContent="Adding…";
    const res=await post("/api/import/apply",{run:r.name,category_keys:[...S.imSel],history:$("#imHistory").checked}).catch(e=>({error:String(e)}));
    if(res.error){toast("Couldn't add: "+res.error,true);$("#imAdd").disabled=false;return}S.loaded={};
    toast(`Added ${res.categories_added} categor${res.categories_added==1?"y":"ies"} and ${res.questions_added} new question${res.questions_added==1?"":"s"}${res.history_rows?` · ${res.history_rows} history rows`:""}`);
    showImport(r.name)}}

/* AI settings */
async function loadAiCard(){const st=await get("/api/ai/status");S.aiStatus=st;
  $("#aiProvider").value=st.provider;$("#aiModel").value=(S.settings||[]).find(x=>x.key==="ai_model")?.value||"";
  $("#aiModel").placeholder=st.provider==="anthropic"?"claude-opus-5-5 (default)":"pick one with List models";
  $("#aiEffort").value=st.effort;$("#aiWeb").checked=st.web;
  $("#aiKeyA").placeholder=st.keys.anthropic?"saved ✓ (type to replace)":"not set";$("#aiKeyG").placeholder=st.keys.gemini?"saved ✓ (type to replace)":"not set";
  const miss=["anthropic","gemini"].filter(p=>!st.sdk[p]).map(p=>p==="anthropic"?"anthropic":"google-genai");
  $("#aiModelHint").innerHTML=miss.length?`To use ${miss.join(" / ")}, install it: <span class="mono">python3 -m pip install --user --break-system-packages ${miss.join(" ")}</span>`:"Blank = the default for the provider.";
  const u=st.usage;$("#aiUsage").textContent=u.calls?`This month: ${u.calls} AI request${u.calls==1?"":"s"} · about $${u.cost_usd.toFixed(2)}${u.unpriced_calls?` (+${u.unpriced_calls} Gemini, not priced)`:""}${u.web_searches?` · ${u.web_searches} web searches`:""}`:"No AI requests yet this month."}
$("#aiProvider").onchange=()=>{$("#aiModel").value="";$("#aiModelList").innerHTML="";$("#aiModel").placeholder=$("#aiProvider").value==="anthropic"?"claude-opus-5-5 (default)":"pick one with List models"};
$("#aiList").onclick=async()=>{$("#aiList").disabled=true;$("#aiList").textContent="Asking…";
  const r=await post("/api/ai/models",{provider:$("#aiProvider").value}).catch(e=>({error:String(e)}));$("#aiList").disabled=false;$("#aiList").textContent="List models";
  if(r.error){toast(r.error,true);return}$("#aiModelList").innerHTML=r.models.map(m=>`<option value="${esc(m)}">`).join("");
  toast(`${r.models.length} models available: click the Model box to pick one`)};
$("#aiSave").onclick=async()=>{const keys={anthropic_api_key:$("#aiKeyA").value.trim()||null,gemini_api_key:$("#aiKeyG").value.trim()||null};
  if(keys.anthropic_api_key||keys.gemini_api_key){const r=await post("/api/ai/keys",keys);if(r.error){toast(r.error,true);return}}
  const want={ai_provider:$("#aiProvider").value,ai_model:$("#aiModel").value.trim(),ai_effort:$("#aiEffort").value,ai_web_search:$("#aiWeb").checked?"Y":"N"};
  const have=Object.fromEntries((S.settings||[]).map(x=>[x.key,x.value]));
  const meaning={ai_provider:"AI used by Check/Generate: anthropic or gemini",ai_model:"AI model (blank = provider default)",ai_effort:"AI effort: low, medium or high",ai_web_search:"Y = AI may search the web to verify facts"};
  for(const [k,v] of Object.entries(want)){if(have[k]===v||(have[k]===undefined&&v===""))continue;
    const vals=have[k]===undefined?{value:v,meaning:meaning[k]}:{value:v};if(!await edit({op:"upsert",tab:"Settings",key:{key:k},values:vals}))return}
  $("#aiKeyA").value="";$("#aiKeyG").value="";toast("AI settings saved");await loadSettings()};

/* ---------- runs (shared) ---------- */
function cls(l){return/^(ERROR|Stopped|Google Sheet:|Traceback|Video failed|Nothing to build)/.test(l)?"ln-err":/^WARNING/.test(l)?"ln-warn":/^built/.test(l)?"ln-ok":/^(SKIPPED|NOTE)/.test(l)?"ln-skip":""}
async function run(body){const r=await post("/api/run",body);if(r.error){alert(r.error);return}
  $("#runcard").classList.remove("hidden");$("#log").innerHTML="";$("#runfiles").innerHTML="";
  if(body.mode==="video"){$("#vProgress").classList.remove("hidden");$("#vBar").style.width="0";$("#vPct").textContent="";$("#vTitle").textContent=`Making video for ${body.set_id}…`}
  S.polling=setInterval(poll,800);poll();setBusy(true)}
function setBusy(on,txt){$("#busy").classList.toggle("hidden",!on);if(txt)$("#busy").textContent=txt;$("#vBuild").disabled=on;countSel()}
async function poll(){const j=await get("/api/job");$("#runtitle").textContent=j.title+(j.done?"":"…");
  $("#log").innerHTML=j.lines.map(l=>`<span class="${cls(l)}">${esc(l)}</span>`).join("\n");$("#log").scrollTop=1e9;
  if(j.progress&&j.kind==="import"){const [a,b]=j.progress,p=Math.round(100*a/b);$("#imBar").style.width=p+"%";$("#imPct").textContent=`${a} of ${b} decks`;setBusy(true,`Import ${p}%`)}
  else if(j.progress){const [a,b]=j.progress,p=Math.round(100*a/b);$("#vBar").style.width=p+"%";$("#vPct").textContent=`${a} of ${b} slides`;setBusy(true,`Video ${p}%`)}
  else if(!j.done)setBusy(true,j.kind==="check"?"Checking…":"Building…");
  $("#runfiles").innerHTML=j.files.map(f=>`<span class="tag ok">${esc(f.split("/").pop())}</span> <button class="sm" data-o="${esc(f)}">Open</button>`).join(" ");
  $$("#runfiles [data-o]").forEach(b=>b.onclick=()=>post("/api/open",{file:b.dataset.o}));
  if(j.done){clearInterval(S.polling);S.polling=null;setBusy(false);
    const warns=j.lines.filter(l=>cls(l)==="ln-warn").length,decks=j.files.filter(f=>f.endsWith(".pptx")).length;
    $("#runtitle").textContent=j.rc!==0?`${j.title}: stopped, see the messages below`:
      j.kind==="video"||j.kind==="import"?`${j.title}: done`:decks?`Done: ${decks} deck${decks>1?"s":""} built${warns?`, ${warns} warning${warns>1?"s":""}`:""}`:
      warns?`Check finished: ${warns} warning${warns>1?"s":""}`:"Check finished: no problems found";
    if(j.kind==="import"){$("#imPTitle").textContent=j.rc===0?"Scan finished":"Scan failed: see the messages below";
      if(j.rc===0&&j.import_dir){$("#imBar").style.width="100%";S.imRun=null;await loadImportTab();showImport(j.import_dir)}}
    if(j.kind==="video"){$("#vTitle").textContent=j.rc===0?"Video ready":"Video failed: see the messages below";
      if(j.rc===0){$("#vBar").style.width="100%";await loadVideoTab();playVideo(S.sets.find(s=>s.set_id===$("#vSet").value))}}
    if(!$('[data-pane="files"]').classList.contains("hidden"))loadFiles()}}
$("#runHide").onclick=()=>$("#runcard").classList.add("hidden");$("#busy").onclick=()=>$("#runcard").classList.remove("hidden");

/* ---------- header ---------- */
async function refreshSource(){const s=await get("/api/status");$("#source").textContent="Reading: "+(s.source||"?")+(s.loaded_at?` · ${s.loaded_at}`:"");
  $("#openSheet").classList.toggle("hidden",!(s.sheet_url||s.source==="trivia_bank.xlsx"));$("#openSheet").textContent=s.sheet_url?"Open Google Sheet":"Open spreadsheet";return s}
$("#reloadData").onclick=async()=>{$("#reloadData").disabled=true;$("#reloadData").textContent="Reading…";const r=await post("/api/reload");
  $("#reloadData").disabled=false;$("#reloadData").textContent="Reload data";if(r.error){alert("Couldn't read the spreadsheet: "+r.error);return}
  S.loaded={};$$("#qCat option").slice(1).forEach(o=>o.remove());refreshSource();showTab(store.get("tab","nights"))};
$("#openSheet").onclick=()=>post("/api/open-sheet");
$("#quit").onclick=async()=>{if(S.polling&&!confirm("Something is still running. Quit anyway?"))return;await post("/api/quit");
  document.body.innerHTML='<main><h1><b>TBT</b>gen stopped.</h1><p>You can close this tab.</p></main>'};

/* ---------- start ---------- */
(async()=>{const days=store.get("days",null);if(days)S.dayOn=new Set(days);chips();
  $("#vSec").value=store.get("#vSec",30);$("#vAns").value=store.get("#vAns",6);$("#vMedia").checked=store.get("vMedia",true);
  S.vHeight=store.get("vHeight","720");$$("#vSize button").forEach(b=>b.classList.toggle("on",b.dataset.h===S.vHeight));
  const s=await refreshSource();const notes=[];
  if(s.missing.length)notes.push(`Missing Python packages: <b>${s.missing.join(", ")}</b>. In Terminal run:<br><code>python3 -m pip install --user --break-system-packages ${s.missing.join(" ")}</code><br>then restart TBTgen.`);
  if(!s.ffmpeg)notes.push(`<b>ffmpeg</b> isn't installed, so music/movie clips and videos won't work. In Terminal run: <code>brew install ffmpeg</code>`);
  $("#notes").innerHTML=notes.map(n=>`<div class="note">${n}</div>`).join("");
  const j=await get("/api/job");if(!j.done){$("#runcard").classList.remove("hidden");S.polling=setInterval(poll,800);setBusy(true)}
  if(!s.missing.length)showTab((location.hash||"").slice(1)||store.get("tab","nights"));else showTab("status")})();
</script></body></html>
"""

if __name__ == "__main__":
    main()
