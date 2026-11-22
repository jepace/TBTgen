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
                       "host": str(r.get("host") or ""), "set_id": str(r.get("set_id") or ""), "status": status})
    nights.sort(key=lambda n: (n["date"], n["venue"]))
    return {"start": str(start), "end": str(end), "nights": nights}


def video_info(set_id):
    meta = OUT / "video" / f"{set_id} at-home.json"
    if meta.exists():
        info = json.loads(meta.read_text())
        info["file"] = f"video/{set_id} at-home.mp4"
        info["poster"] = f"video/{set_id} at-home.png"
        return info
    return None


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
                    "last_played": hist[0][0] if hist else "",
                    "played_at": sorted({venue_name(bank, v) for _, v in hist})[:6]})
        out.append({"set_id": set_id, "rounds": rounds, "errors": errs, "warnings": warns,
                    "nights": sorted(nights.get(set_id, []), key=lambda n: n["date"]), "video": video_info(set_id)})
    return {"sets": out}


def api_questions(q, cat, limit=300):
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
        hay = f"{qid} {qq['question']} {qq['answer']} {qq.get('notes', '')}".lower()
        if any(w not in hay for w in words):
            continue
        hist = sorted(by_q.get(qid, []), reverse=True)
        res.append({"id": qid, "question": str(qq["question"]), "answer": str(qq["answer"]),
                    "question_image": str(qq.get("question_image") or ""), "answer_image": str(qq.get("answer_image") or ""),
                    "media": str(qq.get("media") or ""), "retired": build.is_yes(qq.get("retired")),
                    "notes": str(qq.get("notes") or ""),
                    "categories": [{"id": c, "title": bank["cats"].get(c, {}).get("title", c)} for c in cats_of.get(qid, [])],
                    "times": len(hist), "last": hist[0][0] if hist else "",
                    "history": [{"date": d, "venue": venue_name(bank, v)} for d, v in hist[:12]]})
    res.sort(key=lambda r: r["id"])
    categories = sorted(({"id": c, "title": v.get("title", c), "retired": build.is_yes(v.get("retired"))}
                         for c, v in bank["cats"].items()), key=lambda c: c["title"].lower())
    return {"total": len(bank["qs"]), "matches": len(res), "results": res[:limit], "categories": categories}


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


def start_job(kind, title, args):
    with JOB_LOCK:
        if not JOB["done"]:
            return None
        JOB.update(id=JOB["id"] + 1, kind=kind, title=title, lines=[], done=False, rc=None, files=[], progress=None)
        job_id = JOB["id"]

    def run():
        env = dict(os.environ, PYTHONUNBUFFERED="1")
        p = subprocess.Popen([sys.executable, str(ROOT / "build.py"), *args], cwd=ROOT, stdout=subprocess.PIPE,
                             stderr=subprocess.STDOUT, text=True, env=env)
        for line in p.stdout:
            line = line.rstrip("\n")
            with JOB_LOCK:
                m = re.match(r"video: (\d+)/(\d+)$", line)
                if m:
                    JOB["progress"] = [int(m.group(1)), int(m.group(2))]
                    continue  # don't flood the log
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

    def do_GET(self):
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
                return self.send_json(api_questions(q.get("q", ""), q.get("cat", "")))
            if u.path == "/api/files":
                return self.send_json(api_files())
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
        if path == "/api/run":
            mode = body.get("mode", "build")
            if mode == "video":
                set_id = str(body.get("set_id", ""))
                if not set_id:
                    return self.send_json({"error": "pick a set"}, 400)
                args = ["--video", set_id, "--video-seconds", str(int(body.get("seconds", 30))),
                        "--video-answer-seconds", str(float(body.get("answer_seconds", 6))),
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
@media (max-width:760px){.grid3{grid-template-columns:1fr}.hide-sm{display:none}input[type=search]{min-width:0}}
</style></head><body><main>
<header><h1><b>TBT</b>gen</h1><span class="pill" id="source">…</span><span class="pill busy hidden" id="busy" title="Show progress"></span>
<span class="spacer"></span><button id="reloadData" title="Read the spreadsheet again">Reload data</button>
<button id="openSheet" class="hidden">Open spreadsheet</button><button id="quit" title="Stop TBTgen">Quit</button></header>
<nav id="tabs"><a data-tab="nights">Nights</a><a data-tab="sets">Sets</a><a data-tab="questions">Questions</a>
<a data-tab="video">Video</a><a data-tab="files">Files</a><a data-tab="status">Status</a></nav>
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
<section data-pane="sets" class="hidden"><div id="setsList"><div class="empty">Loading…</div></div></section>

<!-- ============================== QUESTIONS -->
<section data-pane="questions" class="hidden">
<div class="card"><div class="row"><input type="search" id="qSearch" placeholder="Search questions, answers, IDs…">
  <select id="qCat"><option value="">All categories</option></select><span class="muted small" id="qCount"></span></div>
  <p class="muted small" style="margin:10px 0 0">Click a question for details. Edit questions in the spreadsheet, then Reload data.</p></div>
<div class="card" style="padding:6px 16px"><table><thead><tr><th class="hide-sm"></th><th>ID</th><th>Question</th><th>Answer</th>
  <th class="hide-sm">Categories</th><th class="hide-sm">Asked</th></tr></thead><tbody id="qRows"></tbody></table></div>
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
  ({nights:()=>S.loaded.nights||loadWeek(),sets:loadSets,questions:()=>S.loaded.q||loadQuestions(),video:loadVideoTab,files:loadFiles,status:loadStatus}[t]||(()=>{}))()}
$$("nav a").forEach(a=>a.onclick=()=>showTab(a.dataset.tab));

/* ---------- nights ---------- */
function chips(){$("#days").innerHTML=DAYS.map(d=>`<button class="chip${S.dayOn.has(d)?" on":""}" data-d="${d}">${d}</button>`).join("");
  $$("#days .chip").forEach(b=>b.onclick=()=>{const d=b.dataset.d;S.dayOn.has(d)?S.dayOn.delete(d):S.dayOn.add(d);store.set("days",[...S.dayOn]);chips();renderNights()})}
function renderNights(){const shown=S.nights.filter(n=>S.dayOn.has(n.day));
  if(S.rangeTxt)$("#range").textContent=S.rangeTxt+(shown.length<S.nights.length?` · ${shown.length} shown`:"");
  $("#rows").innerHTML=shown.length?shown.map(n=>{const ok=n.status==="ok";
    const tag=ok?`<span class="tag ok">${esc(n.set_id)}</span>`:n.status==="no_set"?`<span class="tag warn">no set yet</span>`:
      n.status==="bad_set"?`<span class="tag bad">${esc(n.set_id)}: not on Sets tab</span>`:`<span class="tag bad">venue not on Venues tab</span>`;
    return `<tr class="${ok?"":"dim"}"><td class="c"><input type="checkbox" data-k="${n.date}:${esc(n.venue_id)}" ${ok&&n.sel!==false?"checked":""} ${ok?"":"disabled"}></td>
      <td>${n.day}</td><td class="hide-sm">${fmtDate(n.date)}</td><td>${esc(n.venue)}</td><td>${esc(n.host)}</td><td>${tag}</td></tr>`}).join("")
    :`<tr><td colspan="6" class="empty">No nights on the schedule for these days.</td></tr>`;
  $$("#rows input").forEach(c=>c.onchange=()=>{const n=S.nights.find(x=>`${x.date}:${x.venue_id}`===c.dataset.k);if(n)n.sel=c.checked;countSel()});countSel()}
const selected=()=>$$("#rows input:checked").map(c=>c.dataset.k);
function countSel(){const k=selected().length;$("#build").textContent=k?`Build ${k} deck${k>1?"s":""}`:"Build";$("#build").disabled=$("#check").disabled=!k||!!S.polling}
async function loadWeek(){S.loaded.nights=true;$("#range").textContent="Loading schedule…";$("#rows").innerHTML="";
  const r=await get("/api/week?week="+encodeURIComponent(S.week)).catch(e=>({error:String(e)}));
  if(r.error){$("#range").textContent="Couldn't load the schedule";$("#rows").innerHTML=`<tr><td colspan="6" class="empty">${esc(r.error)}</td></tr>`;return}
  S.nights=r.nights;S.rangeTxt=`${fmtDate(r.start)} – ${fmtDate(r.end)} · ${S.nights.length} night${S.nights.length==1?"":"s"}`;renderNights();refreshSource()}
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
      <button class="sm" data-video="${esc(s.set_id)}">${s.video?"Watch / remake video":"Make video"}</button></div>
      <div class="grid3">${rounds}</div><div class="small muted" style="margin-top:10px">Nights: ${nights}</div>
      ${probs?`<div style="margin-top:8px">${probs}</div>`:""}</div>`}).join(""):`<div class="card empty">No sets on the Sets tab yet.</div>`;
  $$("[data-video]").forEach(b=>b.onclick=()=>{store.set("vSet",b.dataset.video);showTab("video")})}

/* ---------- questions ---------- */
async function loadQuestions(){S.loaded.q=true;const q=$("#qSearch").value,cat=$("#qCat").value;
  const r=await get(`/api/questions?q=${encodeURIComponent(q)}&cat=${encodeURIComponent(cat)}`).catch(e=>({error:String(e)}));
  if(r.error){$("#qRows").innerHTML=`<tr><td colspan="6" class="empty">${esc(r.error)}</td></tr>`;return}
  if($("#qCat").options.length<=1)$("#qCat").innerHTML='<option value="">All categories</option>'+r.categories.map(c=>`<option value="${esc(c.id)}">${esc(c.title)}${c.retired?" (retired)":""}</option>`).join("");
  $("#qCount").textContent=r.matches===r.total?`${r.total} questions`:`${r.matches} of ${r.total}`+(r.results.length<r.matches?` (first ${r.results.length} shown)`:"");
  $("#qRows").innerHTML=r.results.length?r.results.map((x,i)=>{const img=x.question_image||x.answer_image;
    return `<tr class="click${x.retired?" dim":""}" data-i="${i}"><td class="hide-sm">${img?`<img class="thumb" loading="lazy" src="${assetUrl(img)}" alt="">`:x.media?'<span class="tag">'+(/\.(mp3|m4a|wav|aac)$/i.test(x.media)?"♪ audio":"▶ video")+'</span>':""}</td>
      <td class="mono">${esc(x.id)}</td><td>${esc(x.question)}${x.retired?' <span class="tag">retired</span>':""}</td><td><strong>${esc(x.answer)}</strong></td>
      <td class="hide-sm small">${x.categories.map(c=>esc(c.title)).join("<br>")||'<span class="muted">none</span>'}</td>
      <td class="hide-sm small" style="white-space:nowrap">${x.times?`${x.times}× · ${fmtDate(x.last)}`:'<span class="muted">never</span>'}</td></tr>`}).join("")
    :`<tr><td colspan="6" class="empty">No questions match.</td></tr>`;
  $$("#qRows tr.click").forEach(tr=>tr.onclick=()=>{const nx=tr.nextElementSibling;if(nx&&nx.classList.contains("detail")){nx.remove();return}
    const x=r.results[+tr.dataset.i];const d=document.createElement("tr");d.className="detail";
    d.innerHTML=`<td colspan="6"><div class="row" style="align-items:flex-start">
      ${x.question_image?`<div><div class="small muted">Question picture</div><img src="${assetUrl(x.question_image)}" alt=""></div>`:""}
      ${x.answer_image?`<div><div class="small muted">Answer picture</div><img src="${assetUrl(x.answer_image)}" alt=""></div>`:""}
      ${x.media?`<div><div class="small muted">Clip: ${esc(x.media)}</div>${/\.(mp3|m4a|wav|aac)$/i.test(x.media)?`<audio controls preload="none" src="${assetUrl(x.media)}"></audio>`:`<video controls preload="none" style="max-width:320px" src="${assetUrl(x.media)}"></video>`}</div>`:""}
      <div style="flex:1;min-width:220px"><div style="white-space:pre-wrap">${esc(x.question)}</div><div style="margin-top:6px">Answer: <strong>${esc(x.answer)}</strong></div>
      ${x.notes?`<div class="small muted">Notes: ${esc(x.notes)}</div>`:""}
      <div class="small" style="margin-top:6px">In: ${x.categories.map(c=>`${esc(c.title)} <span class="muted mono">${esc(c.id)}</span>`).join(", ")||"no category"}</div>
      <div class="small muted" style="margin-top:4px">${x.history.length?"Asked: "+x.history.map(h=>`${fmtDate(h.date)} ${esc(h.venue)}`).join(" · "):"Never asked yet"}</div></div></div></td>`;
    tr.after(d)})}
$("#qSearch").oninput=()=>{clearTimeout(S.qTimer);S.qTimer=setTimeout(loadQuestions,250)};$("#qCat").onchange=loadQuestions;

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
  if(s.counts)html+=`<div class="card"><h2>In the spreadsheet</h2><div class="kv">${Object.entries(s.counts).map(([k,v])=>`<div>${esc(k)}</div><div>${v}</div>`).join("")}</div></div>`;
  if(s.assets_missing&&s.assets_missing.length)html+=`<div class="card note">Files referenced by questions but not in the assets folder:<br><span class="mono">${s.assets_missing.map(esc).join("<br>")}</span></div>`;
  if(s.fonts)html+=`<div class="card"><h2>Video fonts</h2><p class="muted small" style="margin-top:0">Decks use Trebuchet MS, Savoye LET and Brush Script MT by name. Videos draw text themselves with these font files:</p>
    <div class="kv">${Object.entries(s.fonts).map(([k,v])=>`<div>${esc(k)}</div><div class="mono">${esc(v)}</div>`).join("")}</div></div>`;
  if(s.settings)html+=`<div class="card"><h2>Settings tab</h2><div class="kv">${Object.entries(s.settings).map(([k,v])=>`<div class="mono">${esc(k)}</div><div>${esc(v)}</div>`).join("")}</div></div>`;
  $("#statusBody").innerHTML=html;
  $("#testConn").onclick=async()=>{$("#testRes").textContent="Reading…";const r=await post("/api/reload");
    const msg=r.error?`<span class="bad-i">✗ ${esc(r.error)}</span>`:`<span class="ok-i">✓ read ${esc(r.source)} at ${esc(r.loaded_at)}</span>`;
    if(!r.error){refreshSource();await loadStatus()}$("#testRes").innerHTML=msg}}

/* ---------- runs (shared) ---------- */
function cls(l){return/^(ERROR|Stopped|Google Sheet:|Traceback|Video failed|Nothing to build)/.test(l)?"ln-err":/^WARNING/.test(l)?"ln-warn":/^built/.test(l)?"ln-ok":/^(SKIPPED|NOTE)/.test(l)?"ln-skip":""}
async function run(body){const r=await post("/api/run",body);if(r.error){alert(r.error);return}
  $("#runcard").classList.remove("hidden");$("#log").innerHTML="";$("#runfiles").innerHTML="";
  if(body.mode==="video"){$("#vProgress").classList.remove("hidden");$("#vBar").style.width="0";$("#vPct").textContent="";$("#vTitle").textContent=`Making video for ${body.set_id}…`}
  S.polling=setInterval(poll,800);poll();setBusy(true)}
function setBusy(on,txt){$("#busy").classList.toggle("hidden",!on);if(txt)$("#busy").textContent=txt;$("#vBuild").disabled=on;countSel()}
async function poll(){const j=await get("/api/job");$("#runtitle").textContent=j.title+(j.done?"":"…");
  $("#log").innerHTML=j.lines.map(l=>`<span class="${cls(l)}">${esc(l)}</span>`).join("\n");$("#log").scrollTop=1e9;
  if(j.progress){const [a,b]=j.progress,p=Math.round(100*a/b);$("#vBar").style.width=p+"%";$("#vPct").textContent=`${a} of ${b} slides`;setBusy(true,`Video ${p}%`)}
  else if(!j.done)setBusy(true,j.kind==="check"?"Checking…":"Building…");
  $("#runfiles").innerHTML=j.files.map(f=>`<span class="tag ok">${esc(f.split("/").pop())}</span> <button class="sm" data-o="${esc(f)}">Open</button>`).join(" ");
  $$("#runfiles [data-o]").forEach(b=>b.onclick=()=>post("/api/open",{file:b.dataset.o}));
  if(j.done){clearInterval(S.polling);S.polling=null;setBusy(false);
    const warns=j.lines.filter(l=>cls(l)==="ln-warn").length,decks=j.files.filter(f=>f.endsWith(".pptx")).length;
    $("#runtitle").textContent=j.rc!==0?`${j.title}: stopped, see the messages below`:
      j.kind==="video"?`${j.title}: done`:decks?`Done: ${decks} deck${decks>1?"s":""} built${warns?`, ${warns} warning${warns>1?"s":""}`:""}`:
      warns?`Check finished: ${warns} warning${warns>1?"s":""}`:"Check finished: no problems found";
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
