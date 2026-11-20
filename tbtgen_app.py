#!/usr/bin/env python3
"""
TBTgen app: a small local web page for building decks.

  python3 tbtgen_app.py        (or double-click TBTgen.command on a Mac)

Opens your browser at a page that shows a week of the schedule. Tick the nights you want, then Check or
Build. It runs build.py underneath, so everything the command line does still applies. Only this computer
can reach the page (it listens on 127.0.0.1). Close the Terminal window, or click Quit, to stop it.
"""
import datetime as dt
import json
import os
import secrets
import shutil
import socket
import subprocess
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "out"
TOKEN = secrets.token_urlsafe(16)  # the page sends this back, so other websites can't drive the app
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


# ----------------------------------------------------------------------------------------- data
def week_bounds(spec):
    today = dt.date.today()
    ref = {"this": today, "next": today + dt.timedelta(days=7)}.get(spec) or dt.date.fromisoformat(spec)
    start = ref - dt.timedelta(days=ref.weekday())
    return start, start + dt.timedelta(days=6)


def load_week(spec):
    cfg = build.load_config()
    tables = build.tables_from_sheet(cfg) if cfg else build.tables_from_xlsx(ROOT / "trivia_bank.xlsx")
    bank = build.load_bank(tables)
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
        nights.append({
            "date": str(d), "day": DAYS[d.weekday()], "venue_id": r["venue_id"],
            "venue": venue["name"] if venue else (r.get("calendar_event") or r["venue_id"] or "?"),
            "host": str(r.get("host") or ""), "set_id": str(r.get("set_id") or ""), "status": status,
        })
    nights.sort(key=lambda n: (n["date"], n["venue"]))
    return {"start": str(start), "end": str(end), "nights": nights,
            "source": "Google Sheet" if cfg else "trivia_bank.xlsx"}


# ------------------------------------------------------------------------------------------ jobs
JOB = {"id": 0, "lines": [], "done": True, "rc": None, "files": []}
JOB_LOCK = threading.Lock()


def start_job(mode, nights, record):
    with JOB_LOCK:
        if not JOB["done"]:
            return None
        JOB.update(id=JOB["id"] + 1, lines=[], done=False, rc=None, files=[])
        job_id = JOB["id"]
    args = [sys.executable, str(ROOT / "build.py"), "--only", ",".join(nights)]
    if mode == "check":
        args.append("--check")
    elif record:
        args.append("--record")

    def run():
        env = dict(os.environ, PYTHONUNBUFFERED="1")
        p = subprocess.Popen(args, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, env=env)
        for line in p.stdout:
            line = line.rstrip("\n")
            with JOB_LOCK:
                JOB["lines"].append(line)
                if line.startswith("built  "):
                    name = line[len("built  "):].split("   (")[0]
                    JOB["files"].append(Path(name).name)
        rc = p.wait()
        with JOB_LOCK:
            JOB.update(done=True, rc=rc)

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


# ------------------------------------------------------------------------------------------ http
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def send_json(self, obj, code=200):
        b = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

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
        if q.get("t") != TOKEN:
            return self.send_json({"error": "forbidden"}, 403)
        if u.path == "/api/status":
            return self.send_json({
                "missing": MISSING, "ffmpeg": bool(shutil.which("ffmpeg") and shutil.which("ffprobe")),
                "source": "Google Sheet" if (not MISSING and build.load_config()) else "trivia_bank.xlsx",
                "out": str(OUT)})
        if u.path == "/api/week":
            try:
                return self.send_json(load_week(q.get("week", "next")))
            except Exception as e:  # show the reason on the page
                return self.send_json({"error": str(e)}, 500)
        if u.path == "/api/job":
            with JOB_LOCK:
                return self.send_json(dict(JOB))
        self.send_json({"error": "not found"}, 404)

    def do_POST(self):
        if self.headers.get("X-TBTgen-Token") != TOKEN:
            return self.send_json({"error": "forbidden"}, 403)
        body = json.loads(self.rfile.read(int(self.headers.get("Content-Length") or 0)) or b"{}")
        path = urlparse(self.path).path
        if path == "/api/run":
            nights = [n for n in body.get("nights", []) if ":" in n]
            if not nights:
                return self.send_json({"error": "no nights selected"}, 400)
            job = start_job(body.get("mode", "build"), nights, bool(body.get("record")))
            return self.send_json({"job": job} if job else {"error": "a run is already going"},
                                  200 if job else 409)
        if path == "/api/open":
            name = body.get("file")
            target = OUT / Path(name).name if name else OUT
            OUT.mkdir(exist_ok=True)
            if target.exists():
                open_path(target)
            return self.send_json({"ok": target.exists()})
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
--ok:#1b7f4b;--warn:#a15c00;--bad:#b3261e;--chip:#efede8;--log:#16161a;--logink:#e8e6e3}
@media (prefers-color-scheme:dark){:root{--bg:#141416;--card:#1e1e22;--ink:#ececef;--muted:#9a9aa2;--line:#33333a;
--accent:#ff4f9a;--accent-ink:#1a0710;--ok:#5fd39a;--warn:#f0b35a;--bad:#ff8a80;--chip:#2a2a30;--log:#0d0d0f;--logink:#d6d4d0}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",system-ui,sans-serif}
main{max-width:980px;margin:0 auto;padding:24px 16px 48px}
header{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:16px}
h1{font-size:22px;margin:0;letter-spacing:-.01em}h1 b{color:var(--accent)}
.pill{font-size:12px;padding:3px 10px;border-radius:99px;background:var(--chip);color:var(--muted)}
.spacer{flex:1}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:16px}
.row{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.label{font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin-right:4px}
button,.btn{font:inherit;border:1px solid var(--line);background:var(--card);color:var(--ink);padding:7px 14px;border-radius:8px;cursor:pointer}
button:hover{border-color:var(--muted)}button:disabled{opacity:.45;cursor:default}
button.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-ink);font-weight:600}
.seg button{border-radius:0;margin-left:-1px}.seg button:first-child{border-radius:8px 0 0 8px;margin-left:0}
.seg button:last-of-type{border-radius:0 8px 8px 0}.seg button.on,.chip.on{background:var(--ink);color:var(--bg);border-color:var(--ink)}
.chip{padding:5px 11px;border-radius:99px;font-size:13px}
input[type=date]{font:inherit;padding:6px 8px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink)}
table{width:100%;border-collapse:collapse;font-size:14px}th,td{text-align:left;padding:8px 6px;border-bottom:1px solid var(--line)}
th{font-size:12px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);font-weight:600}
tr.dim td{color:var(--muted)}td.c{width:30px}
.tag{font-size:12px;padding:2px 8px;border-radius:99px;background:var(--chip);white-space:nowrap}
.tag.ok{color:var(--ok)}.tag.warn{color:var(--warn)}.tag.bad{color:var(--bad)}
.empty{color:var(--muted);padding:20px 0;text-align:center}
pre{background:var(--log);color:var(--logink);border-radius:10px;padding:12px;margin:0;max-height:300px;overflow:auto;font:12.5px/1.5 ui-monospace,Menlo,monospace;white-space:pre-wrap}
.files li{margin:4px 0}.files button{padding:2px 10px;font-size:13px;margin-left:8px}
.note{border-left:3px solid var(--warn);padding:8px 12px;background:var(--chip);border-radius:6px;margin-bottom:12px}
.note code{font-size:12.5px}
.ln-err{color:#ff8a80}.ln-warn{color:#f0b35a}.ln-ok{color:#5fd39a}.ln-skip{color:#9a9aa2}
@media (max-width:640px){.hide-sm{display:none}}
</style></head><body><main>
<header><h1><b>TBT</b>gen</h1><span class="pill" id="source">…</span><span class="spacer"></span>
<button id="quit" title="Stop TBTgen">Quit</button></header>
<div id="notes"></div>
<section class="card">
  <div class="row" style="margin-bottom:12px">
    <span class="label">Week</span>
    <div class="seg row" style="gap:0"><button data-week="this">This week</button><button data-week="next" class="on">Next week</button></div>
    <input type="date" id="pick" title="Any date in the week">
    <span class="spacer"></span><button id="reload">Reload schedule</button>
  </div>
  <div class="row"><span class="label">Days</span><span id="days" class="row" style="gap:6px"></span></div>
</section>
<section class="card">
  <div class="row" style="margin-bottom:8px"><strong id="range">Loading…</strong><span class="spacer"></span>
    <button id="all">Select all</button><button id="none">Select none</button></div>
  <table><thead><tr><th class="c"></th><th>Day</th><th class="hide-sm">Date</th><th>Venue</th><th>Host</th><th>Set</th></tr></thead>
  <tbody id="rows"></tbody></table>
  <div class="row" style="margin-top:14px">
    <button id="check">Check</button><button id="build" class="primary">Build</button>
    <label class="row" style="gap:6px" title="Writes what was played to the Usage tab, so repeat checks know about it">
      <input type="checkbox" id="record"> Log as played (Usage)</label>
    <span class="spacer"></span><button id="openout">Open decks folder</button>
  </div>
</section>
<section class="card" id="runcard" hidden>
  <div class="row" style="margin-bottom:8px"><strong id="runtitle">Running…</strong></div>
  <pre id="log"></pre>
  <ul class="files" id="files"></ul>
</section>
</main>
<script>
const T="__TOKEN__";const $=s=>document.querySelector(s);
const DAYS=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"];let week="next",nights=[],dayOn=new Set(DAYS),polling=null;
const get=p=>fetch(p+(p.includes("?")?"&":"?")+"t="+T).then(r=>r.json());
const post=(p,b)=>fetch(p,{method:"POST",headers:{"Content-Type":"application/json","X-TBTgen-Token":T},body:JSON.stringify(b||{})}).then(r=>r.json());
const esc=s=>String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
function chips(){$("#days").innerHTML=DAYS.map(d=>`<button class="chip${dayOn.has(d)?" on":""}" data-d="${d}">${d}</button>`).join("");
  document.querySelectorAll("#days .chip").forEach(b=>b.onclick=()=>{const d=b.dataset.d;dayOn.has(d)?dayOn.delete(d):dayOn.add(d);chips();render()});}
function fmtDate(s){const d=new Date(s+"T12:00:00");return d.toLocaleDateString(undefined,{month:"short",day:"numeric"})}
function render(){const shown=nights.filter(n=>dayOn.has(n.day));
  if(window.rangeTxt)$("#range").textContent=rangeTxt+(shown.length<nights.length?` · ${shown.length} shown`:"");
  $("#rows").innerHTML=shown.length?shown.map(n=>{const ok=n.status==="ok";
    const tag=ok?`<span class="tag ok">${esc(n.set_id)}</span>`:n.status==="no_set"?`<span class="tag warn">no set yet</span>`:
      n.status==="bad_set"?`<span class="tag bad">${esc(n.set_id)}?</span>`:`<span class="tag bad">venue not on Venues tab</span>`;
    return `<tr class="${ok?"":"dim"}"><td class="c"><input type="checkbox" data-k="${n.date}:${esc(n.venue_id)}" ${ok&&n.sel!==false?"checked":""} ${ok?"":"disabled"}></td>
      <td>${n.day}</td><td class="hide-sm">${fmtDate(n.date)}</td><td>${esc(n.venue)}</td><td>${esc(n.host)}</td><td>${tag}</td></tr>`}).join("")
    :`<tr><td colspan="6" class="empty">No nights on the schedule for these days.</td></tr>`;
  document.querySelectorAll("#rows input").forEach(c=>c.onchange=()=>{const n=nights.find(x=>`${x.date}:${x.venue_id}`===c.dataset.k);if(n)n.sel=c.checked;count()});count();}
function selected(){return[...document.querySelectorAll("#rows input:checked")].map(c=>c.dataset.k)}
function count(){const k=selected().length;$("#build").textContent=k?`Build ${k} deck${k>1?"s":""}`:"Build";$("#build").disabled=$("#check").disabled=!k||!!polling}
async function load(){$("#range").textContent="Loading schedule…";$("#rows").innerHTML="";
  const r=await get("/api/week?week="+encodeURIComponent(week)).catch(e=>({error:String(e)}));
  if(r.error){$("#range").textContent="Couldn't load the schedule";$("#rows").innerHTML=`<tr><td colspan="6" class="empty">${esc(r.error)}</td></tr>`;return}
  nights=r.nights;$("#source").textContent="Reading: "+r.source;
  window.rangeTxt=`${fmtDate(r.start)} – ${fmtDate(r.end)} · ${nights.length} night${nights.length==1?"":"s"}`;render()}
function cls(l){return/^(ERROR|Stopped|Google Sheet:|Traceback)/.test(l)?"ln-err":/^WARNING/.test(l)?"ln-warn":/^built/.test(l)?"ln-ok":/^(SKIPPED|NOTE)/.test(l)?"ln-skip":""}
async function run(mode){const ns=selected();if(!ns.length)return;
  const r=await post("/api/run",{mode,nights:ns,record:$("#record").checked});if(r.error){alert(r.error);return}
  $("#runcard").hidden=false;$("#runtitle").textContent=mode==="check"?"Checking…":"Building…";$("#log").innerHTML="";$("#files").innerHTML="";
  polling=setInterval(poll,700);count()}
async function poll(){const j=await get("/api/job");
  $("#log").innerHTML=j.lines.map(l=>`<span class="${cls(l)}">${esc(l)}</span>`).join("\n");$("#log").scrollTop=1e9;
  $("#files").innerHTML=j.files.map(f=>`<li>${esc(f)}<button data-f="${esc(f)}">Open</button></li>`).join("");
  document.querySelectorAll("#files button").forEach(b=>b.onclick=()=>post("/api/open",{file:b.dataset.f}));
  if(j.done){clearInterval(polling);polling=null;count();
    const errs=j.lines.filter(l=>cls(l)==="ln-err").length,warns=j.lines.filter(l=>cls(l)==="ln-warn").length;
    $("#runtitle").textContent=j.rc!==0?"Stopped: see the messages below":
      j.files.length?`Done: ${j.files.length} deck${j.files.length>1?"s":""} built${warns?`, ${warns} warning${warns>1?"s":""}`:""}`:
      warns?`Check finished: ${warns} warning${warns>1?"s":""}`:"Check finished: no problems found"}}
document.querySelectorAll(".seg button").forEach(b=>b.onclick=()=>{document.querySelectorAll(".seg button").forEach(x=>x.classList.remove("on"));b.classList.add("on");$("#pick").value="";week=b.dataset.week;load()});
$("#pick").onchange=e=>{if(!e.target.value)return;document.querySelectorAll(".seg button").forEach(x=>x.classList.remove("on"));week=e.target.value;load()};
$("#reload").onclick=load;$("#check").onclick=()=>run("check");$("#build").onclick=()=>run("build");
$("#all").onclick=()=>{nights.forEach(n=>n.sel=true);render()};$("#none").onclick=()=>{nights.forEach(n=>n.sel=false);render()};
$("#openout").onclick=()=>post("/api/open",{});
$("#quit").onclick=async()=>{await post("/api/quit");document.body.innerHTML='<main><h1><b>TBT</b>gen stopped.</h1><p>You can close this tab.</p></main>'};
(async()=>{chips();const s=await get("/api/status");$("#source").textContent="Reading: "+s.source;const notes=[];
  if(s.missing.length)notes.push(`Missing Python packages: <b>${s.missing.join(", ")}</b>. In Terminal run:<br><code>python3 -m pip install --user --break-system-packages ${s.missing.join(" ")}</code><br>then restart TBTgen.`);
  if(!s.ffmpeg)notes.push(`<b>ffmpeg</b> isn't installed, so decks with music or movie clips will fail. In Terminal run: <code>brew install ffmpeg</code>`);
  $("#notes").innerHTML=notes.map(n=>`<div class="note">${n}</div>`).join("");if(!s.missing.length)load();})();
</script></body></html>
"""

if __name__ == "__main__":
    main()
