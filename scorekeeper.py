"""
The scorekeeper: one self-contained web page per night, written next to the deck. It works with no internet
(open it in Safari on the laptop, or AirDrop it to a phone): teams, players, round scores, the size penalty,
live ranking and the tiebreaker. It saves as you type (in that browser) and, when there's a connection, sends
the results to the Google Sheet (Results tab) with the sheet's results-only key. Without a connection: text the
results, or save a results file and import it in TBTgen (Scores tab).
"""
import json
from pathlib import Path

import build
import scores


def page_data(bank, row, venue):
    set_id = row["set_id"]
    slots = bank["sets"].get(set_id, [])
    bonus_round = next((int(s["round"]) for s in slots if build.is_yes(s.get("bonus"))), None)
    tbs = [{"question": build.plain(t["question"]), "answer": str(t.get("answer", "")).strip()}
           for t in build.set_tiebreakers(bank, set_id)]
    st = scores.season_standings(bank, row["venue_id"], upto=row["date"])
    cfg = build.load_config()
    send = {"url": cfg["sheet_url"], "key": build.RESULTS_KEY} if cfg and build.RESULTS_KEY else None
    return {
        "venue_id": row["venue_id"], "venue": build.plain(venue.get("name", "")), "date": str(build.as_date(row["date"])),
        "host": build.plain(row.get("host", "") or ""), "set_id": set_id, "bonus_round": bonus_round,
        "rules": scores.venue_rules(venue), "roster": scores.roster(bank, row["venue_id"], row["date"]),
        "tiebreakers": tbs, "send": send, "sms": str(bank["settings"].get("results_sms", "") or "").strip(),
        "season": st["season"], "standings": scores.standings_lines(st, limit=10) if st["rows"] else [],
    }


def write(bank, row, venue, outdir, stem):
    data = json.dumps(page_data(bank, row, venue)).replace("</", "<\\/")
    title = f"Scores · {build.plain(venue.get('name', ''))} · {build.as_date(row['date'])}"
    html = PAGE.replace("__TITLE__", title.replace("<", "&lt;")).replace("__DATA__", data)
    path = Path(outdir) / f"{stem} - scorekeeper.html"
    path.write_text(html, encoding="utf-8")
    return path


PAGE = r"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>__TITLE__</title>
<style>
:root{--pink:#c2185b;--ink:#1f2937;--mut:#6b7280;--line:#e5e7eb;--bg:#f6f5f2;--ok:#15803d;--warn:#b45309;--bad:#b91c1c;--gold:#b8860b}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.4 -apple-system,system-ui,sans-serif}
main{max-width:980px;margin:0 auto;padding:14px 16px 80px}h1{font-size:21px;margin:0}h2{font-size:16px;margin:0 0 8px}
.card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px;margin:12px 0}
.mut{color:var(--mut)}.small{font-size:13px}.row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.sp{flex:1}
button{font:inherit;border:1px solid var(--line);background:#fff;border-radius:8px;padding:8px 12px;cursor:pointer}
button.pri{background:var(--pink);border-color:var(--pink);color:#fff;font-weight:600}button.sm{padding:4px 9px;font-size:13px}
input{font:inherit;border:1px solid #d1d5db;border-radius:7px;padding:7px 8px;width:100%;background:#fff}
input.n{text-align:center}table{width:100%;border-collapse:collapse}th{font-size:12px;color:var(--mut);text-align:left;font-weight:600;padding:4px}
td{padding:4px;vertical-align:middle;border-top:1px solid var(--line)}td.tot{font-weight:700;font-size:17px;text-align:center}
.chip{display:inline-block;border:1px solid var(--line);border-radius:16px;padding:4px 10px;margin:3px;background:#fff;cursor:pointer;font-size:14px}
.chip.on{opacity:.35;cursor:default}.pen{color:var(--bad);font-size:12px}.rank{font-weight:700;width:2.2em}
.tie{color:var(--warn);font-weight:600}.ok{color:var(--ok)}.bad{color:var(--bad)}textarea{width:100%;font:13px ui-monospace,monospace;border-radius:8px;border:1px solid var(--line);padding:8px}
.lead li{margin:2px 0}.bonus{background:#fff7e0}
@media (max-width:640px){thead{display:none}table,tbody,tr,td{display:block;width:100%}tr{border:1px solid var(--line);border-radius:10px;margin:8px 0;padding:6px;background:#fff}
 td{border:0;display:inline-block;width:auto;padding:3px}td.name{display:block;width:100%}td.num{width:22%}td.num::before{content:attr(data-l);display:block;font-size:11px;color:var(--mut)}
 td.tot::before{content:"Total";display:block;font-size:11px;color:var(--mut);font-weight:400}}
</style></head><body><main>
<div class="row"><div><h1 id="h"></h1><div class="mut small" id="sub"></div></div><span class="sp"></span><span class="small mut" id="saved"></span></div>
<div class="card"><div class="row"><h2 style="margin:0">Teams tonight</h2><span class="sp"></span><span class="small mut" id="rules"></span></div>
 <div class="small mut" style="margin:6px 0">Tap a team that's here to add it, or type a new one at the bottom.</div><div id="roster"></div></div>
<div class="card" style="padding:8px 10px"><table><thead><tr><th>Team</th><th>Players</th><th id="h1">Round 1</th><th id="h2">Round 2</th><th id="h3">Round 3</th><th style="text-align:center">Total</th><th></th></tr></thead>
 <tbody id="rows"></tbody></table>
 <div class="row" style="margin-top:8px"><input id="newName" placeholder="New team name" style="max-width:320px"><button id="addNew">Add team</button></div></div>
<div class="card"><h2>Standings tonight</h2><ol class="lead" id="lead"></ol><div id="tieBox"></div></div>
<div class="card"><h2>Send the results</h2>
 <div class="row"><button class="pri" id="send">Send to Trivia By Troy</button><button id="text">Text the results</button><button id="file">Save results file</button><span class="sp"></span><button class="sm" id="clear">Start over</button></div>
 <div class="small" id="sendMsg" style="margin-top:8px"></div>
 <textarea id="summary" rows="6" style="margin-top:8px" readonly></textarea></div>
<div class="card" id="seasonCard"><h2 id="seasonH">Season standings before tonight</h2><ol class="lead small" id="season"></ol></div>
</main>
<script>
const D=__DATA__;
const KEY="tbt-scores-"+D.venue_id+"-"+D.date;
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const $=s=>document.querySelector(s);
let S;try{S=JSON.parse(localStorage.getItem(KEY))}catch(e){}
S=S||{teams:[],tb:{},sentAt:""};
const save=()=>{try{localStorage.setItem(KEY,JSON.stringify(S));$("#saved").textContent="Saved on this device "+new Date().toLocaleTimeString([], {hour:"numeric",minute:"2-digit"})}catch(e){$("#saved").textContent="⚠ can't save in this browser"}};
const num=v=>{const n=parseFloat(v);return isFinite(n)?n:0};
const R=D.rules;
const penalty=t=>{const extra=Math.max(0,Math.floor(num(t.players))-R.team_max),p=extra*R.size_penalty;return R.penalty_per==="round"?p*3:p};
const total=t=>num(t.r1)+num(t.r2)+num(t.r3)-penalty(t);
const tbAnswers=D.tiebreakers.map(t=>{const m=String(t.answer).match(/-?\d+(\.\d+)?/);return m?parseFloat(m[0]):null});
const tbOff=t=>{const g=S.tb[t.key];if(!g)return null;let s=0,any=false;tbAnswers.forEach((a,i)=>{if(a!==null&&g[i]!==undefined&&g[i]!==""){s+=Math.abs(num(g[i])-a);any=true}});return any?s:null};
function ranked(){return S.teams.filter(t=>t.name.trim()).map(t=>({t,tot:total(t),off:tbOff(t)})).sort((a,b)=>b.tot-a.tot||((a.off??1e9)-(b.off??1e9))||a.t.name.localeCompare(b.t.name))}
function header(){$("#h").textContent=D.venue+" · Trivia Night";
  $("#sub").textContent=new Date(D.date+"T12:00").toLocaleDateString([], {weekday:"long",month:"long",day:"numeric"})+(D.host?" · host "+D.host:"")+" · set "+D.set_id;
  $("#rules").textContent=`${R.team_max} per team · −${R.size_penalty} per extra member ${R.penalty_per==="round"?"each round":"per game"}`;
  [1,2,3].forEach(i=>{const n=S.teams.filter(t=>t.name.trim()&&t["r"+i]!==""&&t["r"+i]!==undefined).length;$("#h"+i).innerHTML=`Round ${i}${D.bonus_round===i?" ★":""}<div style="font-weight:400">${n} scored</div>`});
  if(!D.standings.length)$("#seasonCard").style.display="none";else{$("#seasonH").textContent=`Season ${D.season??""} standings before tonight`;$("#season").innerHTML=D.standings.map(x=>`<li style="list-style:none">${esc(x)}</li>`).join("")}}
function drawRoster(){const here=new Set(S.teams.map(t=>t.team_id).filter(Boolean));
  $("#roster").innerHTML=D.roster.length?D.roster.map(r=>`<span class="chip${here.has(r.team_id)?" on":""}" data-id="${esc(r.team_id)}">${esc(r.name)}</span>`).join(""):'<span class="small mut">No teams on file for this venue yet.</span>';
  document.querySelectorAll("#roster .chip:not(.on)").forEach(c=>c.onclick=()=>{const r=D.roster.find(x=>x.team_id===c.dataset.id);addTeam(r.name,r.team_id)})}
function addTeam(name,id){name=String(name||"").trim();if(!name)return;
  if(S.teams.some(t=>t.name.trim().toLowerCase()===name.toLowerCase())){alert(name+" is already on the list");return}
  const known=id||((D.roster.find(r=>r.name.toLowerCase()===name.toLowerCase())||{}).team_id)||"";
  S.teams.push({key:Date.now()+"-"+Math.random().toString(36).slice(2,7),team_id:known,name,players:"",r1:"",r2:"",r3:""});save();drawAll()}
function drawRows(){$("#rows").innerHTML=S.teams.map((t,i)=>{const p=penalty(t);
  return `<tr data-i="${i}"><td class="name"><input data-f="name" value="${esc(t.name)}">${t.team_id?"":'<div class="small mut">new team</div>'}</td>
    <td class="num" data-l="Players"><input class="n" data-f="players" inputmode="numeric" value="${esc(t.players)}"><div class="pen">${p?`−${p} pts`:""}</div></td>
    ${[1,2,3].map(r=>`<td class="num${D.bonus_round===r?" bonus":""}" data-l="Round ${r}${D.bonus_round===r?" ★":""}"><input class="n" data-f="r${r}" inputmode="decimal" value="${esc(t["r"+r])}"></td>`).join("")}
    <td class="tot">${total(t)}</td><td><button class="sm" data-del="${i}" title="Remove">✕</button></td></tr>`}).join("")||'<tr><td colspan="7" class="mut small">No teams yet.</td></tr>';
  document.querySelectorAll("#rows input").forEach(inp=>inp.oninput=()=>{const t=S.teams[+inp.closest("tr").dataset.i];t[inp.dataset.f]=inp.value;
    if(inp.dataset.f==="name")t.team_id=((D.roster.find(r=>r.name.toLowerCase()===inp.value.trim().toLowerCase())||{}).team_id)||"";
    save();const tr=inp.closest("tr");tr.querySelector(".tot").textContent=total(t);
    const pn=penalty(t);tr.querySelector(".pen").textContent=pn?`−${pn} pts`:"";drawLead();header()});
  document.querySelectorAll("[data-del]").forEach(b=>b.onclick=()=>{const t=S.teams[+b.dataset.del];if(confirm("Remove "+(t.name||"this team")+"?")){S.teams.splice(+b.dataset.del,1);save();drawAll()}})}
function drawLead(){const r=ranked();let html="",prev=null,rank=0;
  r.forEach((x,i)=>{const k=x.tot+"|"+(x.off??"");if(prev===null||k!==prev)rank=i+1;prev=k;html+=`<li style="list-style:none"><span class="rank">${rank}.</span> ${esc(x.t.name)} <b>${x.tot}</b>${x.off!==null?` <span class="small mut">(tiebreak ${x.off} off)</span>`:""}</li>`});
  $("#lead").innerHTML=html||'<li class="mut small" style="list-style:none">Scores will show here.</li>';
  // ties for 1st, 2nd or 3rd: the tiebreaker
  const top=r.length?[...new Set(r.map(x=>x.tot))].slice(0,3):[];const tied=top.map(v=>r.filter(x=>x.tot===v)).filter(g=>g.length>1&&num(g[0].tot)>0);
  if(!tied.length||!D.tiebreakers.length){$("#tieBox").innerHTML=tied.length?'<div class="tie">Tie! (no tiebreaker questions in this set)</div>':"";summary();return}
  $("#tieBox").innerHTML=`<div class="tie" style="margin-top:8px">Tie! Use the tiebreaker: enter each tied team's answers; lowest total years off wins.</div>
   ${D.tiebreakers.map((q,i)=>`<div class="small" style="margin-top:6px">${i+1}. ${esc(q.question)} <b>${esc(q.answer)}</b></div>`).join("")}
   <table style="margin-top:6px"><tbody>${tied.flat().map(x=>`<tr><td>${esc(x.t.name)}</td>${D.tiebreakers.map((q,i)=>`<td><input class="n" inputmode="numeric" data-tb="${esc(x.t.key)}" data-q="${i}" value="${esc((S.tb[x.t.key]||[])[i]??"")}" placeholder="year"></td>`).join("")}<td class="small">${x.off!==null?x.off+" off":""}</td></tr>`).join("")}</tbody></table>`;
  document.querySelectorAll("[data-tb]").forEach(inp=>inp.onchange=()=>{(S.tb[inp.dataset.tb]=S.tb[inp.dataset.tb]||[])[+inp.dataset.q]=inp.value;save();drawLead()});summary()}
function results(){return ranked().map(x=>({team_id:x.t.team_id,name:x.t.name.trim(),players:x.t.players,round1:x.t.r1,round2:x.t.r2,round3:x.t.r3,tiebreak:x.off??""}))}
function summaryText(){return `${D.venue} ${D.date}${D.host?" (host "+D.host+")":""}\n`+ranked().map((x,i)=>`${i+1}. ${x.t.name}: ${num(x.t.r1)}+${num(x.t.r2)}+${num(x.t.r3)}${penalty(x.t)?" −"+penalty(x.t):""} = ${x.tot} (${num(x.t.players)} players)${x.off!==null?" tiebreak "+x.off+" off":""}`).join("\n")}
function summary(){$("#summary").value=summaryText()}
function drawAll(){header();drawRoster();drawRows();drawLead()}
$("#addNew").onclick=()=>{addTeam($("#newName").value);$("#newName").value=""};
$("#newName").onkeydown=e=>{if(e.key==="Enter")$("#addNew").click()};
$("#clear").onclick=()=>{if(confirm("Clear every team and score for tonight?")){S={teams:[],tb:{},sentAt:""};save();drawAll()}};
$("#send").onclick=async()=>{if(!D.send){$("#sendMsg").innerHTML='<span class="bad">This page can\'t send directly. Use "Text the results" or "Save results file".</span>';return}
  if(!results().length){alert("No teams yet");return}
  $("#sendMsg").textContent="Sending…";
  try{await fetch(D.send.url,{method:"POST",mode:"no-cors",headers:{"Content-Type":"text/plain"},
      body:JSON.stringify({token:D.send.key,action:"submit_results",venue_id:D.venue_id,date:D.date,teams:results()})});
    S.sentAt=new Date().toLocaleString();save();$("#sendMsg").innerHTML=`<span class="ok">Sent ${esc(S.sentAt)}.</span> <span class="mut">Sending again later just replaces tonight's results.</span>`}
  catch(e){$("#sendMsg").innerHTML='<span class="bad">No connection right now: your scores are saved on this device. Try again when online, or text them.</span>'}};
$("#text").onclick=()=>{const t=summaryText();try{navigator.clipboard.writeText(t)}catch(e){}$("#summary").select();
  location.href="sms:"+(D.sms||"")+(/iPhone|iPad|Mac/.test(navigator.userAgent)?"&":"?")+"body="+encodeURIComponent(t)};
$("#file").onclick=()=>{const blob=new Blob([JSON.stringify({tbtgen_results:1,venue_id:D.venue_id,date:D.date,host:D.host,teams:results()},null,1)],{type:"application/json"});
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`${D.date} ${D.venue} results.json`;a.click()};
if(S.sentAt)$("#sendMsg").innerHTML=`<span class="ok">Sent ${esc(S.sentAt)}.</span>`;
drawAll();
</script></body></html>"""
