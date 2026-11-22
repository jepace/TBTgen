"""
AI assistant for TBTgen: check questions, and generate questions for a topic.

Runs only when asked (the app's "Check with AI" / "Generate with AI" buttons). Nothing is changed in the
spreadsheet by this module: it returns suggestions, and the person decides what to keep.

Providers (pick one on the app's Settings tab -> ai_provider):
  anthropic  Claude, via the `anthropic` package. Default model claude-opus-5-5.
  gemini     Google Gemini, via the `google-genai` package. Pick a model with "List models".

API keys live in tbtgen_config.json next to this file ("anthropic_api_key", "gemini_api_key"; never in the
spreadsheet or Git), or in the ANTHROPIC_API_KEY / GEMINI_API_KEY environment variables.
Each call is logged to ai_usage.jsonl (tokens, web searches, estimated cost) for the app's monthly total.
"""
import datetime as dt
import json
import os
import random
import re
from pathlib import Path

import build

ROOT = build.ROOT
USAGE_LOG = ROOT / "ai_usage.jsonl"
PROVIDERS = ("anthropic", "gemini")
DEFAULT_MODEL = {"anthropic": "claude-opus-5-5", "gemini": ""}
# $ per million input / output tokens (Anthropic list prices); web searches are billed separately.
PRICES = {"claude-opus-5-5": (4.0, 20.0), "claude-sonnet-5-5": (2.0, 10.0), "claude-haiku-5-5": (0.10, 0.50),
          "claude-fable-5-2": (8.0, 40.0)}
# models that take the server-side refusal fallback ("default" routing)
FALLBACK_OK = {"claude-opus-5-5", "claude-sonnet-5-5", "claude-fable-5-2"}


class AIError(Exception):
    pass


# --------------------------------------------------------------------------------------- settings & keys
def _config():
    try:
        return json.loads(build.CONFIG.read_text()) if build.CONFIG.exists() else {}
    except ValueError:
        raise AIError(f"{build.CONFIG.name} isn't valid JSON")


def api_key(provider):
    cfg = _config()
    if provider == "anthropic":
        return cfg.get("anthropic_api_key") or os.environ.get("ANTHROPIC_API_KEY") or ""
    return cfg.get("gemini_api_key") or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or ""


def save_keys(updates):
    """updates: {"anthropic_api_key": str|None, "gemini_api_key": str|None}; None = leave, "" = remove."""
    cfg = _config()
    for k in ("anthropic_api_key", "gemini_api_key"):
        v = updates.get(k)
        if v is None:
            continue
        if v.strip():
            cfg[k] = v.strip()
        else:
            cfg.pop(k, None)
    build.CONFIG.write_text(json.dumps(cfg, indent=2) + "\n")
    try:
        os.chmod(build.CONFIG, 0o600)
    except OSError:
        pass


def sdk_available(provider):
    try:
        if provider == "anthropic":
            import anthropic  # noqa: F401
        else:
            from google import genai  # noqa: F401
        return True
    except ImportError:
        return False


def options(settings):
    provider = str(settings.get("ai_provider") or "anthropic").strip().lower()
    if provider not in PROVIDERS:
        provider = "anthropic"
    model = str(settings.get("ai_model") or "").strip() or DEFAULT_MODEL[provider]
    effort = str(settings.get("ai_effort") or "medium").strip().lower()
    web = str(settings.get("ai_web_search") or "Y").strip().upper() not in ("N", "NO", "FALSE", "0")
    return {"provider": provider, "model": model, "effort": effort if effort in ("low", "medium", "high") else "medium",
            "web": web}


def status(settings):
    o = options(settings)
    return {**o, "keys": {p: bool(api_key(p)) for p in PROVIDERS}, "sdk": {p: sdk_available(p) for p in PROVIDERS},
            "usage": usage_summary()}


# ------------------------------------------------------------------------------------------- providers
class Anthropic:
    name = "Claude"

    def __init__(self, key, model, effort):
        try:
            import anthropic
        except ImportError:
            raise AIError("the 'anthropic' Python package isn't installed: "
                          "python3 -m pip install --user --break-system-packages anthropic")
        self.sdk = anthropic
        self.client = anthropic.Anthropic(api_key=key) if key else anthropic.Anthropic()
        self.model, self.effort = model, effort

    def _errors(self, fn):
        a = self.sdk
        try:
            return fn()
        except a.AuthenticationError:
            raise AIError("Anthropic rejected the API key: check it on the Settings tab")
        except a.PermissionDeniedError as e:
            raise AIError(f"Anthropic: not allowed ({e.message})")
        except a.NotFoundError:
            raise AIError(f"Anthropic doesn't know the model '{self.model}': pick one with List models")
        except a.RateLimitError:
            raise AIError("Anthropic says too many requests or spending limit reached: wait a minute and try again")
        except a.BadRequestError as e:
            raise AIError(f"Anthropic: {e.message}")
        except a.APIStatusError as e:
            raise AIError(f"Anthropic had a problem ({e.status_code}): try again shortly")
        except a.APIConnectionError:
            raise AIError("couldn't reach Anthropic: check the internet connection")

    def run(self, system, user, web):
        tools = [{"type": "web_search_20260209", "name": "web_search", "max_uses": 5}] if web else []
        messages = [{"role": "user", "content": user}]
        usage = {"input_tokens": 0, "output_tokens": 0, "web_searches": 0}
        resp = None
        for _ in range(4):  # a long web-search turn can pause; resume it
            kwargs = dict(model=self.model, max_tokens=16000, system=system, messages=messages,
                          output_config={"effort": self.effort})
            if tools:
                kwargs["tools"] = tools
            if self.model in FALLBACK_OK:  # if the model declines, the API retries on a fallback model
                kwargs.update(betas=["server-side-fallback-2026-07-01"], fallbacks="default")
            resp = self._errors(lambda: self.client.beta.messages.create(**kwargs))
            u = resp.usage
            usage["input_tokens"] += u.input_tokens or 0
            usage["output_tokens"] += u.output_tokens or 0
            if u.server_tool_use and u.server_tool_use.web_search_requests:
                usage["web_searches"] += u.server_tool_use.web_search_requests
            if resp.stop_reason != "pause_turn":
                break
            messages = [*messages, {"role": "assistant", "content": resp.content}]
        if resp.stop_reason == "refusal":
            raise AIError("the AI declined this request")
        if resp.stop_reason == "max_tokens":
            raise AIError("the AI's answer was too long and got cut off: try fewer questions at once")
        usage["model"] = getattr(resp, "model", self.model) or self.model
        return "".join(b.text for b in resp.content if b.type == "text"), usage

    def list_models(self):
        return self._errors(lambda: sorted((m.id for m in self.client.models.list()), reverse=True))


class Gemini:
    name = "Gemini"

    def __init__(self, key, model, effort):
        try:
            from google import genai
            from google.genai import types
        except ImportError:
            raise AIError("the 'google-genai' Python package isn't installed: "
                          "python3 -m pip install --user --break-system-packages google-genai")
        if not key:
            raise AIError("no Gemini API key: add one on the Settings tab")
        self.genai, self.types = genai, types
        self.client = genai.Client(api_key=key)
        self.model, self.effort = model, effort

    def _errors(self, fn):
        errors = self.genai.errors
        try:
            return fn()
        except errors.ClientError as e:
            code = getattr(e, "code", None)
            if code in (401, 403):
                raise AIError("Google rejected the Gemini API key: check it on the Settings tab")
            if code == 404:
                raise AIError(f"Gemini doesn't know the model '{self.model}': pick one with List models")
            if code == 429:
                raise AIError("Gemini says too many requests or quota used up: wait and try again")
            raise AIError(f"Gemini: {getattr(e, 'message', None) or e}")
        except errors.ServerError:
            raise AIError("Gemini had a problem: try again shortly")
        except errors.APIError as e:
            raise AIError(f"Gemini: {e}")
        except OSError:
            raise AIError("couldn't reach Google: check the internet connection")

    def run(self, system, user, web):
        if not self.model:
            raise AIError("pick a Gemini model on the Settings tab (use List models)")
        t = self.types
        cfg = t.GenerateContentConfig(system_instruction=system,
                                      tools=[t.Tool(google_search=t.GoogleSearch())] if web else None)
        resp = self._errors(lambda: self.client.models.generate_content(model=self.model, contents=user, config=cfg))
        um = resp.usage_metadata
        usage = {"input_tokens": (um.prompt_token_count or 0) if um else 0,
                 "output_tokens": ((um.candidates_token_count or 0) + (um.thoughts_token_count or 0)) if um else 0,
                 "web_searches": 0, "model": self.model}
        text = resp.text or ""
        if not text:
            raise AIError("Gemini returned no answer (it may have declined): try rewording")
        return text, usage

    def list_models(self):
        def names():
            out = []
            for m in self.client.models.list():
                actions = getattr(m, "supported_actions", None) or []
                if not actions or "generateContent" in actions:
                    out.append((m.name or "").replace("models/", ""))
            return sorted((n for n in out if n.startswith("gemini")), reverse=True)
        return self._errors(names)


def provider(settings, effort=None):
    o = options(settings)
    if effort:
        o["effort"] = effort
    cls = Anthropic if o["provider"] == "anthropic" else Gemini
    return cls(api_key(o["provider"]), o["model"], o["effort"]), o


def gen_line(gens):
    gens = build.parse_gens(gens)
    return ", ".join(f"{g} (born {build.GEN_YEARS[g]})" for g in gens)


# ------------------------------------------------------------------------------------------- usage log
def log_usage(task, prov, usage, items):
    pin, pout = PRICES.get(usage.get("model", ""), (None, None))
    cost = None if pin is None else round(usage["input_tokens"] / 1e6 * pin + usage["output_tokens"] / 1e6 * pout, 4)
    rec = {"ts": dt.datetime.now().isoformat(timespec="seconds"), "task": task, "provider": prov, "items": items,
           **usage, "cost_usd": cost}
    with open(USAGE_LOG, "a") as f:
        f.write(json.dumps(rec) + "\n")
    return rec


def usage_summary():
    month = dt.date.today().strftime("%Y-%m")
    s = {"month": month, "calls": 0, "input_tokens": 0, "output_tokens": 0, "web_searches": 0, "cost_usd": 0.0,
         "unpriced_calls": 0}
    if USAGE_LOG.exists():
        for line in USAGE_LOG.read_text().splitlines():
            try:
                r = json.loads(line)
            except ValueError:
                continue
            if not str(r.get("ts", "")).startswith(month):
                continue
            s["calls"] += 1
            for k in ("input_tokens", "output_tokens", "web_searches"):
                s[k] += r.get(k) or 0
            if r.get("cost_usd") is None:
                s["unpriced_calls"] += 1
            else:
                s["cost_usd"] += r["cost_usd"]
    s["cost_usd"] = round(s["cost_usd"], 2)
    return s


# ------------------------------------------------------------------------------------------- parsing
def extract_json(text):
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip(), flags=re.S)
    a, b = t.find("{"), t.rfind("}")
    if a < 0 or b <= a:
        raise AIError("couldn't read the AI's answer (no JSON found): try again")
    try:
        return json.loads(t[a:b + 1])
    except ValueError:
        raise AIError("couldn't read the AI's answer (bad JSON): try again")


def _s(v):
    return "" if v is None else str(v).strip()


def _list(v):
    return [_s(x) for x in v if _s(x)] if isinstance(v, list) else ([_s(v)] if _s(v) else [])


# --------------------------------------------------------------------------------------------- check
CHECK_SYSTEM = """You are the fact-checker and copy editor for Trivia By Troy, a live bar trivia company. Teams hear each question read aloud and see it on a slide, have about a minute, and write one short answer. The host must be able to mark answers right or wrong quickly and fairly.

For each question, judge:
1. Accuracy: is the given answer correct?{verify}
2. Uniqueness: is it unambiguously the single best answer? Would another answer reasonably have to be marked correct (other names, spellings, nicknames, a tie, several people or things fitting the clue, "it depends")?
3. Wording: is the question clear and precise? Flag facts that change over time and need a date ("as of 2025"), vague superlatives, missing units, giveaways, double negatives, and anything a team could fairly argue with.
4. Typos and grammar.

Music and movie questions ("Name this song", "What movie is this scene from?") depend on a clip you can't hear or see: check only that the answer text is accurate and well formatted (title, artist or year).

Don't nitpick. If a question is fine, give it verdict "ok" with no issues. Suggest a rewrite only when it fixes a real problem, keep the host's voice and length, and never change what the question is about.

Reply with only a JSON object and no other text:
{{"results": [{{"id": "<question id>", "verdict": "ok" | "revise" | "wrong", "issues": ["<short, specific problem>"], "suggested_question": "<rewrite, or empty>", "suggested_answer": "<corrected answer, or empty>", "accept_also": ["<other answers the host should accept>"], "fun_fact": "<one sentence the host can share, or empty>", "sources": ["<URL you relied on>"], "confidence": "high" | "medium" | "low"}}]}}
Verdict "wrong" means the given answer is incorrect; "revise" means the answer is right but the question or answer should change."""


def check_questions(bank, qids, context=""):
    prov, o = provider(bank["settings"])
    qs = [(qid, bank["qs"][qid]) for qid in qids if qid in bank["qs"]]
    if not qs:
        raise AIError("no questions to check")
    verify = (" If you aren't certain, verify it with a web search." if o["web"] else
              " If you aren't certain, say so and lower your confidence.")
    lines = [f"Category: {context}" if context else "", "Check these trivia questions:", ""]
    for qid, q in qs:
        lines += [f"ID: {qid}", f"Question: {_s(q['question'])}", f"Answer: {_s(q['answer'])}"]
        if _s(q.get("media")):
            lines.append("(This question plays a song or movie clip.)")
        lines.append("")
    text, usage = prov.run(CHECK_SYSTEM.format(verify=verify), "\n".join(lines).strip(), o["web"])
    data = extract_json(text)
    by_id = {}
    for r in data.get("results") or []:
        rid = _s(r.get("id"))
        verdict = _s(r.get("verdict")).lower()
        by_id[rid] = {"id": rid, "verdict": verdict if verdict in ("ok", "revise", "wrong") else "revise",
                      "issues": _list(r.get("issues")), "suggested_question": _s(r.get("suggested_question")),
                      "suggested_answer": _s(r.get("suggested_answer")), "accept_also": _list(r.get("accept_also")),
                      "fun_fact": _s(r.get("fun_fact")), "sources": [u for u in _list(r.get("sources"))
                                                                     if u.startswith(("http://", "https://"))][:5],
                      "confidence": _s(r.get("confidence")).lower() or "medium"}
    results = []
    for qid, q in qs:
        r = by_id.get(qid) or {"id": qid, "verdict": "unchecked", "issues": ["The AI didn't return a result for this one."],
                               "suggested_question": "", "suggested_answer": "", "accept_also": [], "fun_fact": "",
                               "sources": [], "confidence": "low"}
        r.update(question=_s(q["question"]), answer=_s(q["answer"]))
        results.append(r)
    rec = log_usage("check", o["provider"], usage, len(qs))
    return {"results": results, "provider": prov.name, "model": usage["model"], "web": o["web"],
            "cost_usd": rec["cost_usd"], "web_searches": usage["web_searches"]}


# ------------------------------------------------------------------------------------------ generate
GEN_SYSTEM = """You write questions for Trivia By Troy, a live bar trivia company. Teams hear each question read aloud and see it on a slide, have about a minute, and write one short answer on paper.

Rules for every question:
- Exactly one short, unambiguous correct answer a team could write down (a name, word, number or short phrase). If a well-known alternative should also be accepted, list it.
- Fair and fun: no trick questions, nothing that needs a calculator, nothing offensive.
- Facts that won't go stale, or say "as of <year>". Be accurate.{verify}
- Match the length and voice of the examples. Vary the answer types.
- Don't repeat or closely paraphrase any of the existing questions listed.

- If a target audience is given, pick facts, people, songs, shows and events those generations grew up with or follow now, and phrase things so they land with them.
- Tag each question with the generations likely to know the answer: Boomer (born 1946-64), Gen X (1965-80), Millennial (1981-96), Gen Z (1997-2012), Gen Alpha (2013+).

Also write a category intro blurb in the same playful voice as the example blurbs (one or two sentences, usually ending with an ellipsis).

Reply with only a JSON object and no other text:
{{"blurb": "<category intro>", "questions": [{{"question": "...", "answer": "...", "accept_also": ["..."], "fun_fact": "<one sentence the host can share>", "difficulty": "easy" | "medium" | "hard", "picture_idea": "<a photo that would suit the slide without giving the answer away, or empty>", "generations": ["Boomer" | "Gen X" | "Millennial" | "Gen Z" | "Gen Alpha"], "sources": ["<URL you relied on>"]}}]}}"""


def generate(bank, topic, count=5, difficulty="mixed", notes="", ctype="standard", generations=None):
    prov, o = provider(bank["settings"])
    topic = _s(topic)
    if not topic:
        raise AIError("give a topic")
    count = max(1, min(int(count or 5), 10))
    plain = [q for q in bank["qs"].values() if not _s(q.get("media")) and not build.is_yes(q.get("retired"))]
    rnd = random.Random(topic)
    examples = rnd.sample(plain, min(12, len(plain)))
    blurbs = [_s(c.get("blurb")) for c in bank["cats"].values() if _s(c.get("blurb")) and _s(c.get("type")) in ("", "standard")]
    words = {w for w in re.findall(r"[a-z]{4,}", topic.lower())}
    related = [q for q in bank["qs"].values()
               if words & set(re.findall(r"[a-z]{4,}", f"{_s(q['question'])} {_s(q['answer'])}".lower()))][:40]
    verify = " Verify facts with a web search." if o["web"] else ""
    parts = [f"Topic: {topic}", f"Number of questions: {count}",
             f"Difficulty: {difficulty if difficulty in ('easy', 'medium', 'hard') else 'a mix of easy, medium and hard'}"]
    if ctype == "music":
        parts.append("This is a Music category: each question is 'Name this song. Who performs it?' with a song clip; "
                     "suggest songs that fit the topic, with the answer as 'Song Title – Artist'.")
    elif ctype == "movie":
        parts.append("This is a Movies category: each question is 'What movie is this scene from?' with a clip; "
                     "suggest movies that fit the topic, with the answer as 'Title (year)'.")
    if build.parse_gens(generations):
        parts.append(f"Target audience: {gen_line(generations)}")
    if _s(notes):
        parts.append(f"Notes from the host: {_s(notes)}")
    parts += ["", "Example questions in our style:"] + [f"- Q: {_s(q['question'])}  A: {_s(q['answer'])}" for q in examples]
    parts += ["", "Example category blurbs:"] + [f"- {b}" for b in rnd.sample(blurbs, min(4, len(blurbs)))]
    if related:
        parts += ["", "Existing questions on similar subjects (don't repeat these):"] + [
            f"- {_s(q['question'])} ({_s(q['answer'])})" for q in related]
    text, usage = prov.run(GEN_SYSTEM.format(verify=verify), "\n".join(parts), o["web"])
    data = extract_json(text)
    qs = []
    for q in data.get("questions") or []:
        if not _s(q.get("question")) or not _s(q.get("answer")):
            continue
        d = _s(q.get("difficulty")).lower()
        qs.append({"question": _s(q["question"]), "answer": _s(q["answer"]), "accept_also": _list(q.get("accept_also")),
                   "fun_fact": _s(q.get("fun_fact")), "difficulty": d if d in ("easy", "medium", "hard") else "",
                   "picture_idea": _s(q.get("picture_idea")), "generations": build.parse_gens(q.get("generations")),
                   "sources": [u for u in _list(q.get("sources")) if u.startswith(("http://", "https://"))][:5]})
    if not qs:
        raise AIError("the AI didn't return any usable questions: try again or reword the topic")
    rec = log_usage("generate", o["provider"], usage, len(qs))
    return {"blurb": _s(data.get("blurb")), "questions": qs, "provider": prov.name, "model": usage["model"],
            "web": o["web"], "cost_usd": rec["cost_usd"], "web_searches": usage["web_searches"]}


# ------------------------------------------------------------------------------- generations tagging
TAG_SYSTEM = """You help a bar trivia company balance its games across age groups. For each trivia question, decide which generations of American adults (and kids, for family nights) are likely to know the answer or get the reference:
Boomer (born 1946-64), Gen X (1965-80), Millennial (1981-96), Gen Z (1997-2012), Gen Alpha (2013+).

Tag a generation when a typical member would plausibly know it from their own life and culture. Evergreen general knowledge (science, geography, holidays, long-running classics) gets every generation old enough to know it. Be realistic: an 80s hit is Boomer/Gen X/older Millennial; a 2020s TikTok trend is Gen Z/Gen Alpha. Most questions get 1-4 tags.

Reply with only a JSON object and no other text:
{"results": [{"id": "<question id>", "generations": ["Boomer" | "Gen X" | "Millennial" | "Gen Z" | "Gen Alpha"], "why": "<a few words>"}]}"""


def tag_generations(bank, qids):
    """Suggests generation tags. Cheap: low effort, no web search."""
    prov, o = provider(bank["settings"], effort="low")
    qs = [(qid, bank["qs"][qid]) for qid in qids if qid in bank["qs"]]
    if not qs:
        raise AIError("no questions to tag")
    lines = ["Tag these trivia questions:", ""]
    for qid, q in qs:
        lines += [f"ID: {qid}", f"Question: {_s(q['question'])}", f"Answer: {_s(q['answer'])}", ""]
    text, usage = prov.run(TAG_SYSTEM, "\n".join(lines).strip(), False)
    data = extract_json(text)
    got = {_s(r.get("id")): r for r in data.get("results") or []}
    results = []
    for qid, q in qs:
        r = got.get(qid, {})
        results.append({"id": qid, "question": _s(q["question"]), "answer": _s(q["answer"]),
                        "current": build.parse_gens(q.get("generations")), "suggested": build.parse_gens(r.get("generations")),
                        "why": _s(r.get("why")), "missing": qid not in got})
    rec = log_usage("tag", o["provider"], usage, len(qs))
    return {"results": results, "provider": prov.name, "model": usage["model"], "cost_usd": rec["cost_usd"]}


# ------------------------------------------------------------------------------------ category ideas
IDEAS_SYSTEM = """You brainstorm category ideas for Trivia By Troy, a live bar trivia company. A category is a themed group of 5 questions with a short title and a playful intro blurb read by the host. Games have 9 categories, including one Music category (name the song) and sometimes a Movies category (name the movie from a clip).

Give fresh, varied, specific ideas (not just "History" or "Sports"): clever angles, wordplay, seasonal hooks, crossovers between topics. Don't repeat or closely resemble the existing categories listed. Don't write the questions; just the idea.

Reply with only a JSON object and no other text:
{"ideas": [{"title": "<short, catchy>", "blurb": "<one or two sentences in a playful host voice, usually ending with an ellipsis>", "type": "standard" | "music" | "movie", "angle": "<one line on what the 5 questions would cover>", "generations": ["Boomer" | "Gen X" | "Millennial" | "Gen Z" | "Gen Alpha"], "season": "<month or holiday it suits, or empty>"}]}"""


def category_ideas(bank, theme="", count=10, generations=None, notes=""):
    prov, o = provider(bank["settings"])
    count = max(3, min(int(count or 10), 20))
    existing = sorted({_s(c.get("title")) for c in bank["cats"].values() if _s(c.get("title"))})
    blurbs = [_s(c.get("blurb")) for c in bank["cats"].values() if _s(c.get("blurb")) and _s(c.get("type")) in ("", "standard")]
    parts = [f"Number of ideas: {count}", f"Theme or occasion: {_s(theme) or 'anything (a varied mix)'}",
             f"Today's date: {dt.date.today():%B %d, %Y}"]
    if build.parse_gens(generations):
        parts.append(f"Target audience: {gen_line(generations)}. Make the ideas land with them.")
    if _s(notes):
        parts.append(f"Notes from the host: {_s(notes)}")
    parts += ["", "Example blurbs in our voice:"] + [f"- {b}" for b in random.Random(theme).sample(blurbs, min(5, len(blurbs)))]
    parts += ["", "Existing categories (don't repeat these):"] + [f"- {t}" for t in existing]
    text, usage = prov.run(IDEAS_SYSTEM, "\n".join(parts), False)
    data = extract_json(text)
    ideas = []
    taken = {t.lower() for t in existing}
    for i in data.get("ideas") or []:
        title = _s(i.get("title"))
        if not title or title.lower() in taken:
            continue
        t = _s(i.get("type")).lower()
        ideas.append({"title": title, "blurb": _s(i.get("blurb")), "type": t if t in ("standard", "music", "movie") else "standard",
                      "angle": _s(i.get("angle")), "generations": build.parse_gens(i.get("generations")),
                      "season": _s(i.get("season"))})
    if not ideas:
        raise AIError("the AI didn't return any usable ideas: try again")
    rec = log_usage("ideas", o["provider"], usage, len(ideas))
    return {"ideas": ideas, "provider": prov.name, "model": usage["model"], "cost_usd": rec["cost_usd"]}


def list_models(settings, which=None):
    s = dict(settings)
    if which:
        s["ai_provider"] = which
        s["ai_model"] = ""
    prov, _ = provider(s)
    return prov.list_models()
