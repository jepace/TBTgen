"""
At-home video: turns one set into an MP4 (no PowerPoint needed).

  python3 build.py --video 2026-W48-A [--video-seconds 30] [--video-answer-seconds 6]
                   [--video-height 720|1080] [--video-no-media]

Slides are drawn with Pillow in the same style as the decks (black background, Trebuchet / Savoye /
Brush Script when installed, gold answers, the Trivia By Troy logo filling up as the countdown), then
each slide becomes a short clip with ffmpeg and the clips are joined. Music and movie clips play on
their question slides. Progress is printed as "video: N/M" lines (the app shows them as a progress bar).

At-home pacing: a shorter timer than the bar (default 30 s), answers after each round, no venue
branding, no tiebreaker.
"""
import json
import os
import re
import shutil
import subprocess
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

import build

ASSETS = build.ASSETS
FPS = 25
GOLD = (245, 194, 66)
WHITE = (255, 255, 255)
GREY = (150, 150, 150)
LOGO = ASSETS / "brand/tbt_logo.png"

# ------------------------------------------------------------------------------------------ fonts
FONT_DIRS = [Path.home() / "Library/Fonts", Path("/Library/Fonts"), Path("/System/Library/Fonts"),
             Path("/System/Library/Fonts/Supplemental"), Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
             Path("C:/Windows/Fonts")]
FONT_FILES = {  # first match wins; patterns are case-insensitive file-name globs
    "body": ["Trebuchet MS.ttf", "trebuc.ttf", "DejaVuSans.ttf", "LiberationSans-Regular.ttf"],
    "bold": ["Trebuchet MS Bold.ttf", "trebucbd.ttf", "DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf"],
    "italic": ["Trebuchet MS Italic.ttf", "trebucit.ttf", "DejaVuSans-Oblique.ttf", "LiberationSans-Italic.ttf"],
    "bolditalic": ["Trebuchet MS Bold Italic.ttf", "trebucbi.ttf", "DejaVuSans-BoldOblique.ttf",
                   "LiberationSans-BoldItalic.ttf"],
    "script": ["SavoyeLET.ttc", "Savoye LET.ttc", "*avoye*", "LiberationSerif-Italic.ttf", "DejaVuSerif-Italic.ttf"],
    "brush": ["Brush Script.ttf", "BRUSHSCI.TTF", "*rush*cript*", "LiberationSerif-BoldItalic.ttf",
              "DejaVuSerif-BoldItalic.ttf"],
}
_found = {}
_index = None


def _font_index():
    """All font files in the usual folders, by lower-case file name (built once)."""
    global _index
    if _index is None:
        _index = {}
        for d in FONT_DIRS:
            if d.exists():
                for p in d.rglob("*"):
                    if p.suffix.lower() in (".ttf", ".ttc", ".otf") and p.name.lower() not in _index:
                        _index[p.name.lower()] = p
    return _index


def font_path(kind):
    if kind not in _found:
        import fnmatch
        idx = _font_index()
        _found[kind] = next((idx[n] for pat in FONT_FILES[kind] for n in sorted(idx)
                             if fnmatch.fnmatch(n, pat.lower())), None)
    return _found[kind]


def fonts_report():
    """Which font file is used for each role (shown on the app's Status tab)."""
    return {k: (str(font_path(k)) if font_path(k) else "Pillow built-in") for k in FONT_FILES}


class Canvas:
    """A 16:9 slide drawn in inch coordinates, like the decks (13.333 x 7.5 in)."""

    def __init__(self, height):
        self.H = height
        self.W = round(height * 16 / 9)
        self.px_per_in = height / 7.5
        self.img = Image.new("RGB", (self.W, self.H), "black")
        self.d = ImageDraw.Draw(self.img)

    def px(self, inches):
        return int(round(inches * self.px_per_in))

    def font(self, kind, pt):
        size = max(8, int(round(pt * self.H / 540)))
        p = font_path(kind)
        return ImageFont.truetype(str(p), size) if p else ImageFont.load_default(size=size)

    def wrap(self, text, font, width_px):
        lines = []
        for para in str(text).split("\n"):
            words, line = para.split(), ""
            if not words:
                lines.append("")
                continue
            for w in words:
                trial = (line + " " + w).strip()
                if self.d.textlength(trial, font=font) <= width_px or not line:
                    line = trial
                else:
                    lines.append(line)
                    line = w
            lines.append(line)
        return lines

    STYLE_KIND = {(False, False): "body", (True, False): "bold", (False, True): "italic", (True, True): "bolditalic"}

    def rich(self, x, y, w, text, kind="body", pt=20, color=WHITE, spacing=1.2):
        """Like text() (left-aligned) but honours *italic* / **bold** / [red]..[/red] markup (see build.markup_runs);
        returns the bottom (inches)."""
        base_b, base_i = {"bold": (True, False), "bolditalic": (True, True), "italic": (False, True)}.get(kind, (False, False))
        fonts = {}

        def fnt(b, i):
            k = self.STYLE_KIND[(b, i)]
            if k not in fonts:
                fonts[k] = self.font(k, pt)
            return fonts[k]
        width = self.px(w)
        out = []  # lines of [(text, font, fill)]
        for runs in build.markup_runs(text, base_b, base_i):
            words, cur = [], []   # words = [[(piece, font, fill), ...]]
            for t, b, i, col in runs:
                fill = build.MARK_COLORS[col] if col else color
                for piece in re.findall(r"\s+|\S+", t):
                    if piece.isspace():
                        if cur:
                            words.append(cur)
                            cur = []
                    else:
                        cur.append((piece, fnt(b, i), fill))
            if cur:
                words.append(cur)
            line, lw = [], 0
            for grp in words:
                gw = sum(self.d.textlength(p, font=f) for p, f, _ in grp)
                sp = self.d.textlength(" ", font=grp[0][1]) if line else 0
                if line and lw + sp + gw > width:
                    out.append(line)
                    line, lw, sp = [], 0, 0
                if line:
                    line.append((" ", grp[0][1], grp[0][2]))
                line += grp
                lw += sp + gw
            out.append(line)
        lh = fnt(base_b, base_i).size * spacing
        top = self.px(y)
        for n, line in enumerate(out):
            lx = self.px(x)
            for p, f, fill in line:
                self.d.text((lx, top + n * lh), p, font=f, fill=fill)
                lx += self.d.textlength(p, font=f)
        return (top + lh * len(out)) / self.px_per_in

    def text(self, x, y, w, text, kind="body", pt=20, color=WHITE, align="left", h=None, valign="top",
             spacing=1.2):
        """Draws wrapped text in the box starting at (x, y) inches with width w; returns the bottom (inches)."""
        f = self.font(kind, pt)
        lines = self.wrap(text, f, self.px(w))
        lh = f.size * spacing
        total = lh * len(lines)
        top = self.px(y)
        if h is not None and valign != "top":
            room = self.px(h) - total
            top += room / 2 if valign == "middle" else room
        for i, line in enumerate(lines):
            lw = self.d.textlength(line, font=f)
            lx = self.px(x) + {"left": 0, "center": (self.px(w) - lw) / 2, "right": self.px(w) - lw}[align]
            self.d.text((lx, top + i * lh), line, font=f, fill=color)
        return (top + total) / self.px_per_in

    def picture(self, path, x, y, w, h):
        """Pastes an image fitted inside the box; returns its pixel box."""
        with Image.open(path) as im:
            im = im.convert("RGBA")
            bw, bh = self.px(w), self.px(h)
            scale = min(bw / im.width, bh / im.height)
            im = im.resize((max(1, int(im.width * scale)), max(1, int(im.height * scale))), Image.LANCZOS)
            ox, oy = self.px(x) + (bw - im.width) // 2, self.px(y) + (bh - im.height) // 2
            self.img.paste(im, (ox, oy), im)
            return ox, oy, im.width, im.height

    def logo(self):
        self.picture(LOGO, 11.85, 0.12, 1.3, 1.3)

    def save(self, path):
        self.img.save(path)
        return path


def logo_box_px(c):
    return c.px(11.85), c.px(0.12), c.px(1.3), c.px(1.3)


# ----------------------------------------------------------------------------------------- slides
def slide_title(c, set_id):
    c.picture(LOGO, 4.67, 0.5, 4.0, 3.6)
    c.text(0, 4.2, 13.333, "Trivia Night", "script", 110, align="center")
    c.text(0, 5.75, 13.333, "At-Home Edition", "bolditalic", 28, GOLD, align="center")
    c.text(0, 6.6, 13.333, "Grab a pen and paper!", "body", 20, GREY, align="center")


def slide_howto(c, seconds, has_bonus):
    c.logo()
    c.text(0.35, 0.2, 11, "How to Play", "body", 28)
    items = ["3 rounds of questions, in categories of 5.",
             f"You have {seconds} seconds per question. Watch the logo fill up! Pause any time.",
             "Write down your answers, then check them when the answers come up after each round.",
             "One point per question."]
    if has_bonus:
        items.append("Tonight's bonus category is worth double points!")
    items.append("Have fun!")
    y = 1.5
    for it in items:
        y = c.text(0.8, y, 11.5, "•  " + it, "bold", 24) + 0.25


def slide_script(c, script, sub="", head="Trivia Night"):
    c.logo()
    c.text(0.35, 0.2, 11, head, "body", 28)
    c.text(0, 2.2, 13.333, script, "script", 96, align="center", h=1.8, valign="middle")
    if sub:
        c.text(0, 4.6, 13.333, sub, "bolditalic", 24, align="center")


def slide_category(c, cat, first_q, nq, bonus):
    c.logo()
    c.text(0.35, 0.2, 11, f"Questions {first_q} - {first_q + nq - 1}", "body", 28)
    c.text(1.2, 2.0, 10.9, cat["title"], "script", 80, align="center", h=1.8, valign="middle")
    if bonus:
        c.text(0.5, 1.4, 4, "Tonight’s\nBonus Category", "brush", 28, GOLD)
        c.text(9.6, 3.9, 3.4, "Double Points!", "brush", 32, GOLD)
    c.text(0.5, 4.8, 11.5, cat["blurb"], "bold", 16, h=1.9, valign="bottom")
    c.text(0.5, 6.75, 6, "Good Luck!", "bolditalic", 20)


def question_top(c, num, question, note=""):
    c.text(0.35, 0.2, 11, f"Question {num}", "body", 28)
    bottom = c.rich(0.35, 0.85, 11.3, question, "bold", 20)
    if str(note or "").strip():
        bottom = c.rich(0.35, bottom + 0.02, 11.3, str(note).strip(), "bolditalic", 16)
    return max(2.25, bottom + 0.15)


def slide_question(c, qq, num, kind, logo_static=False):
    """Returns the media box (inches) for a movie, else None. The countdown logo is added by ffmpeg."""
    top = question_top(c, num, qq["question"], qq.get("question_note", ""))
    box = (2.4, top, 8.5, 7.2 - top)
    if logo_static:
        c.logo()
    else:
        c.picture(dim_logo(), 11.85, 0.12, 1.3, 1.3)
    if qq["media"] and kind_of(qq, kind) == "movie":
        return box
    if qq["media"]:
        c.picture(build.music_frame(), *box)
    elif qq["question_image"]:
        c.picture(ASSETS / qq["question_image"], *box)
    return None


def slide_answer(c, qq, num, kind, revealed):
    top = question_top(c, num, qq["question"], qq.get("question_note", ""))
    c.logo()
    right = (5.6, top, 7.4, 7.2 - top)
    if qq["question_image"] and not qq["media"]:
        c.picture(ASSETS / qq["question_image"], *right)
    elif qq["media"] and kind_of(qq, kind) == "music":
        c.picture(build.music_frame(), *right)
    if not revealed:
        return
    below = c.rich(0.5, top + 0.15, 4.9, qq["answer"], "bold", 26, GOLD)
    img_top = top + 1.6
    if str(qq.get("answer_note", "") or "").strip():
        below = c.rich(0.5, below + 0.1, 4.9, str(qq["answer_note"]).strip(), "bolditalic", 16)
        img_top = max(img_top, below + 0.2)
    if qq["answer_image"]:
        img_box = (0.5, img_top, 4.9, 7.2 - img_top) if (qq["question_image"] and not qq["media"]) else right
        if img_box == right:  # cover the question picture
            c.d.rectangle([c.px(right[0]), c.px(right[1]), c.px(right[0] + right[2]), c.px(right[1] + right[3])],
                          fill="black")
        c.picture(ASSETS / qq["answer_image"], *img_box)


def slide_pencils(c, rnd):
    c.logo()
    c.text(0.35, 0.2, 11, "Trivia Night", "body", 28)
    c.text(0, 1.9, 13.333, "Pencils Down!", "script", 96, align="center", h=1.8, valign="middle")
    c.text(0, 4.3, 13.333, f"Here come the answers for Round {build.NUM_WORDS[rnd]}…", "bolditalic", 24,
           align="center")


def slide_end(c, settings):
    c.picture(LOGO, 4.92, 0.4, 3.5, 3.2)
    c.text(0, 3.75, 13.333, "Thanks for playing!", "script", 96, align="center")
    c.text(0, 5.4, 13.333, "Add up your points. How did you do?", "bolditalic", 24, align="center")
    social = "   ·   ".join(x for x in [f"Instagram: {settings.get('instagram', '')}",
                                         f"Facebook: {settings.get('facebook', '')}"] if not x.endswith(": "))
    c.text(0, 6.4, 13.333, social, "body", 20, GREY, align="center")


def kind_of(qq, cat_kind):
    if not qq["media"]:
        return "none"
    if cat_kind == "music" or Path(qq["media"]).suffix.lower() in (".mp3", ".m4a", ".wav", ".aac"):
        return "music"
    return "movie"


_dim = None


def dim_logo():
    global _dim
    if _dim is None:
        _dim = build.dim_logo()
    return _dim


def logo_fill_frames(c, folder, steps=60):
    """Frames of the countdown: grey logo filling with colour from the bottom."""
    _, _, w, h = logo_box_px(c)
    color = Image.open(LOGO).convert("RGBA")
    grey = Image.open(dim_logo()).convert("RGBA")
    scale = min(w / color.width, h / color.height)
    size = (max(1, int(color.width * scale)), max(1, int(color.height * scale)))
    color, grey = color.resize(size, Image.LANCZOS), grey.resize(size, Image.LANCZOS)
    for i in range(steps + 1):
        frame = grey.copy()
        cut = int(size[1] * (1 - i / steps))
        frame.paste(color.crop((0, cut, size[0], size[1])), (0, cut), color.crop((0, cut, size[0], size[1])))
        frame.save(folder / f"fill_{i:03d}.png")
    ox = c.px(11.85) + (w - size[0]) // 2
    oy = c.px(0.12) + (h - size[1]) // 2
    return ox, oy, steps + 1


# ------------------------------------------------------------------------------------------ ffmpeg
ENC_V = ["-c:v", "libx264", "-preset", "veryfast", "-crf", "22", "-pix_fmt", "yuv420p", "-r", str(FPS)]
ENC_A = ["-c:a", "aac", "-b:a", "160k", "-ar", "48000", "-ac", "2"]


def run_ffmpeg(args):
    p = subprocess.run(["ffmpeg", "-v", "error", "-y", *args], capture_output=True, text=True)
    if p.returncode:
        raise RuntimeError("ffmpeg failed: " + p.stderr.strip()[-600:])


def seg_still(img, dur, out):
    run_ffmpeg(["-loop", "1", "-framerate", str(FPS), "-t", f"{dur:.3f}", "-i", str(img),
                "-f", "lavfi", "-t", f"{dur:.3f}", "-i", "anullsrc=r=48000:cl=stereo",
                "-shortest", *ENC_V, "-tune", "stillimage", *ENC_A, str(out)])


def seg_countdown(img, dur, out, fill, media=None, media_kind="none", box_px=None, clip=(None, None)):
    """A question slide: background + filling logo; optional music (audio) or movie (overlaid in box)."""
    folder, ox, oy, nframes = fill
    args = ["-loop", "1", "-framerate", str(FPS), "-t", f"{dur:.3f}", "-i", str(img),
            "-framerate", f"{nframes}/{dur:.3f}", "-i", str(folder / "fill_%03d.png")]
    trim = (["-ss", str(clip[0])] if clip[0] not in (None, "") else []) + \
           (["-to", str(clip[1])] if clip[1] not in (None, "") else [])
    if media:
        args += [*trim, "-i", str(media)]
    else:
        args += ["-f", "lavfi", "-t", f"{dur:.3f}", "-i", "anullsrc=r=48000:cl=stereo"]
    fc = f"[0:v][1:v]overlay={ox}:{oy}:eof_action=repeat[v1]"
    if media and media_kind == "movie":
        bx, by, bw, bh = box_px
        fc += (f";[2:v]scale={bw}:{bh}:force_original_aspect_ratio=decrease,setsar=1[m];"
               f"[v1][m]overlay=x={bx}+({bw}-overlay_w)/2:y={by}+({bh}-overlay_h)/2:eof_action=repeat[v]")
    else:
        fc += ";[v1]null[v]"
    if media and has_audio(media):
        fc += f";[2:a]aresample=48000,apad,atrim=0:{dur:.3f}[a]"
    elif media:
        args += ["-f", "lavfi", "-t", f"{dur:.3f}", "-i", "anullsrc=r=48000:cl=stereo"]
        fc += ";[3:a]anull[a]"
    else:
        fc += ";[2:a]anull[a]"
    run_ffmpeg([*args, "-filter_complex", fc, "-map", "[v]", "-map", "[a]", "-t", f"{dur:.3f}",
                *ENC_V, *ENC_A, str(out)])


def has_audio(path):
    p = subprocess.run(["ffprobe", "-v", "error", "-select_streams", "a", "-show_entries", "stream=index",
                        "-of", "csv=p=0", str(path)], capture_output=True, text=True)
    return bool(p.stdout.strip())


def media_seconds(path, clip):
    full = build.ffprobe_dur(path)
    start = float(clip[0]) if clip[0] not in (None, "") else 0.0
    end = float(clip[1]) if clip[1] not in (None, "") else full
    return max(0.5, min(end, full) - start)


# ------------------------------------------------------------------------------------------- plan
def make_video(bank, set_id, outdir, seconds=30, answer_seconds=6, height=720, media=True, workers=None):
    if not shutil.which("ffmpeg"):
        raise RuntimeError("ffmpeg is not installed (on a Mac: brew install ffmpeg)")
    settings = bank["settings"]
    media_extra = float(settings.get("video_media_extra_seconds", 10) or 10)
    slots = sorted(bank["sets"][set_id], key=lambda r: (int(r["round"]), int(r["slot"])))
    has_bonus = any(build.is_yes(r["bonus"]) for r in slots)
    outdir = Path(outdir) / "video"
    outdir.mkdir(parents=True, exist_ok=True)
    final = outdir / f"{set_id} at-home.mp4"
    tmp = Path(tempfile.mkdtemp(prefix="tbtvideo_", dir=build.CACHE if build.CACHE.exists() else None))
    jobs = []  # (kind, args...) in play order

    def new_canvas():
        return Canvas(height)

    def still(name, draw, dur):
        c = new_canvas()
        draw(c)
        jobs.append(("still", c.save(tmp / f"{len(jobs):04d}_{name}.png"), dur))

    still("title", lambda c: slide_title(c, set_id), 6)
    still("howto", lambda c: slide_howto(c, seconds, has_bonus), 12)

    fill_cache = {}
    skipped = []
    for rnd in (1, 2, 3):
        rslots = [r for r in slots if int(r["round"]) == rnd]
        still(f"round{rnd}", lambda c, r=rnd: slide_script(c, f"Round {build.NUM_WORDS[r]}", build.ROUND_SUB[r]), 5)
        played = []
        num = 1
        for sl in rslots:
            cat = bank["cats"][sl["category_id"]]
            qs = build.cat_questions(bank, cat["category_id"])
            if not media and any(q["media"] for q in qs):
                skipped.append(cat["title"])
                continue
            bonus = build.is_yes(sl["bonus"])
            still("cat", lambda c, cat=cat, n=num, k=len(qs), b=bonus: slide_category(c, cat, n, k, b), 10)
            for qq in qs:
                kind = kind_of(qq, cat["type"])
                c = new_canvas()
                box = slide_question(c, qq, num, cat["type"])
                img = c.save(tmp / f"{len(jobs):04d}_q{num}.png")
                clip = (qq.get("clip_start"), qq.get("clip_end"))
                mpath = ASSETS / qq["media"] if qq["media"] else None
                dur = float(qq.get("video_seconds") or 0) or seconds
                if mpath:
                    dur = max(dur, media_seconds(mpath, clip) + media_extra)
                key = round(dur, 3)
                if key not in fill_cache:
                    folder = tmp / f"fill_{len(fill_cache)}"
                    folder.mkdir()
                    fill_cache[key] = (folder, *logo_fill_frames(c, folder))
                box_px = (c.px(box[0]), c.px(box[1]), c.px(box[2]), c.px(box[3])) if box else None
                jobs.append(("count", img, dur, fill_cache[key], mpath, kind, box_px, clip))
                played.append((qq, num, cat["type"]))
                num += 1
        still(f"pencils{rnd}", lambda c, r=rnd: slide_pencils(c, r), 5)
        for qq, n, ctype in played:
            for revealed, dur in ((False, 2.5), (True, answer_seconds)):
                c = new_canvas()
                slide_answer(c, qq, n, ctype, revealed)
                jobs.append(("still", c.save(tmp / f"{len(jobs):04d}_a{n}{'r' if revealed else ''}.png"), dur))
    still("end", lambda c: slide_end(c, settings), 10)

    if skipped:
        print("NOTE: left out (music/movie): " + ", ".join(skipped))
    total = len(jobs)
    print(f"video: 0/{total}", flush=True)
    done = [0]

    def encode(i_job):
        i, job = i_job
        out = tmp / f"seg_{i:04d}.mp4"
        if job[0] == "still":
            seg_still(job[1], job[2], out)
        else:
            _, img, dur, fill, mpath, kind, box_px, clip = job
            seg_countdown(img, dur, out, fill, mpath, kind, box_px, clip)
        done[0] += 1
        print(f"video: {done[0]}/{total}", flush=True)
        return out

    with ThreadPoolExecutor(max_workers=workers or max(1, (os.cpu_count() or 2) - 1)) as ex:
        segs = list(ex.map(encode, enumerate(jobs)))
    listing = tmp / "list.txt"
    listing.write_text("".join(f"file '{p.as_posix()}'\n" for p in segs))
    run_ffmpeg(["-f", "concat", "-safe", "0", "-i", str(listing), "-c", "copy", "-movflags", "+faststart",
                str(final)])
    length = build.ffprobe_dur(final)
    # keep a poster image for the app, then clean up
    shutil.copy(next(p for p in sorted(tmp.glob("0000_*.png"))), outdir / f"{set_id} at-home.png")
    (outdir / f"{set_id} at-home.json").write_text(json.dumps(
        {"set_id": set_id, "seconds": seconds, "answer_seconds": answer_seconds, "height": height,
         "media": media, "length_seconds": round(length), "skipped": skipped}))
    shutil.rmtree(tmp, ignore_errors=True)
    return final, length
