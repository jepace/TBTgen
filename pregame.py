"""
The pregame deck (what plays while teams arrive) and the winners picture for social media.

  build.py --pregame VENUE_ID [--date YYYY-MM-DD]   -> out/pregame/<date> <venue> - pregame.pptx
                                                       out/pregame/<date> <venue> - winners.png (1080x1080)

Follows Troy's pregame deck: title, welcome, last week's winners, format, contact, "a quick note", rules,
social, season rules, this season's weekly winners, standings, upcoming nights, "it should go without saying".
The season slides come from the Scores data (Results/Seasons); the winners picture from the photo attached to last
week's winning team (Scores tab > Winners photo). The deck loops, advancing every `pregame_seconds` (Settings, 15).
"""
import datetime as dt
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont, ImageOps
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

import build
import scores
from build import ASSETS, GOLD, SCRIPT, W, header, picture, text, venue_logo

PREGAME = ASSETS / "brand/pregame"
WELCOME = ("We are excited for you to join us!\n\nGet a clipboard from the host for your team, fill out all the "
           "information in the yellow highlighted boxes, grab a drink, and get ready for a fun game!")


# ---------------------------------------------------------------------------------------------- data
def last_night(bank, venue_id, before):
    """The venue's latest night with results before `before`: (date, [(row, score)]) or (None, [])."""
    before = build.as_date(before)
    dates = sorted({build.as_date(r["date"]) for r in bank.get("Results", [])
                    if r.get("venue_id") == venue_id and r.get("date") and build.as_date(r["date"]) < before})
    if not dates:
        return None, []
    return dates[-1], scores.night_results(bank, venue_id, dates[-1])


def season_dates(bank, venue_id, st, date):
    """The season's game nights: those played, then the venue's scheduled nights, up to the season length."""
    played = [build.as_date(w["date"]) for w in st["weeks"]]
    future = sorted({build.as_date(r["date"]) for r in bank.get("Schedule", [])
                     if r.get("venue_id") == venue_id and r.get("date") and build.as_date(r["date"]) >= build.as_date(date)
                     and (not played or build.as_date(r["date"]) > played[-1])})
    return (played + future)[:st["weeks_total"]], len(played)


# ---------------------------------------------------------------------------------------------- picture
def _font(kind, size):
    import video
    p = video.font_path(kind)
    return ImageFont.truetype(str(p), size) if p else ImageFont.load_default(size=size)


def winners_image(bank, venue_id, date, out, size=1080):
    """Square 'Trivia Winners!' picture: the winners' photo, Trivia By Troy logo, venue logo and team name
    (like Troy's). With no photo yet: the logos and the name on black."""
    night_date, night = last_night(bank, venue_id, build.as_date(date) + dt.timedelta(days=1))
    if not night:
        return None
    row = night[0][0]
    team = build.plain(scores.team_name(bank, str(row["team_id"])))
    photo = next((r.get("photo") for r, _ in night if r.get("photo")), None)
    canvas = Image.new("RGB", (size, size), "black")
    if photo and not build.asset_problem(photo):
        with Image.open(build.slide_image(build.asset(photo))) as im:
            im = ImageOps.exif_transpose(im).convert("RGB")
            canvas.paste(ImageOps.fit(im, (size, size), Image.LANCZOS), (0, 0))
    d = ImageDraw.Draw(canvas)
    s = size / 1080

    def outlined(xy, txt, font, anchor, fill="white"):
        d.text(xy, txt, font=font, fill=fill, anchor=anchor, stroke_width=max(2, int(7 * s)), stroke_fill="black")
    venue = bank["venues"].get(venue_id, {})
    tbt_box, venue_box = int(250 * s), int(270 * s)
    # Trivia By Troy logo (neon on black): blended so the black disappears and the glow stays
    with Image.open(build.slide_image(ASSETS / "brand/tbt_logo.png")) as lg:
        lg = lg.convert("RGB")
        lg = lg.crop((6, 6, lg.width - 6, lg.height - 6))   # the file has a thin light border line
        lg.thumbnail((tbt_box, tbt_box), Image.LANCZOS)
        x, y = size - lg.width - int(20 * s), int(15 * s)
        mask = lg.convert("L").point(lambda v: 0 if v < 45 else min(255, v * 3))   # neon opaque; black (and a stray line) gone
        canvas.paste(lg, (x, y), mask)
    # venue logo: the black around it is cut away
    with Image.open(build.slide_image(venue_logo(venue))) as lg:
        lg = lg.convert("RGBA")
        lg.thumbnail((venue_box, venue_box), Image.LANCZOS)
        r, g, b, a = lg.split()
        dark = Image.eval(Image.merge("RGB", (r, g, b)).convert("L"), lambda v: 0 if v < 14 else 255)
        lg.putalpha(ImageChops.multiply(a, dark.filter(ImageFilter.GaussianBlur(1))))
        canvas.paste(lg, (int(20 * s), size - lg.height - int(20 * s)), lg)
    title_f = _font("bold", int(110 * s))
    room = size - tbt_box - int(60 * s)
    while title_f.getlength("Trivia Winners!") > room and title_f.size > 30:
        title_f = _font("bold", title_f.size - 4)
    outlined((int(30 * s), int(30 * s)), "Trivia Winners!", title_f, "la")
    f = _font("bold", int(96 * s))
    while f.getlength(team) > size - venue_box - int(70 * s) and f.size > 30:
        f = _font("bold", f.size - 4)
    outlined((size - int(30 * s), size - int(30 * s)), team, f, "rd", fill=(245, 194, 66) if not photo else "white")
    if not photo:
        outlined((size // 2, size // 2), build.ordinal_date(night_date), _font("bold", int(60 * s)), "mm")
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    canvas.save(out)
    return Path(out)


# ---------------------------------------------------------------------------------------------- deck
def loop_show(prs):
    """Make the slideshow loop (presentation properties: showPr loop)."""
    from lxml import etree
    for part in prs.part.package.iter_parts():
        if str(part.partname).endswith("/presProps.xml"):
            root = etree.fromstring(part.blob)
            ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
            show = root.find(f"{{{ns}}}showPr")
            if show is None:
                show = etree.Element(f"{{{ns}}}showPr")
                root.insert(0, show)
            show.set("loop", "1")
            show.set("showNarration", "1")
            part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def build_pregame(bank, venue_id, date, outdir):
    venue = bank["venues"][venue_id]
    settings = bank["settings"]
    secs = float(settings.get("pregame_seconds") or 15)
    st = scores.season_standings(bank, venue_id, upto=build.as_date(date) - dt.timedelta(days=1))
    deck = build.Deck()
    logo = venue_logo(venue)
    stem = build.safe_name(f"{build.as_date(date)} {build.plain(venue.get('name', venue_id))}")
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    def slide(fn, *a):
        s = fn(*a) or deck.prs.slides[-1]
        build.auto_advance(s, secs)

    slide(build.slide_venue_title, deck, venue)

    def welcome():
        s = deck.slide()
        header(s, "Trivia Night", logo)
        text(s, 0, Inches(1.5), W, Inches(1.4), "Welcome to Trivia Night!", size=72, font=SCRIPT, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, color=GOLD)
        text(s, Inches(1.2), Inches(3.3), Inches(10.9), Inches(3.5), settings.get("pregame_welcome") or WELCOME, size=28,
             bold=True, align=PP_ALIGN.CENTER)
        return s
    slide(welcome)

    img = winners_image(bank, venue_id, build.as_date(date) - dt.timedelta(days=1), outdir / f"{stem} - winners.png")
    if img:
        def winners():
            s = deck.slide()
            text(s, Inches(0.3), Inches(0.4), Inches(2.6), Inches(6.7), "\n".join("LAST WEEK"), size=40, bold=True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, color=GOLD, spacing=0)
            picture(s, img, Inches(3.1), Inches(0.15), Inches(7.2), Inches(7.2))
            return s
        slide(winners)

    slide(build.slide_bullets, deck, venue, "Trivia Format", build.FORMAT_ITEMS, 24)

    def contact():
        s = deck.slide()
        header(s, "Trivia By Troy", logo)
        text(s, 0, Inches(2.0), W, Inches(1.6), "Questions?  Feedback?  Want to book a gig?\nContact us!", size=40, bold=True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, 0, Inches(4.2), W, Inches(1.0), settings.get("contact_email") or "triviabytroy@gmail.com", size=44,
             font=SCRIPT, align=PP_ALIGN.CENTER, color=GOLD)
        return s
    slide(contact)

    def full_image(name):
        def f():
            s = deck.slide()
            picture(s, PREGAME / name, 0, 0, W, build.H)
            return s
        return f
    if (PREGAME / "note_everyone.jpg").exists():
        slide(full_image("note_everyone.jpg"))
    slide(build.slide_bullets, deck, venue, "Trivia Rules", build.rules_items(venue), 22)

    def social():
        s = deck.slide()
        header(s, "Follow Us!", logo)
        text(s, Inches(0.8), Inches(1.8), Inches(11.7), Inches(1.6),
             "Follow us on social media to find out where we’ll be, what’s new, and game hints!", size=34, bold=True,
             align=PP_ALIGN.CENTER)
        picture(s, ASSETS / "brand/instagram.png", Inches(2.6), Inches(4.2), Inches(1.0), Inches(1.0))
        text(s, Inches(3.8), Inches(4.2), Inches(6), Inches(1.0), settings.get("instagram", ""), size=36, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
        picture(s, ASSETS / "brand/facebook.png", Inches(2.6), Inches(5.5), Inches(1.0), Inches(1.0))
        text(s, Inches(3.8), Inches(5.5), Inches(6), Inches(1.0), settings.get("facebook", ""), size=36, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
        return s
    slide(social)

    if st["season"] is not None:
        dates, n_played = season_dates(bank, venue_id, st, date)
        weeks, drop = st["weeks_total"], st["drop_lowest"]
        end = f" and ends {build.ordinal_date(dates[-1])}" if len(dates) == weeks else ""

        def season_rules():
            s = deck.slide()
            header(s, "Trivia Night Seasons", logo)
            text(s, Inches(0.6), Inches(1.3), Inches(12), Inches(5.9),
                 f"Team Points will accumulate each week throughout the {weeks} week Season.\n\n"
                 + (f"Each Team’s lowest {'score' if drop == 1 else f'{drop} scores'} will be dropped at the end of the season "
                    f"(Totals are based on best {weeks - drop} weeks for each team). " if drop else "")
                 + "The Team with the most Points at the end of the Season wins!\n\n"
                 "In the event of a tie, the team with the most total points (before dropping lowest scores) wins. "
                 "If there’s still a tie, team with more wins for the season takes it all.\n\n"
                 f"Season {st['season']} started {build.ordinal_date(st['start'])}{end}.", size=22, bold=True)
            return s
        slide(season_rules)

        def season_winners():
            s = deck.slide()
            header(s, f"Season {st['season']}", logo)
            text(s, Inches(9.0), Inches(0.25), Inches(2.6), Inches(0.7), f"{weeks} Weeks", size=24, bold=True, italic=True)
            winners = {w["date"]: build.plain(w["winner"]) for w in st["weeks"]}
            lines = [f"{build.ordinal_date(d)} Winners - {winners[str(d)]}" if str(d) in winners else build.ordinal_date(d)
                     for d in dates]
            size = 20 if len(lines) <= 13 else 16
            text(s, Inches(0.8), Inches(1.25), Inches(11.5), Inches(6), "\n".join(lines), size=size, bold=True, spacing=2)
            return s
        slide(season_winners)

        if st["rows"]:
            def standings():
                s = deck.slide()
                best = st["weeks_played"] - drop if st["weeks_played"] > 1 else st["weeks_played"]
                header(s, f"Season {st['season']} Standings", logo)
                text(s, Inches(0.35), Inches(0.85), Inches(11), Inches(0.5),
                     f"as of Week {st['weeks_played']} reflecting {best} best score{'' if best == 1 else 's'}", size=20, italic=True)
                lines = scores.standings_lines(st)
                cols = [lines[:13], lines[13:26]] if len(lines) > 13 else [lines]
                for i, col in enumerate(cols):
                    text(s, Inches(0.8 + 6.1 * i), Inches(1.5), Inches(6), Inches(5.8), "\n".join(col), size=22, bold=True, spacing=3)
                return s
            slide(standings)

    slide(build.slide_upcoming, deck, venue, settings)
    if (PREGAME / "note_not_endorsement.jpg").exists():
        slide(full_image("note_not_endorsement.jpg"))
    loop_show(deck.prs)
    path = outdir / f"{stem} - pregame.pptx"
    deck.prs.save(path)
    return path, img
