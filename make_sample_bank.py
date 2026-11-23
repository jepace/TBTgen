"""Creates trivia_bank.xlsx with sample data (run once; after that, edit the spreadsheet)."""
import datetime as dt

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

wb = Workbook()
HEAD = PatternFill("solid", fgColor="1F2937")
AUTOFILL = PatternFill("solid", fgColor="6B7280")
AUTO_ROWS = 500   # helper formulas are pre-filled this far down


def sheet(title, cols, rows, widths=None, auto=None, date_cols=()):
    """auto: list of (header, formula_template) helper columns appended on the right; '{r}' = row number."""
    auto = auto or []
    ws = wb.create_sheet(title)
    ws.append(cols + [h for h, _ in auto])
    for i, c in enumerate(ws[1]):
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = AUTOFILL if i >= len(cols) else HEAD
    for r in rows:
        ws.append(list(r))
    if auto:
        for r in range(2, AUTO_ROWS + 2):
            for j, (_, f) in enumerate(auto):
                c = ws.cell(r, len(cols) + 1 + j, f.format(r=r))
                c.font = Font(italic=True, color="6B7280")
                if "MAXIFS" in f:
                    c.number_format = "yyyy-mm-dd"
    for col in date_cols:
        for r in range(2, ws.max_row + 1):
            ws.cell(r, col).number_format = "yyyy-mm-dd"
    ws.freeze_panes = "A2"
    for i, w in enumerate(widths or [18] * (len(cols) + len(auto))):
        ws.column_dimensions[chr(65 + i)].width = w
    for row in ws.iter_rows(min_row=2, max_row=len(rows) + 1):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    return ws


def lookup(key_col, tab, out_col, key_tab_col="A"):
    return (f'=IF(${key_col}{{r}}="","",IFERROR(INDEX({tab}!${out_col}:${out_col},'
            f'MATCH(${key_col}{{r}},{tab}!${key_tab_col}:${key_tab_col},0)),"?"))')


wb.remove(wb.active)

# ---------------------------------------------------------------- Settings
sheet("Settings", ["key", "value", "meaning"], [
    ("default_seconds", 60, "Seconds a question slide stays up (and the logo countdown length)"),
    ("media_extra_seconds", 30, "For music/movie slides with no 'seconds' set: clip length + this many seconds"),
    ("repeat_lookback_days", 365, "Warn if a venue saw a category within this many days"),
    ("company_name", "Trivia By Troy", ""),
    ("instagram", "triviabytroy", ""),
    ("facebook", "Trivia By Troy", ""),
    ("calendar_id", "8c65afbf52eeb5b75e6502c813588834f034bd7b5f8933e6eb67709921e1c983@group.calendar.google.com",
     "The 'Trivia By Troy Schedule' Google Calendar (used by TBTgen > Pull schedule from calendar)"),
], [24, 18, 80])

# ---------------------------------------------------------------- Venues
sheet("Venues", ["venue_id", "name", "logo", "region", "night", "upcoming_nights", "active", "calendar_name", "audience"], [
    ("wine_press", "101 Wine Press", "venues/101_wine_press.jpg", "North", "Wed",
     "November 25th No Trivia - Happy Thanksgiving!\nDecember 2nd\nDecember 9th", "Y", "101 Wine Press", "Gen X, Millennial"),
    ("taproom", "Sample Taproom", "venues/sample_taproom.png", "South", "Tue", "December 1st\nDecember 8th", "Y",
     "Sample Taproom", "Millennial, Gen Z"),
    ("pub", "Sample Pub", "venues/sample_pub.png", "North", "Wed", "December 3rd\nDecember 10th", "Y", "Sample Pub",
     "Boomer, Gen X"),
], [14, 22, 30, 10, 8, 50, 8, 22, 22])

# ---------------------------------------------------------------- Categories
cats = [
    ("turkey", "Turkey", "The main dish spotlight…", "standard", "November"),
    ("tg_potluck", "Thanksgiving Potluck", "Just like your plate on thanksgiving, a little bit of everything…", "standard", "November"),
    ("tg_mishaps", "Thanksgiving Dinner Mishaps",
     "Sometimes it’s not about the “perfect” holiday meal, and more about who you share it with; "
     "Name the TV shows these infamous Thanksgiving meals were inspired by…", "standard", "November"),
    ("nov_bdays", "November Birthday Bash", "Famous people born in November…", "standard", "November"),
    ("nov_pop", "November in Pop Culture",
     "Falling leaves, rising stars, and the moments everyone was talking about…", "standard", "November"),
    ("music_tg", "Music",
     "Name these songs!\nWho performs them?\nWhen naming the artist please use the primary credited artist "
     "(usually whose album it is on), not the collaborators.\nOne point awarded for EACH!", "music", ""),
    ("holiday_travel", "There’s no Place Like Home for the Holidays",
     "I don’t care how we get there but please just get me home… Some classic movies based on holiday travel…",
     "standard", "November-December"),
    ("potluck_1", "Potluck", "A little bit of everything…", "standard", ""),
    ("movies_tg", "Movies",
     "Name the movies these scenes are from.\nAnswer as complete as possible; sequels, prequels, etc. should be "
     "the full name it is in the series. Example: Don’t just write “Star Wars” if it’s “The Empire Strikes Back”.",
     "movie", ""),
]

# ---------------------------------------------------------------- Questions
Q = []


def q(cat, n, question, answer, qimg="", aimg="", media="", start="", end="", secs="", notes=""):
    Q.append((cat, n, question, answer, qimg, aimg, media, start, end, secs, notes))


q("turkey", 1, "What do you call the process of pouring or squeezing juices over a turkey while it's cooking to help keep it moist?",
  "Basting", "images/turkey.jpg", "images/basting.jpg")
q("turkey", 2, "What is the name for an adult male turkey?", "A tom (or gobbler)")
q("turkey", 3, "What is the name of the fleshy flap that hangs over a turkey's beak?", "Snood")
q("turkey", 4, "Which Founding Father wrote that the turkey was “a much more respectable Bird” than the bald eagle?", "Benjamin Franklin")
q("turkey", 5, "Which U.S. state raises the most turkeys?", "Minnesota")

q("tg_potluck", 1, "Which state produces the most cranberries?", "Wisconsin", "images/cranberries.jpg", "images/wisconsin_map.jpg")
q("tg_potluck", 2, "Which president proclaimed Thanksgiving a national holiday in 1863?", "Abraham Lincoln")
q("tg_potluck", 3, "In what year was the first Macy's Thanksgiving Day Parade held?", "1924")
q("tg_potluck", 4, "Which NFL team has hosted a Thanksgiving Day game almost every year since 1934?", "Detroit Lions")
q("tg_potluck", 5, "Canadian Thanksgiving falls on which Monday of October?", "The second Monday")

q("tg_mishaps", 1, "What TV show does this Thanksgiving meal represent?", "A Charlie Brown Thanksgiving",
  "images/snack_dinner.jpg", "images/charlie_brown.jpg")
q("tg_mishaps", 2, "Name this T.V. show from the Thanksgiving scene below.", "Gossip Girl", "images/gossip_girl_scene.jpg")
q("tg_mishaps", 3, "In which sitcom does a character get a whole turkey stuck on her head?", "Friends")
q("tg_mishaps", 4, "“As God is my witness, I thought turkeys could fly.” Which sitcom's Thanksgiving episode?", "WKRP in Cincinnati")
q("tg_mishaps", 5, "Which sitcom had a Thanksgiving episode called “Slapsgiving”?", "How I Met Your Mother")

q("nov_bdays", 1, "This actor, born November 11, 1974, starred in Titanic, The Wolf of Wall Street, and The Revenant. Who is he?",
  "Leonardo DiCaprio", "", "images/dicaprio.jpg")
q("nov_bdays", 2, "Born November 27, 1942, this guitarist recorded “Purple Haze.” Who is he?", "Jimi Hendrix")
q("nov_bdays", 3, "Born November 30, 1835, Samuel Clemens wrote under what pen name?", "Mark Twain")
q("nov_bdays", 4, "Born November 12, 1980, this actor played Ken in 2023's Barbie. Who is he?", "Ryan Gosling")
q("nov_bdays", 5, "Born November 22, 1984, this actress played Black Widow in the Marvel movies. Who is she?", "Scarlett Johansson")

q("nov_pop", 1, "Which Michael Jackson album was released on November 30, 1982?", "Thriller",
  "images/michael_jackson.jpg", "images/thriller_cover.jpg")
q("nov_pop", 2, "The first Harry Potter film opened in U.S. theaters in November 2001. What was its U.S. title?",
  "Harry Potter and the Sorcerer's Stone")
q("nov_pop", 3, "Which Disney movie featuring the song “Let It Go” was released in November 2013?", "Frozen")
q("nov_pop", 4, "Which cartoon character made his debut in “Steamboat Willie” in November 1928?", "Mickey Mouse")
q("nov_pop", 5, "Which Nintendo console with a motion-sensing controller launched in North America in November 2006?", "Wii")

M = "Name this song.  Who performs it?  One Point Each!"
q("music_tg", 1, M, "Connected – Stereo MC's (1992)", "", "images/connected_album.jpg", "media/connected.mp3", 0, 30)
q("music_tg", 2, M, "Alice's Restaurant – Arlo Guthrie", "", "", "media/alices_restaurant.mp3", notes="placeholder audio")
q("music_tg", 3, M, "Thank You – Dido", "", "", "media/thank_you_dido.mp3", notes="placeholder audio")
q("music_tg", 4, M, "Gravy (For My Mashed Potatoes) – Dee Dee Sharp", "", "", "media/gravy.mp3", notes="placeholder audio")
q("music_tg", 5, M, "Thank You (Falettinme Be Mice Elf Agin) – Sly & the Family Stone", "", "", "media/thank_you_sly.mp3",
  notes="placeholder audio")

q("holiday_travel", 1, "What movie does this describe?\n\nAfter their flight from New York to Chicago is diverted, two men endure "
  "a chaotic three-day road trip using various modes of transportation to try and make it home for Thanksgiving.",
  "Planes, Trains and Automobiles", "images/car_with_turkey.jpg", "images/planes_trains_poster.jpg")
q("holiday_travel", 2, "What movie does this describe?\n\nA family flies to Paris for Christmas and accidentally leaves their "
  "8-year-old son home alone to fend off two burglars.", "Home Alone")
q("holiday_travel", 3, "What movie does this describe?\n\nWhen fog grounds their tropical getaway, a couple is forced to visit "
  "all four of their divorced parents on Christmas Day.", "Four Christmases")
q("holiday_travel", 4, "What movie does this describe?\n\nAfter losing her job, a single mom flies home to Baltimore for a "
  "chaotic Thanksgiving with her family (1995).", "Home for the Holidays")
q("holiday_travel", 5, "What movie does this describe?\n\nEmpty nesters decide to skip Christmas and take a Caribbean cruise "
  "instead, much to their neighbors' outrage.", "Christmas with the Kranks")

q("potluck_1", 1, "What is the largest planet in our solar system?", "Jupiter")
q("potluck_1", 2, "How many players does each team have on the field in soccer?", "11")
q("potluck_1", 3, "What is the chemical symbol for gold?", "Au")
q("potluck_1", 4, "The city of Marrakech is in which country?", "Morocco")
q("potluck_1", 5, "In what year did the Titanic sink?", "1912")

MV = "What movie is this scene from?"
for n, (f, a) in enumerate([("addams_family_values", "Addams Family Values (1993)"), ("free_birds", "Free Birds (2013)"),
                            ("pieces_of_april", "Pieces of April (2003)"), ("scent_of_a_woman", "Scent of a Woman (1992)"),
                            ("hannah_and_her_sisters", "Hannah and Her Sisters (1986)")], 1):
    q("movies_tg", n, MV, a, "", "", f"media/{f}.mp4", notes="placeholder clip")

# ---------------------------------------------------------------- IDs + linking
# An older (retired) category that once used the DiCaprio question -> shows a question living in two categories.
cats.append(("best_actors", "Best Actors", "Leading men and their big roles…", "standard", ""))
CID = {slug: f"C{i:04d}" for i, (slug, *_) in enumerate(cats, 1)}
QID = {}
for i, (cat, n, *_rest) in enumerate(Q, 1):
    QID[(cat, n)] = f"Q{i:05d}"

sheet("Categories", ["category_id", "title", "blurb", "type", "season_tag", "retired", "notes"],
      [(CID[slug], t, b, ty, st, "Y" if slug == "best_actors" else "",
        "older category, kept for history" if slug == "best_actors" else "")
       for slug, t, b, ty, st in cats],
      [11, 34, 60, 10, 16, 8, 26, 12, 14],
      auto=[("(auto) questions", '=IF($A{r}="","",COUNTIF(CategoryQuestions!$A:$A,$A{r}))'),
            ("(auto) last played", '=IF($A{r}="","",IF(COUNTIF(Usage!$D:$D,$A{r})=0,"",MAXIFS(Usage!$A:$A,Usage!$D:$D,$A{r})))')])

# Who is likely to know each question (sample tags; the app's "Tag ages with AI" can suggest these)
ALL = "Boomer, Gen X, Millennial, Gen Z"
GEN = {("turkey", n): ALL for n in range(1, 6)}
GEN.update({("tg_potluck", n): "Boomer, Gen X, Millennial" for n in range(1, 6)})
GEN.update({("potluck_1", n): ALL for n in range(1, 6)})
GEN.update(dict(zip([("tg_mishaps", n) for n in range(1, 6)], [
    "Boomer, Gen X, Millennial", "Millennial, Gen Z", "Gen X, Millennial", "Boomer", "Millennial"])))
GEN.update(dict(zip([("nov_bdays", n) for n in range(1, 6)], [
    "Gen X, Millennial", "Boomer", "Boomer, Gen X", "Millennial, Gen Z", "Millennial, Gen Z"])))
GEN.update(dict(zip([("nov_pop", n) for n in range(1, 6)], [
    "Boomer, Gen X", "Millennial, Gen Z", "Millennial, Gen Z, Gen Alpha", ALL + ", Gen Alpha", "Millennial, Gen Z"])))
GEN.update(dict(zip([("music_tg", n) for n in range(1, 6)], ["Gen X", "Boomer", "Gen X, Millennial", "Boomer", "Boomer"])))
GEN.update(dict(zip([("holiday_travel", n) for n in range(1, 6)], [
    "Boomer, Gen X", "Gen X, Millennial", "Millennial", "Boomer, Gen X", ""])))  # one left untagged on purpose
GEN.update(dict(zip([("movies_tg", n) for n in range(1, 6)], [
    "Gen X, Millennial", "Gen Z, Gen Alpha", "Gen X, Millennial", "Boomer, Gen X", "Boomer"])))

# Sample difficulty (for a typical bar team): easy = most teams get it, hard = few do
HARD = {("turkey", 3), ("turkey", 5), ("tg_potluck", 1), ("tg_potluck", 3), ("tg_mishaps", 4), ("nov_bdays", 2),
        ("music_tg", 1), ("music_tg", 4), ("music_tg", 5), ("movies_tg", 3), ("movies_tg", 5)}
EASY = {("turkey", 1), ("turkey", 2), ("tg_potluck", 2), ("tg_potluck", 4), ("tg_mishaps", 3), ("nov_bdays", 1),
        ("nov_bdays", 3), ("nov_pop", 1), ("nov_pop", 3), ("nov_pop", 4), ("holiday_travel", 1), ("holiday_travel", 2),
        ("potluck_1", 1), ("potluck_1", 2), ("potluck_1", 3)}
DIFF = {k: "hard" if k in HARD else "easy" if k in EASY else "medium" for k in [(c, n) for c, n, *_ in Q]}
DIFF[("holiday_travel", 5)] = ""  # one left unrated on purpose

# Extra text shown on the slides (question_note under the question; answer_note revealed with the answer)
QNOTE = {("music_tg", n): "One point for the title, one for the artist" for n in range(1, 6)}
QNOTE[("nov_bdays", 3)] = "Hint: he wrote Tom Sawyer"
ANOTE = {("music_tg", n): "One Point for Title · One Point for Artist" for n in range(1, 6)}
ANOTE[("turkey", 5)] = "Minnesota raises around 40 million turkeys a year"
ANOTE[("tg_potluck", 1)] = "Wisconsin grows more than half of the country's cranberries"

sheet("Questions", ["question_id", "question", "answer", "question_image", "answer_image", "media",
                    "clip_start", "clip_end", "seconds", "retired", "notes", "generations", "difficulty",
                    "question_note", "answer_note"],
      [(QID[(cat, n)], qu.replace("  One Point Each!", "") if cat == "music_tg" else qu, an, qi, ai, me, st, en, se, "",
        no, GEN.get((cat, n), ""), DIFF.get((cat, n), ""), QNOTE.get((cat, n), ""), ANOTE.get((cat, n), ""))
       for cat, n, qu, an, qi, ai, me, st, en, se, no in Q],
      [10, 60, 34, 26, 26, 26, 9, 9, 8, 8, 18, 22, 10, 30, 30, 12, 12, 14],
      auto=[("(auto) in categories", '=IF($A{r}="","",COUNTIF(CategoryQuestions!$C:$C,$A{r}))'),
            ("(auto) times asked", '=IF($A{r}="","",COUNTIF(Usage!$E:$E,$A{r}))'),
            ("(auto) last asked", '=IF($A{r}="","",IF(COUNTIF(Usage!$E:$E,$A{r})=0,"",MAXIFS(Usage!$A:$A,Usage!$E:$E,$A{r})))')])

links = [(CID[cat], n, QID[(cat, n)]) for cat, n, *_ in Q]
links.append((CID["best_actors"], 1, QID[("nov_bdays", 1)]))   # same question, second category
sheet("CategoryQuestions", ["category_id", "position", "question_id"], links, [12, 9, 12, 30, 60, 30],
      auto=[("(auto) category", lookup("A", "Categories", "B")),
            ("(auto) question", lookup("C", "Questions", "B")),
            ("(auto) answer", lookup("C", "Questions", "C"))])

# ---------------------------------------------------------------- Sets
S = "2026-W48-A"
order = [("turkey", ""), ("tg_potluck", ""), ("tg_mishaps", ""), ("nov_bdays", ""), ("nov_pop", ""),
         ("music_tg", "Y"), ("holiday_travel", ""), ("potluck_1", ""), ("movies_tg", "")]
sheet("Sets", ["set_id", "round", "slot", "category_id", "bonus"],
      [(S, i // 3 + 1, i % 3 + 1, CID[slug], b) for i, (slug, b) in enumerate(order)],
      [12, 7, 6, 12, 7, 34], auto=[("(auto) category", lookup("D", "Categories", "B"))])

sheet("Tiebreakers", ["set_id", "order", "question", "answer"], [
    (S, 1, "What year was Dolly Parton born?", "1946"),
    (S, 2, "What year did World War I officially end?", "1918"),
    (S, 3, "What year did Sabrina Carpenter release her debut album, ‘Eyes Wide Open’?", "2015"),
], [12, 7, 70, 12])

D = dt.date
sheet("Schedule", ["date", "venue_id", "set_id", "host", "calendar_event"], [
    (D(2026, 11, 23), "wine_press", S, "James", "101 Wine Press - James"),
    (D(2026, 11, 24), "taproom", S, "Troy", "Sample Taproom - Troy"),
    (D(2026, 11, 25), "pub", S, "Aaron", "Sample Pub - Aaron"),
], [12, 14, 12, 14, 30], date_cols=[1])

# Pretend history, to show the repeat checks:
#  - Sample Pub played the Turkey category last December          -> category warning
#  - Sample Taproom heard the DiCaprio question in March, in another category -> question warning
usage = [(D(2025, 12, 3), "pub", "2025-W49-B", CID["turkey"], QID[("turkey", n)]) for n in range(1, 6)]
usage.append((D(2026, 3, 10), "taproom", "2026-W11-A", CID["best_actors"], QID[("nov_bdays", 1)]))
sheet("Usage", ["date", "venue_id", "set_id", "category_id", "question_id"], usage, [12, 12, 12, 12, 12],
      date_cols=[1])

wb._sheets = [wb[n] for n in ["Settings", "Venues", "Categories", "Questions", "CategoryQuestions", "Sets",
                              "Tiebreakers", "Schedule", "Usage"]]

ws = wb["Categories"]
dv = DataValidation(type="list", formula1='"standard,music,movie"')
ws.add_data_validation(dv); dv.add("D2:D2000")
dv = DataValidation(type="list", formula1='"easy,medium,hard"', allow_blank=True)
wb["Questions"].add_data_validation(dv); dv.add("M2:M5000")
for tab, col in (("Categories", "F"), ("Questions", "J")):  # retired
    dv = DataValidation(type="list", formula1='"Y"', allow_blank=True)
    wb[tab].add_data_validation(dv); dv.add(f"{col}2:{col}5000")
dv = DataValidation(type="list", formula1="=Sets!$A$2:$A$1000", allow_blank=True)
wb["Schedule"].add_data_validation(dv); dv.add("C2:C2000")
ws = wb["CategoryQuestions"]
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5")
ws.add_data_validation(dv); dv.add("B2:B10000")

wb.save("trivia_bank.xlsx")
print("wrote trivia_bank.xlsx")
