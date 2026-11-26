# TBTgen: Trivia By Troy deck generator

Builds a complete PowerPoint trivia deck for each venue-night from one question bank (a Google Sheet), plus
an at-home video version, from a browser-based app or the command line.

## What a generated deck does
- Venue slides (title, host, format, rules, upcoming nights, round/intermission, results, tip QR) carry the venue
  logo (venues without one get the Trivia By Troy logo). Slide 2 is "Tonight's host" + the host's name.
- 3 rounds × 3 categories × 5 questions, one category per night marked as the double-points bonus.
- Question slides auto-advance after 60 s (or clip length + 30 s for music/movie). The Trivia By Troy logo starts
  grey and fills with colour from the bottom over that time, as a countdown. Pictures appear immediately.
- Music/movie clips play automatically on the question slide.
- Answer rounds: each question re-appears (without its picture); one click reveals the answer (+ answer note,
  answer picture; music replays).
- Tiebreaker slides, and an answer key `.txt` next to each deck.
- File names include the date, venue and host: `2026-11-23 101 Wine Press (James).pptx`.
- Checks before building: warns if a venue already saw a category, or a question (even in a different category),
  within `repeat_lookback_days`, or if venues in the same region on back-to-back nights share questions; stops on
  retired or missing questions, a category without exactly positions 1–5, or a question used twice in one set.
- Plays correctly in PowerPoint (Mac or Windows). Keynote's import does not auto-advance the timed slides.

## Mac setup (once)
```
brew install python ffmpeg
python3 -m pip install --user --break-system-packages python-pptx openpyxl pillow lxml
```
(`--break-system-packages` is Homebrew's way of making you confirm a per-user install; with `--user` it only
affects your account.) For the optional AI assistant also install `anthropic` and/or `google-genai` the same way.

## Connect to the Google Sheet (once)
1. In the Trivia Bank sheet: Extensions → Apps Script, paste in `tbtgen_sheet.gs`, Save, reload the sheet.
2. **TBTgen → Set up generator access…** shows a secret key and these steps:
   in the Apps Script editor, **Deploy → New deployment → Web app**, Execute as **Me**,
   Who has access **Anyone**, Deploy, copy the **Web app URL**.
3. In the TBTgen folder (next to `build.py`), create `tbtgen_config.json` (ignored by Git; never commit it):
   ```json
   {"sheet_url": "https://script.google.com/macros/s/…/exec", "token": "the key from step 2"}
   ```
Now the app and `build.py` read the sheet live and save edits back to it. Without that file, `trivia_bank.xlsx`
is used; `--bank some.xlsx` forces a file.
- **After `tbtgen_sheet.gs` changes**, paste in the new version and use Deploy → Manage deployments → edit →
  Version: **New version** (the URL stays the same). Until then the app's saves may be refused.
- To cut off access: **TBTgen → Make a new generator key** (then update `tbtgen_config.json`).
- The script also adds **TBTgen → Pull schedule from calendar…**: it reads a week from the Trivia By Troy Schedule
  calendar and adds a Schedule row per night ("Venue - Host" → venue, host). Troy then picks each night's set.

## The app
Double-click **TBTgen.command** (or run `python3 tbtgen_app.py`); your browser opens the app. Leave the Terminal
window open while using it; close it or click **Quit** to stop. The first time, macOS may refuse to open a
downloaded .command file: right-click it → Open → Open.
- **Nights**: a week of the schedule. Pick each night's set from its drop-down (saved to the Schedule tab), filter
  weekdays, tick nights, **Check** or **Build** decks; "Log as played" also writes the Usage tab. Nights with no set
  or an unknown venue can't be ticked. Each night shows the venue's audience and flags sets light on it.
- **Sets**: **Preview deck** builds a deck of the set with the TBT logo and no host (not tied to a night, nothing logged).
  Each set's 9 categories by round, bonus, problems, age and difficulty mix, where/when each category was
  last played, which nights use it, and its video. **New set / Edit set**: a 3×3 grid of category pickers, bonus,
  tiebreakers, with live warnings (played recently, not 5 questions, a question in two categories).
- **Categories**: browse, create and edit categories and pick/reorder their 5 questions (search, or New question).
- **Questions**: search (text, answer, notes, ID) or filter by category, difficulty or generation; click for
  pictures, clips and history; **New / Edit question** (with a live slide preview), uploading pictures and clips.
- **Venues**: add and edit venues (name, calendar name, logo upload, night, region, upcoming nights, audience, active).
- **Video**: make the at-home video for a set, watch it in the page, download it or open it.
- **Files**: every deck, answer key and video, with View / Open / Download.
- **Import**: read Troy's old decks into the bank (see below).
- **Settings**: the Settings tab with plain-English labels and range checks, and the AI assistant setup.
- **Status**: setup checks, a spreadsheet connection test, counts, the fonts videos use, AI spend.

Every table sorts by any column (click the header; again to reverse) and long lists are paged; Sets have
search and sort too. Categories that share a name (Movies, Music…) show a date (last played, else imported) to
tell them apart. Edits are saved straight to the spreadsheet. Uploaded pictures/clips go into this computer's `assets/` folder
(copy them to the Drive folder so other computers have them). Data is cached until you click **Reload data**.

## The spreadsheet
Grey "(auto)" columns are formulas: don't type in them. IDs (`Q00001`, `C0001`) are permanent values, never
formulas, so sorting rows is safe. Don't delete questions or categories; mark them `retired` so history stays.
New rows get IDs automatically (Google Sheets: as you type; .xlsx: on every build, or `build.py --ids`).

| Tab | Columns |
|---|---|
| **Questions** | question, answer, question_image, answer_image, media, clip_start/clip_end (seconds), seconds (timer override), retired, notes (host only), generations, difficulty, question_note, answer_note |
| **Categories** | title, blurb (intro slide), type `standard`/`music`/`movie`, season_tag, retired, notes |
| **CategoryQuestions** | category_id, position 1–5, question_id. A question can be in any number of categories |
| **Sets** | set_id, round 1–3, slot 1–3, category_id, bonus (exactly one `Y`), tiebreaker_id (the set's tiebreaker group) |
| **Tiebreakers** | tiebreaker_id (a group, e.g. `T001`), order, question, answer. A set picks one group; a group can be reused |
| **Venues** | venue_id, name, logo, region, night, upcoming_nights, active, calendar_name (as written in calendar events), audience |
| **Schedule** | date, venue_id, set_id, host, calendar_event. Nights without a set are skipped |
| **Usage** | which venue heard which question, in which category, when (written by "Log as played" / `--record`) |
| **Settings** | default_seconds, media_extra_seconds, repeat_lookback_days, calendar_id, AI settings, … |

Picture and clip paths are relative to `assets/` (e.g. `images/turkey.jpg`, `media/song.mp3`).

### Slide text: notes and formatting
- **question_note** is shown smaller under the question (a hint, "One point each!"); **answer_note** is revealed
  together with the answer (scoring note, fun fact). Both are optional; **notes** is for you only.
- **Formatting** works in the question, answer, both notes and tiebreaker questions:

| Type | Shows as |
|---|---|
| `*italic*` | *italic* (inside an italic note it turns upright, for emphasis) |
| `**bold**` | **bold** |
| `***both***` | bold italic |
| `[red]text[/red]` | red text |
| `[gold]text[/gold]` | gold text (`[/]` also closes a colour) |

A `*` only counts at the edge of a word, so M\*A\*S\*H and Q\*bert stay as typed; `\*` gives a literal star.
Answer keys, searching and the AI use the plain text.

### Balancing for the crowd (generations and difficulty)
Questions have a **generations** tag (Boomer 1946–64, Gen X 1965–80, Millennial 1981–96, Gen Z 1997–2012,
Gen Alpha 2013+) meaning "likely to know it", and a **difficulty** (easy / medium / hard, for a typical bar team).
Set them in the question editor or with **Tag with AI**. Venues have an **audience** (who usually plays there). Then:
- each set shows an **age mix** and **difficulty mix**, live in the set editor too;
- the Nights tab flags "⚠ light on Gen Z" when under a quarter of the set's tagged questions suit part of the
  venue's audience;
- Generate with AI and Category ideas can aim at chosen generations.

## AI assistant (optional)
Runs **only** when you click a button; it suggests, and nothing changes until you accept.
- **Check with AI** (a question, a category, or a set): is the answer correct and the single best answer (what
  else to accept), is the wording precise? Shows issues, a suggested rewrite, fun fact and sources;
  **Use suggestion…** opens the editor pre-filled.
- **Generate with AI…** (category editor): candidate questions in your style for a topic, plus a blurb.
- **Category ideas with AI…** (Categories tab): titles, blurbs and angles to spark thinking; no questions.
- **Tag with AI**: rates difficulty and generations. Cheap: low effort, no web search.
- Providers: **Claude** (Anthropic; default claude-opus-5-5) or **Gemini** (Google; pick a model with List models),
  set up on Settings → AI assistant. Keys are saved in `tbtgen_config.json` (not in Git or the sheet). Spend is
  logged to `ai_usage.jsonl` and shown on Settings and Status. Claude requests use Anthropic's server-side
  fallback: if the model declines a request, the API retries it on a fallback model.

## At-home video
`python3 build.py --preview SET_ID        # a deck of one set, not tied to a night -> out/preview/
python3 build.py --video SET_ID` (or the Video tab) makes `out/video/SET_ID at-home.mp4`: title, how to play,
3 rounds with the logo countdown on every question (default 30 s), the answers after each round, then a thanks
card. Music and movie clips play on their questions (clip length + 10 s, or `video_media_extra_seconds`).
Needs ffmpeg. Uses Trebuchet MS / Savoye LET / Brush Script when the computer has them (a Mac does).
Song/movie clips and web photos may be copyrighted: fine for private use, risky to publish.

## Importing old decks
`python3 import_decks.py FOLDER [--pictures] [--clips]` or the app's **Import** tab ("Choose folder…" on a Mac).
Reads Keynote (.key, folder or single-file; Keynote 6+ fully, '09-era best effort) and PowerPoint (.pptx) decks
directly, with no Keynote needed. "Questions 1 - 5" intros give the categories, "Question N" slides the questions
(+ picture/clip; subtext styled italic or smaller at the end of the question, like "One Point Each!", and any
extra text become question_note), answer slides the answers (+ answer_note, answer pictures).
Repeats across decks are merged, with every appearance (deck, date from the file name, venue) kept.
- Decks are read where they are (nothing is copied or uploaded); only the slide text and layout are read, so even
  huge decks use little memory. Decks in iCloud Drive or Google Drive that are "cloud only" get downloaded in
  full when read (slow for a big archive): first right-click the folder → Download Now / Keep Downloaded (iCloud)
  or Available offline (Google Drive). iCloud placeholders that can't be read are listed in a warning, not skipped silently.
- Pictures and clips are optional; unticked, they're only counted (with total MB) so you can judge the space.
- Output in `imports/<date-time>/`: `review.xlsx` (decks, categories, questions, "check these") and media.
- In the Import tab, tick categories and **Add to the bank**: new questions are appended, existing ones reused,
  media copied into `assets/…/imported/`, play history optionally recorded. Adding the same import twice adds nothing.
- Text formatting (colours, italics) in old decks is not carried over.

## Command line
```
python3 build.py                         # every scheduled night that has a set -> out/
python3 build.py --week next --day wed   # just next week's Wednesday decks
python3 build.py --date 2026-12-02       # one night;  --week this|next|DATE;  --day mon,tue;  --venue a,b
python3 build.py --only 2026-12-02:wine_press,2026-12-02:pops
python3 build.py --from 2026-11-23 --to 2026-11-29 --record   # also log to the Usage tab
python3 build.py --check [--strict]      # only check for repeats/problems (--strict: warnings are errors)
python3 build.py --video SET_ID          # at-home video (--video-seconds, --video-answer-seconds,
                                         #   --video-height 1080, --video-no-media)
python3 build.py --ids                   # .xlsx only: fill in missing IDs
```

## Files
- `tbtgen_app.py` + `TBTgen.command`: the app. `build.py`: the deck generator. `video.py`: the at-home video.
- `bank_write.py`: saving to the sheet/xlsx. `ai.py`: the AI assistant.
- `import_decks.py` + `keynote_read.py`: the old-deck importer.
- `tbtgen_sheet.gs`: the Google Sheet's Apps Script (menu, IDs, calendar pull, the web app the generator talks to).
- `trivia_bank.xlsx`: a sample bank (`make_sample_bank.py` rebuilds it); `make_sample_media.py` makes placeholder clips.
- `keynote_finish.applescript`: Mac only, untested attempt to fix Keynote's auto-advance after importing a .pptx.
- `assets/`: logos and pictures. `assets/media/` (songs, clips) is not in Git: copy it from Google Drive.
- Not in Git: `out/` (decks, videos), `build/` (cache), `imports/`, `tbtgen_config.json`, `ai_usage.jsonl`.

## Where things live
- Code: github.com/jepace/TBTgen (private)
- The question bank: the **Trivia Bank** Google Sheet in the Drive folder **TBTgen** (media in TBTgen/assets)
- The schedule: the "Trivia By Troy Schedule" Google Calendar, pulled into the sheet with the TBTgen menu

## Security notes
- The app only listens on this computer (127.0.0.1), checks every request carries the per-launch key, and refuses
  requests addressed to any other host name, so web pages can't drive it.
- `tbtgen_config.json` holds the sheet key and AI keys: keep it private (when the app saves keys it makes the
  file readable only by you; if you made it by hand, `chmod 600 tbtgen_config.json`).
- The sheet's web app accepts only requests with the key, and only writes the columns TBTgen edits. Text that
  starts with `=` is stored as text, never as a formula.
- Picture/clip paths in the sheet must stay inside `assets/` (absolute paths and `..` are refused by Check), and
  file names made from venue, host or set names can't point outside `out/`. No command is ever run through a shell.
