# TBTgen: Trivia By Troy deck generator (prototype)

Builds a complete PowerPoint trivia deck for each venue-night from one spreadsheet.

## Files
- `trivia_bank.xlsx`: the question bank and schedule (edit this). Grey "(auto)" columns are formulas: don't type in them.
  - **Categories**: `C0001`…: title, intro blurb, type `standard`/`music`/`movie`, season tag, retired
  - **Questions**: `Q00001`…: question, answer, question image, optional answer image, media clip
    (+ optional `clip_start`/`clip_end` in seconds, and `seconds` to override the 60-second timer), retired.
    Shows how many categories use it, how often it was asked, and when it was last asked.
  - **CategoryQuestions**: which question goes in which category, position 1–5. A question can be in any
    number of categories, so reusing a question later = adding a row here.
  - **Sets**: a night's game = 9 categories (round 1–3 × slot 1–3); mark exactly one `bonus` = Y
  - **Venues**: name, logo, region, and `calendar_name` (how the venue is written in calendar event titles)
  - **Schedule**: date, venue, set, host, calendar event. Nights without a set are skipped by the build.
  - **Tiebreakers**, **Settings** (incl. `calendar_id`)
  - **Usage**: which venue heard which question, in which category, when (filled in by `--record`)
- IDs are permanent values, never formulas, so sorting rows is safe. Don't delete questions or categories;
  mark them `retired` so their history stays.
  - In Excel/Numbers: leave the ID blank on new rows and run `python3 build.py --ids` (every build also does it).
  - In Google Sheets: install `tbtgen_sheet.gs` (Extensions → Apps Script) and IDs appear as you type.
- `assets/`: logos, images, audio/video referenced from the sheet (paths are relative to `assets/`).
  `assets/media/` (songs, movie clips) is not in Git: copy it from Google Drive (TBTgen/assets/media),
  or run `python3 make_sample_media.py` to create placeholder clips for the sample bank.
- `keynote_finish.applescript`: Mac only. Opens a generated .pptx in Keynote, fixes the auto-advance timing,
  saves a .key file. (Not yet confirmed working.)
- `tbtgen_sheet.gs`: Apps Script for the Google Sheet (Extensions → Apps Script). Adds a **TBTgen** menu:
  - IDs appear automatically as you type new questions/categories.
  - **Pull schedule from calendar…** reads a week from the Trivia By Troy Schedule calendar and adds a
    Schedule row per night ("Venue - Host" → venue, host). Then pick a set for each night from the drop-down.
- `build.py`: the generator

## Where things live
- Code: this repo (github.com/jepace/TBTgen, private)
- The real question bank, media and images: Google Drive folder **TBTgen** ("Trivia Bank (new layout)" sheet)
- The real schedule: the "Trivia By Troy Schedule" Google Calendar (not wired in yet)

## Connect the generator to the Google Sheet (once)
1. In the Trivia Bank sheet: Extensions → Apps Script, paste in `tbtgen_sheet.gs`, Save, reload the sheet.
2. **TBTgen → Set up generator access…** shows a secret key and these steps:
   in the Apps Script editor, **Deploy → New deployment → Web app**, Execute as **Me**,
   Who has access **Anyone**, Deploy, copy the **Web app URL**.
3. Next to `build.py`, create `tbtgen_config.json` (it's ignored by Git; don't commit it):
   ```json
   {"sheet_url": "https://script.google.com/macros/s/…/exec", "token": "the key from step 2"}
   ```
Now every run reads the sheet live, and `--record` writes the Usage log back to it.
Without that file, `trivia_bank.xlsx` is used; `--bank some.xlsx` forces a file.
If the script changes later: Deploy → Manage deployments → edit → Version: **New version** (same URL).
To cut off access: **TBTgen → Make a new generator key**.

## Mac setup (once)
```
brew install python ffmpeg
python3 -m pip install --user --break-system-packages python-pptx openpyxl pillow lxml
```
(`--break-system-packages` is Homebrew's way of making you confirm a per-user install; with `--user` it only
affects your account. Prefer a virtual environment? `python3 -m venv .venv` and use `.venv/bin/python`.)

## The app
Double-click **TBTgen.command** (or run `python3 tbtgen_app.py`); your browser opens the app.
- **Nights**: a week of the schedule. Filter weekdays, tick nights, **Check** or **Build** decks;
  "Log as played" also writes the Usage tab. Nights with no set or an unknown venue can't be ticked.
- **Sets**: each set's 9 categories by round, bonus, problems, when/where each category was last played,
  which nights use it, and its video. **New set / Edit set**: a 3×3 grid of category pickers, bonus, tiebreakers,
  with live warnings (played recently, not 5 questions, a question in two categories).
- **Categories**: browse, **create and edit** categories and pick/reorder their 5 questions (search, or New question).
- **Questions**: search the bank (text, answer, ID) or filter by category; click for pictures, clips and history;
  **New question / Edit question**, including uploading pictures and clips into the assets folder.
- On **Nights**, each night has a set drop-down: picking one saves it to the Schedule tab.
- **Venues**: list, add and edit venues (name, calendar name, logo upload, night, region, upcoming nights, active).
  New venues get an ID and calendar name suggested from the name; duplicate IDs/calendar names are refused.
- **Settings**: edit the Settings tab with plain-English labels and range checks, or add a new setting.

Everything you edit in the app is saved straight to the spreadsheet (the Google Sheet through its web app, or
trivia_bank.xlsx). After updating `tbtgen_sheet.gs`, redeploy it (Deploy → Manage deployments → edit → Version:
New version) or the app's saves will be refused. Uploaded pictures/clips go into this computer's `assets/` folder.
- **Video**: make the at-home video for a set (seconds per question/answer, 720p/1080p, with or without
  music/movie categories), watch it in the page, download it or open it in your player.
- **Files**: every deck, answer key and video, with View / Open / Download.
- **Status**: setup checks, a spreadsheet connection test, counts, the fonts videos use, and the Settings tab.
The first time, macOS may refuse to open a downloaded .command file: right-click it → Open → Open.
Leave the Terminal window open while using the app; close it or click Quit to stop.

## At-home video
`python3 build.py --video SET_ID` (or the Video tab) makes `out/video/SET_ID at-home.mp4`: title, how to play,
3 rounds with the logo countdown on every question (default 30 s), "Pencils down" and the answers after each
round, then a thanks card. Music and movie clips play on their questions (clip length + 10 s, or the Settings
tab's `video_media_extra_seconds`). Options: `--video-seconds`, `--video-answer-seconds`, `--video-height 1080`,
`--video-no-media`. Needs ffmpeg. Text is drawn with Trebuchet MS / Savoye LET / Brush Script when the
computer has them (a Mac does). Song/movie clips and web photos may be copyrighted: fine for private use,
risky to publish.

## Connect the generator to the Google Sheet (once)
1. In the Trivia Bank sheet: Extensions → Apps Script, paste in `tbtgen_sheet.gs`, Save, reload the sheet.
2. **TBTgen → Set up generator access…** shows a secret key and these steps:
   in the Apps Script editor, **Deploy → New deployment → Web app**, Execute as **Me**,
   Who has access **Anyone**, Deploy, copy the **Web app URL**.
3. Next to `build.py`, create `tbtgen_config.json` (it's ignored by Git; don't commit it):
   ```json
   {"sheet_url": "https://script.google.com/macros/s/…/exec", "token": "the key from step 2"}
   ```
Now every run reads the sheet live, and `--record` writes the Usage log back to it.
Without that file, `trivia_bank.xlsx` is used; `--bank some.xlsx` forces a file.
If the script changes later: Deploy → Manage deployments → edit → Version: **New version** (same URL).
To cut off access: **TBTgen → Make a new generator key**.

## Mac setup (once)
```
brew install python ffmpeg
python3 -m pip install --user --break-system-packages python-pptx openpyxl pillow lxml
```
(`--break-system-packages` is Homebrew's way of making you confirm a per-user install; with `--user` it only
affects your account. Prefer a virtual environment? `python3 -m venv .venv` and use `.venv/bin/python`.)

## The app
Double-click **TBTgen.command** (or run `python3 tbtgen_app.py`). Your browser opens a page showing a week
of the schedule: pick the week and weekdays, tick nights, then **Check** or **Build**; "Log as played" also
writes the Usage tab. Nights with no set or an unknown venue are shown but can't be ticked.
The first time, macOS may refuse to open a downloaded .command file: right-click it → Open → Open.
Leave the Terminal window open while using the app; close it or click Quit to stop.

## Command line (needs Python 3 with python-pptx, openpyxl, Pillow, lxml; and ffmpeg)
```
python3 build.py                      # build every scheduled night that has a set -> out/
python3 build.py --week next --day wed   # just next week's Wednesday decks
python3 build.py --date 2026-12-02    # one night;  --week this|next|DATE;  --day mon,tue;  --venue a,b
python3 build.py --from 2026-11-23 --to 2026-11-29 --record
python3 build.py --check              # only check for repeats/problems
```

## What a generated deck does
- Venue slides (title, format, rules, upcoming nights, round/intermission, results, tip QR) carry the venue logo.
- Question slides auto-advance after 60 s (or clip length + 30 s for music/movie). The Trivia By Troy logo
  starts grey and fills in with colour from the bottom over that time, as a countdown.
- Music/movie clips play automatically on the question slide.
- Answer rounds: each question re-appears; one click reveals the answer (+ answer image; music replays).
- An answer key `.txt` is written next to each deck.
- Checks: warns if a venue already saw a category, or a question (even in a different category), within
  `repeat_lookback_days`; if venues in the same region on back-to-back nights share questions; and stops on
  retired or missing questions, a category without exactly positions 1–5, or a question used twice in one set.
