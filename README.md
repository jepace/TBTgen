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
cd ~/TBTgen
python3 -m venv .venv
.venv/bin/pip install --upgrade pip python-pptx openpyxl pillow lxml
```
Then run everything with `.venv/bin/python` (e.g. `.venv/bin/python build.py`), or `source .venv/bin/activate` first.

## Run (needs Python 3 with python-pptx, openpyxl, Pillow, lxml; and ffmpeg)
```
python3 build.py                      # build every scheduled night that has a set -> out/
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
