# TBTgen notes for Claude

Generates Trivia By Troy game decks (.pptx) from a question bank spreadsheet. See README.md.

- `build.py` is the whole generator. `make_sample_bank.py` writes the sample `trivia_bank.xlsx`.
- Data model: Categories (C####), Questions (Q#####), CategoryQuestions (many-to-many, position 1-5),
  Sets (9 categories, one bonus), Schedule, Usage (per question). IDs are permanent values, never formulas.
  Columns whose header starts with "(auto)" are formula helpers; the generator ignores them.
- Verified by the user: PowerPoint (Windows) plays the 60 s auto-advance, countdown logo wipe, media
  autoplay and click-to-reveal answers. Keynote import does NOT advance timed slides (it adds the
  transition delay after the 60 s build); `keynote_finish.applescript` is the untested fix.
- The user didn't want the question photo animated: it appears immediately. Only the logo countdown animates.
- Hosts are all on Macs today; presenting in PowerPoint is acceptable.
- Real data: Google Drive folder "TBTgen"; schedule in the "Trivia By Troy Schedule" calendar
  (events named "Venue - Host", ~16 venues Mon-Thu).
- Troy assigns a set to each night by hand (no fixed A/B groups). Flow: `tbtgen_sheet.gs` menu
  "Pull schedule from calendar" adds Schedule rows (date, venue_id via Venues.calendar_name, host,
  calendar_event) with set_id blank; Troy picks set_id from a drop-down; build.py skips rows with no set
  or unknown venue (prints SKIPPED). Venues without a logo fall back to the TBT logo.
- The Apps Script is syntax-checked and its parsing was tested in Node, but it has not been run in Google yet.
- Slide 2 is the host slide: "Tonight's host" + name, script font (Schedule tab `host` column).
  Deck file names include the host: "2026-11-23 101 Wine Press (James).pptx".
- The live Google Sheet is "Trivia Bank" (id 1ssBfjj26jdF-gxIXVKcn7VzYJZqV6unsyLx0OXZVRMU) in Drive folder TBTgen.
  Keep its layout in sync with make_sample_bank.py when columns change. It still holds the sample
  questions/sets; Venues and Schedule there are real.
- build.py reads the live Google Sheet through an Apps Script web app (doGet/doPost in tbtgen_sheet.gs,
  secret key in Script Properties) configured in tbtgen_config.json (gitignored). No Google Cloud/API
  libraries needed. Tested against a local stand-in server that mimics Google's 302 redirects, and the
  .gs logic against fake Apps Script services, but NOT yet deployed for real.
- The user won't use a venv. GUI = `tbtgen_app.py`: stdlib http.server on 127.0.0.1 + a page in the
  browser (Tkinter avoided: missing in Homebrew Python, old/crashy in Apple's). It shells out to
  `build.py --only date:venue,...`; per-launch token guards the API. `TBTgen.command` launches it on a Mac.
- Fonts come from Troy's Keynote template: Trebuchet MS (body), Savoye LET (script titles),
  Brush Script MT (bonus labels). Savoye is Mac-only. Sizes were guessed without seeing Savoye rendered.
- Pending: font sizes/layout review on a Mac, packaging as a real Mac app, venue logos for all venues, real questions in the sheet.
- At-home video is BUILT: `video.py` (Pillow draws slides in the deck style, ffmpeg makes one clip per slide
  in parallel, then concat). `build.py --video SET_ID` / app Video tab. Countdown = 61 PNG frames of the logo
  filling, overlaid per question. Tested here (fallback fonts); not yet seen on a Mac with the real fonts.
- App tabs: Nights (set drop-down per night saves to Schedule), Sets (New/Edit set), Categories (create/edit,
  pick+order 5 questions), Questions (New/Edit, upload pictures/clips to assets/), Venues (add/edit, logo upload to assets/venues/),
  Video, Files, Settings (labelled form for the Settings tab, int range checks, add keys), Status.
  Writes go through `bank_write.py` (SheetWriter -> web app doPost actions upsert/replace; XlsxWriter for the
  .xlsx); EDITABLE tab/column whitelist is duplicated in bank_write.py and tbtgen_sheet.gs: keep them in sync.
  New IDs are assigned by the writer (Apps Script nextNumber_ / max+1 in xlsx). Tested via Playwright on the xlsx
  and the .gs logic against fake Apps Script services; not yet against the real deployed sheet.
- Video style parity idea (not done): one shared layout spec used by both the pptx builder and video.py. Data cached until
  "Reload data". `build.check_sets()` = per-set checks without a schedule.
- Future idea (not started): an AI voice host (TTS, possibly a licensed
  clone of Troy's voice; Claude-drafted banter with a sheet column to override; pronunciation column).
  Big caveat discussed with the user: publishing song/movie clips and web photos raises copyright issues.
- Never commit generated decks (out/), cache (build/) or media (assets/media/).
