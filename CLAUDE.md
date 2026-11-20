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
- Pending: fonts/sizing, Mac app packaging, venue logos for all venues, real questions in the sheet.
- Never commit generated decks (out/), cache (build/) or media (assets/media/).
