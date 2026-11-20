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
  (events named "Venue - Host"). Open question: fixed A/B set group per venue vs weekly assignment.
- Slide 2 is the host slide: "Tonight's host" + name, script font (Schedule tab `host` column).
  Deck file names include the host: "2026-11-23 101 Wine Press (James).pptx".
- The live Google Sheet is "Trivia Bank" (id 1ssBfjj26jdF-gxIXVKcn7VzYJZqV6unsyLx0OXZVRMU) in Drive folder TBTgen.
  Keep its layout in sync with make_sample_bank.py when columns change. Pending: fonts/sizing, Mac app packaging, Google Sheets/Calendar reading.
- Never commit generated decks (out/), cache (build/) or media (assets/media/).
