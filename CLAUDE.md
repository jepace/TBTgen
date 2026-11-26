# TBTgen notes for Claude

Generates Trivia By Troy game decks (.pptx) from a question bank spreadsheet. See README.md.

- `build.py` is the whole generator. `make_sample_bank.py` writes the sample `trivia_bank.xlsx`.
- Data model: Categories (C####), Questions (Q#####), CategoryQuestions (many-to-many, position 1-5),
  Sets (9 categories, one bonus), Schedule, Usage (per question). IDs are permanent values, never formulas.
  Columns whose header starts with "(auto)" are formula helpers; the generator ignores them.
- Verified by the user: PowerPoint (Windows) plays the 60 s auto-advance, countdown logo wipe, media
  autoplay and click-to-reveal answers. Keynote import does NOT advance timed slides (it adds the
  transition delay after the 60 s build); `keynote_finish.applescript` fixes it (delay 0.1 on timed slides) and saves
  .key: first real run on the user's Mac worked (45 slides fixed, classic Keynote). It prefers Keynote Creator Studio
  (bundle com.apple.Keynote; classic = com.apple.iWork.Keynote), compiles the Keynote part at run time for the chosen
  app, opens files with `open -b` (sandboxed app refused script-passed paths). TBT_KEYNOTE=<bundle id> overrides.
  build.py --keynote -> make_keynote() runs it after building (Mac only); app: "Also make Keynote copies" checkbox
  (remembered; also used by Preview deck). Files tab lists .key packages as one file (Open only).
- The user didn't want the question photo animated: it appears immediately. Only the logo countdown animates.
- Answer slides never reuse the question photo: only the answer photo (right side), revealed with the answer.
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
- AI assistant (`ai.py`): ONLY on explicit clicks (user requirement: never automatic). Providers Anthropic
  (anthropic SDK 2.x, beta messages + server-side fallback "default", web_search_20260209, effort from Settings,
  pause_turn resume, JSON parsed from the reply) and Gemini (google-genai, GoogleSearch grounding; model must be
  chosen via List models). Settings keys ai_provider/ai_model/ai_effort/ai_web_search (hidden from the generic
  Settings form, edited in the AI card). API keys in tbtgen_config.json. Usage log ai_usage.jsonl (gitignored).
  Tested with stand-in providers (no real keys here) - real API calls not yet exercised.
- Difficulty: Questions.difficulty (easy|medium|hard; build.DIFFS/parse_diff), rated by hand or by the same
  AI tag call as generations ("Tag with AI"); set_mix includes a diff count; live sheet column M.
- Generations: Questions.generations and Venues.audience (comma lists; build.GENS / parse_gens normalises
  aliases). App: set_mix (per-set counts), audience_fit (flag audience gens < 25% of tagged questions),
  question filter, editor ticks. AI: tag_generations (effort low, no web), category_ideas (no questions),
  generate(..., generations). Live sheet got the two columns (Questions L, Venues I).
- Deck importer: `keynote_read.py` decodes Keynote IWA (snappy + generic protobuf walk: show object type 2 field 3 =
  slide nodes (type 4) -> slide (5); text storages type 2001 field 3; drawables 3000-3099 carry data ids matched to
  Data/*-<id>.ext) and .pptx; `import_decks.py` classifies slides (intro / Question N / answer), dedupes by
  normalised question+answer, writes imports/<run>/import.json + review.xlsx; apply_import() bulk-appends via the
  new "append" action (gs + bank_write) and remembers applied keys in applied.json. Tested on Troy's (trimmed)
  sample .key, a single-file .key built from it, and a generated .pptx (all 45 questions recovered). Keynote '09
  (index.apxl) path is untested.
  Decks are read in place, never copied: .key reads only Index/*.iwa (media sizes from the listing, bytes only when
  importing); .pptx is read with zipfile+lxml (not python-pptx, which loaded every video into RAM), pictures/clips
  named by zip CRC+size. 405 MB test deck: 22 MB peak memory.
  find_decks() reports iCloud placeholders (".Name.key.icloud", older macOS) as cloud_only with a WARNING.
- Video style parity idea (not done): one shared layout spec used by both the pptx builder and video.py. Data cached until
  "Reload data". `build.check_sets()` = per-set checks without a schedule.
- Future idea (not started): an AI voice host (TTS, possibly a licensed
  clone of Troy's voice; Claude-drafted banter with a sheet column to override; pronunciation column).
  Big caveat discussed with the user: publishing song/movie clips and web photos raises copyright issues.
- Slide notes: Questions.question_note (smaller italic under the question, on question + answer slides) and
  answer_note (italic, revealed in the same click as the answer). Live sheet columns N/O; `notes` stays host-only.
  Importer fills them from extra slide text (NOTE_SKIP filters structural labels).
- Text markup applies to ALL slide text (build.text markup=True by default; video Canvas.text = rich() with align/valign;
  script/brush fonts: colour only). App: md() in lists, plainTxt() in <option>s and dialog titles. File names use plain().
  Markup syntax (question/answer/notes/tiebreakers/titles/blurbs/venue/host/settings): *italic* (flips), **bold**, ***both***, [red]/[gold]..[/red|/gold|/].
  build.markup_runs/plain + text(markup=True); video Canvas.rich wraps mixed runs; app md() mirrors the regexes
  (keep in sync); AI prompts use plain (ai._p). Colours in build.MARK_COLORS. Importer does not capture formatting.
- Security: app checks Host header (127.0.0.1/localhost:port, anti DNS-rebinding) + per-launch token; import scan
  passes folders after "--"; text starting with "=" is written as text (gs toCell_ prefixes ', XlsxWriter data_type s).
  build.sheet_call always POSTs (key only in the body); reads are POST action "tables"; doGet refuses everything.
- Paths: sheet file cells go through build.asset()/asset_problem() (must resolve inside assets/, else check() errors);
  output names through build.safe_name() (deck stem from venue+host, video from set_id); app import runs via
  import_dir() (strict name regex). All subprocess calls use arg lists (no shell). review.xlsx "=" cells forced text.
- Tiebreakers are groups: Tiebreakers.tiebreaker_id (T###) rows; Sets.tiebreaker_id (written on every row of the set)
  picks one. load_bank: bank["tiebreakers"][gid], bank["set_tiebreaker"][set_id]; legacy set_id-keyed rows still work.
  App: /api/tiebreakers, editTiebreakers() (dialog 2, next T### client-side). Live sheet migrated (Sets F, Tiebreakers A).
- App tables: generic sortRows/pageRows/markSort + RENDER[id] (th data-tbl/data-sort, #pg-id pager, sort saved in
  localStorage). Questions fetched with limit=0 and paged client-side (50); Sets cards paged 10 with search/sort.
  cat_labels(): shared category titles get " · date" (last played, else date in notes, else ID); used in all pickers.
  Set editor hides categories picked in other slots. Header/favicon: /logo.png (brand_logo(), 128 px, no token).
- Importer subtext: keynote_read gives Slide.parts[(main, sub)] per text box from styles (Keynote storage field 5 para
  / field 8 char style tables; CharacterStyleProperties 2=italic 3=size; pptx a:rPr i/sz). Only a TRAILING italic/
  smaller run is sub (mid-sentence emphasis stays). Question box sub -> question_note; answer boxes kept whole.
- Script versioning: tbtgen_sheet.gs SCRIPT_VERSION (returned by "tables") vs build.SCRIPT_VERSION. Version = "YYYY-MM-DD HH:MM";
  run `python3 bump_script_version.py` after EVERY .gs change (stamps .gs first line, .gs SCRIPT_VERSION, build.py). Out-of-date script: terminal WARNING, red banner in the app, Status row, and save errors
  ("not editable"/"unknown action") get the redeploy hint.
- Pictures: build.slide_image() converts formats a .pptx can't hold (AVIF, WebP, HEIC...; pptx = JPEG/PNG/GIF/BMP/TIFF)
  once into build/img-<hash>.png|jpg (Pillow, else sips on a Mac); picture()/fit_box()/video Canvas.picture use it;
  check() reports unreadable pictures via image_problem(). Importer CONVERT includes avif.
- Never commit generated decks (out/), cache (build/) or media (assets/media/).
