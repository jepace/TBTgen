-- Opens generated .pptx decks in Keynote, fixes the auto-advance timing, and saves them as .key files.
--
-- Usage (Terminal, on the Mac):
--   osascript keynote_finish.applescript "/path/to/2026-11-23 101 Wine Press.pptx" [more.pptx ...]
--
-- Why: when Keynote imports a timed PowerPoint slide, it starts counting the transition delay only AFTER
-- the slide's builds finish. Our question slides already have a build lasting the full question time
-- (the logo countdown), so the delay must be ~0 or the slide waits twice as long.

on run argv
	if (count of argv) is 0 then
		return "Give one or more .pptx paths."
	end if
	set report to ""
	repeat with p in argv
		set pptxPath to p as text
		set keyPath to (text 1 thru -6 of pptxPath) & ".key"
		set inFile to POSIX file pptxPath
		set outFile to POSIX file keyPath
		tell application "Keynote"
			activate
			open inFile
			delay 2
			set doc to front document
			set fixed to 0
			repeat with s in slides of doc
				set tp to transition properties of s
				if automatic transition of tp then
					set transition properties of s to {automatic transition:true, transition delay:0.1}
					set fixed to fixed + 1
				end if
			end repeat
			save doc in outFile
			close doc saving no
		end tell
		set report to report & "saved " & keyPath & " (" & fixed & " timed slides fixed)" & linefeed
	end repeat
	return report
end run
