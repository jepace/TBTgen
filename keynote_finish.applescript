-- Opens generated .pptx decks in Keynote, fixes the auto-advance timing, and saves them as .key files.
--
-- Which Keynote: Keynote Creator Studio (app id com.apple.Keynote) when installed, else the classic Keynote
-- (com.apple.iWork.Keynote). Works with either one alone. To force one:
--   TBT_KEYNOTE=com.apple.iWork.Keynote osascript keynote_finish.applescript ...
-- Usage (Terminal, on the Mac):
--   osascript keynote_finish.applescript "/path/to/2026-11-23 101 Wine Press.pptx" [more.pptx ...]
--
-- Why: when Keynote imports a timed PowerPoint slide, it starts counting the transition delay only AFTER
-- the slide's builds finish. Our question slides already have a build lasting the full question time
-- (the logo countdown), so the delay must be ~0 or the slide waits twice as long.

on installed(appId)
	try
		do shell script "mdfind \"kMDItemCFBundleIdentifier == '" & appId & "'\" | grep -q ."
		return true
	on error
		return false
	end try
end installed

on run argv
	if (count of argv) is 0 then
		return "Give one or more .pptx paths."
	end if
	set appId to system attribute "TBT_KEYNOTE"
	if appId is "" then
		if installed("com.apple.Keynote") then
			set appId to "com.apple.Keynote"
		else if installed("com.apple.iWork.Keynote") then
			set appId to "com.apple.iWork.Keynote"
		else
			return "Keynote isn't installed."
		end if
	end if
	set report to "Using " & (name of application id appId) & " (" & appId & ")" & linefeed

	-- The part that talks to Keynote is compiled only now, for the app chosen above, so this script doesn't
	-- need any particular Keynote to be installed.
	set q to quote
	set worker to "on run {pptxPath, keyPath}
	-- open it the way Finder does: a sandboxed Keynote may refuse a file path handed to it by a script
	do shell script " & q & "open -b " & appId & " " & q & " & quoted form of pptxPath
	tell application id " & q & appId & q & "
		activate
		set waited to 0
		repeat until (count documents) > 0
			delay 1
			set waited to waited + 1
			if waited > 180 then error " & q & "Keynote didn't open the deck (a dialog may be waiting): " & q & " & pptxPath
		end repeat
		delay 2
		set doc to front document
		set fixedCount to 0
		repeat with s in slides of doc
			set tp to transition properties of s
			if automatic transition of tp then
				set transition properties of s to {automatic transition:true, transition delay:0.1}
				set fixedCount to fixedCount + 1
			end if
		end repeat
		save doc in (POSIX file keyPath)
		close doc saving no
	end tell
	return fixedCount
end run"

	-- start Keynote and wait until it is ready (opening a file too early fails with "isn't running")
	tell application id appId to launch
	repeat 60 times
		if application id appId is running then exit repeat
		delay 0.5
	end repeat
	delay 2
	repeat with p in argv
		set pptxPath to p as text
		set keyPath to (text 1 thru -6 of pptxPath) & ".key"
		try
			do shell script "test -f " & quoted form of pptxPath
		on error
			return report & "Not found: " & pptxPath
		end try
		set fixedCount to run script worker with parameters {pptxPath, keyPath}
		set report to report & "saved " & keyPath & " (" & fixedCount & " timed slides fixed)" & linefeed
	end repeat
	return report
end run
