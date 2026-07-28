-- Slide Aid UI helper: native macOS dialogs callable from VBA
-- via AppleScriptTask. Install (one time):
--
--   mkdir -p ~/Library/Application\ Scripts/com.microsoft.Powerpoint
--   osacompile -o ~/Library/Application\ Scripts/com.microsoft.Powerpoint/SlideAidUI.scpt tools/SlideAidUI.applescript
--
-- chooseColor: opens the native macOS color panel (wheel, sliders,
-- swatches, eyedropper magnifier). Param and result are "R,G,B"
-- with 0-255 components; returns "" if the user cancels.

-- buildPpam: dev helper for BuildSlideAid (modImportHelper). Runs the
-- ribbon injector in the repo and converts the freshly saved .pptm to
-- .ppam. Fixed command on purpose - not a generic shell runner.
on buildPpam(repoPath)
	try
		do shell script "PATH=/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH; cd " & quoted form of repoPath & " && python3 tools/inject_ribbon.py --make-ppam 'Slide Aid.pptm' 2>&1"
		return "OK"
	on error errMsg
		return "ERR: " & errMsg
	end try
end buildPpam

on chooseColor(param)
	set AppleScript's text item delimiters to ","
	set parts to text items of param
	set r to ((item 1 of parts) as integer) * 257
	set g to ((item 2 of parts) as integer) * 257
	set b to ((item 3 of parts) as integer) * 257
	try
		set c to choose color default color {r, g, b}
	on error
		return ""
	end try
	set rr to (item 1 of c) div 257
	set gg to (item 2 of c) div 257
	set bb to (item 3 of c) div 257
	return (rr as text) & "," & (gg as text) & "," & (bb as text)
end chooseColor
