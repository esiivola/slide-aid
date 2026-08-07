use framework "Foundation"
use framework "AppKit"
use scripting additions

-- Slide Aid UI helper: native macOS dialogs callable from VBA
-- via AppleScriptTask. Install (one time):
--
--   mkdir -p ~/Library/Application\ Scripts/com.microsoft.Powerpoint
--   osacompile -o ~/Library/Application\ Scripts/com.microsoft.Powerpoint/SlideAidUI.scpt apps/powerpoint/tools/SlideAidUI.applescript
--
-- HANDLERS (all return a single string; VBA calls one at a time via
-- AppleScriptTask and blocks until it returns):
--   chooseColor      native color panel (wheel, eyedropper). "R,G,B" 0-255.
--   chartSettings    native slider/checkbox/popup panel for chart parameters.
--   chooseChartColor pick which chart color to change, then the color panel.
--   editColors       native palette editor (swatches + family/slot popups).
--   buildPpam        dev helper (repo build). Fixed command, not a shell runner.
--
-- DESIGN NOTE: the dialogs use ONLY primitives verified to work when the
-- compiled .scpt is invoked via AppleScriptTask from sandboxed PowerPoint:
-- NSAlert's own buttons (coded returns), reading accessory control state
-- AFTER runModal, Cocoa value-bindings for the live slider read-out, and
-- the Standard-Additions `choose color` / `choose from list`. Custom
-- target/action callbacks are deliberately NOT used - they do not dispatch
-- reliably in this runtime - so every action is an NSAlert button and the
-- caller re-invokes the handler for multi-step edits (pick a color, add a
-- swatch, switch family). VBA falls back to on-slide tables if this helper
-- is absent, so the .ppam works without it.

-- buildPpam: dev helper for BuildSlideAid (modImportHelper). Runs the
-- ribbon injector in the repo and converts the freshly saved .pptm to
-- .ppam. Fixed command on purpose - not a generic shell runner.
on buildPpam(buildParam)
	try
		set buildArgs to paragraphs of buildParam
		set repoPath to item 1 of buildArgs
		if (count of buildArgs) > 1 then
			set pptmPath to item 2 of buildArgs
		else
			-- Backward compatibility for a loaded add-in built before
			-- sandboxed temporary builds were introduced.
			set pptmPath to repoPath & "/Slide Aid.pptm"
		end if
		set ppamPath to (text 1 thru -6 of pptmPath) & ".ppam"
		set buildCommand to "PATH=/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH; python3 " & quoted form of (repoPath & "/tools/inject_ribbon.py")
		set buildCommand to buildCommand & " --make-ppam " & quoted form of pptmPath
		set buildCommand to buildCommand & " && mkdir -p " & quoted form of (repoPath & "/dist")
		set buildCommand to buildCommand & " && mv -f " & quoted form of ppamPath & " " & quoted form of (repoPath & "/dist/Slide Aid.ppam") & " 2>&1"
		do shell script buildCommand
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

-- =====================================================================
-- CHART SETTINGS: a native panel of labelled controls, pre-filled with
-- current values. Request = one control per line, tab-separated:
--     #<tab>title<tab>info<tab>(unused)
--     num<tab>key<tab>label<tab>value<tab>min<tab>max
--     check<tab>key<tab>label<tab>value(0|1)
--     popup<tab>key<tab>label<tab>value<tab><tab><tab>opt|opt|opt
--     swatch<tab>key<tab>label<tab>hex           (display only)
-- Return line 1 = OK | APPLY | COLORS | CANCEL, then key=value lines for
-- the num/check/popup controls (colors are held by the caller).
-- =====================================================================
on chartSettings(param)
	set parsed to my parseSpec(param)
	set recs to item 1 of parsed
	set titleText to item 2 of parsed
	set infoText to item 3 of parsed
	set hasColors to item 4 of parsed

	set n to count of recs
	set totalH to (n * 32) + 6
	if totalH < 40 then set totalH to 40
	set theView to current application's NSView's alloc's initWithFrame:(current application's NSMakeRect(0, 0, 480, totalH))
	set readbacks to my addControls(theView, recs, totalH)

	set alert to current application's NSAlert's alloc's init()
	alert's setMessageText:titleText
	if infoText is not "" then alert's setInformativeText:infoText
	alert's setAccessoryView:theView

	set tokens to {}
	alert's addButtonWithTitle:"OK"
	set end of tokens to "OK"
	alert's addButtonWithTitle:"Apply"
	set end of tokens to "APPLY"
	if hasColors then
		alert's addButtonWithTitle:"Colors…"
		set end of tokens to "COLORS"
	end if
	alert's addButtonWithTitle:"Cancel"
	set end of tokens to "CANCEL"

	set tok to my runAlert(alert, tokens)
	if tok is "CANCEL" then return "CANCEL"

	set outLines to my readbackLines(readbacks)
	set AppleScript's text item delimiters to linefeed
	set res to tok & linefeed & (outLines as text)
	set AppleScript's text item delimiters to ""
	return res
end chartSettings

-- chooseChartColor: caller passes a title line then "label<tab>hex" lines.
-- Shows a picker for WHICH color, then the native color panel seeded with
-- its current value. Returns "label<tab>RRGGBB" or "" if cancelled.
on chooseChartColor(param)
	set ls to paragraphs of param
	if (count of ls) < 2 then return ""
	set titleText to item 1 of ls
	set labels to {}
	set hexes to {}
	repeat with i from 2 to count of ls
		set ln to (item i of ls) as text
		if ln is not "" then
			set AppleScript's text item delimiters to tab
			set f to text items of ln
			set AppleScript's text item delimiters to ""
			if (count of f) ≥ 2 then
				set end of labels to (item 1 of f)
				set end of hexes to (item 2 of f)
			end if
		end if
	end repeat
	if labels is {} then return ""

	if (count of labels) = 1 then
		set pick to item 1 of labels
	else
		set chosen to (choose from list labels with title "Chart Aid" with prompt titleText)
		if chosen is false then return ""
		set pick to item 1 of chosen
	end if

	set seedHex to "808080"
	repeat with i from 1 to count of labels
		if (item i of labels) is pick then set seedHex to (item i of hexes)
	end repeat
	set rgb to my hexToRGBList(seedHex)
	try
		set c to choose color default color {(item 1 of rgb) * 257, (item 2 of rgb) * 257, (item 3 of rgb) * 257}
	on error
		return ""
	end try
	set hexOut to my rgbToHex((item 1 of c) div 257, (item 2 of c) div 257, (item 3 of c) div 257)
	set AppleScript's text item delimiters to tab
	set res to pick & tab & hexOut
	set AppleScript's text item delimiters to ""
	return res
end chooseChartColor

-- =====================================================================
-- EDIT COLORS: native palette editor for one chart family. Request:
--     #<tab>title<tab>family<tab>fam1|fam2|fam3
--     swatch<tab>idx<tab>hex   (one per palette color)
-- Family and a "color / action" popup are read back on every button.
-- Return line 1 = CHANGE | DONE | CANCEL, then FAMILY=..., SLOT=...
-- (SLOT is "Color N" or an action item). The caller mutates its palette
-- and re-invokes until DONE.
-- =====================================================================
on editColors(param)
	set ls to paragraphs of param
	set titleText to "Chart colors"
	set family to "BARS"
	set families to {"BARS", "LINES", "PIES"}
	set hexes to {}
	repeat with i from 1 to count of ls
		set ln to (item i of ls) as text
		if ln is not "" then
			set AppleScript's text item delimiters to tab
			set f to text items of ln
			set AppleScript's text item delimiters to ""
			if (item 1 of f) is "#" then
				if (count of f) ≥ 2 then set titleText to item 2 of f
				if (count of f) ≥ 3 then set family to item 3 of f
				if (count of f) ≥ 4 then
					set AppleScript's text item delimiters to "|"
					set families to text items of (item 4 of f)
					set AppleScript's text item delimiters to ""
				end if
			else if (item 1 of f) is "swatch" then
				if (count of f) ≥ 3 then set end of hexes to (item 3 of f)
			end if
		end if
	end repeat

	set slotItems to {}
	repeat with i from 1 to count of hexes
		set end of slotItems to "Color " & i
	end repeat
	set end of slotItems to "+ Add a color"
	set end of slotItems to "- Remove last color"
	set end of slotItems to "* Reset to theme colors"

	set nRows to 2 + (count of hexes)
	set rowH to 30
	set totalH to (nRows * rowH) + 6
	set theView to current application's NSView's alloc's initWithFrame:(current application's NSMakeRect(0, 0, 430, totalH))

	set yTop to totalH - rowH
	my addLabel(theView, "Chart family:", 0, yTop + 2, 120)
	set famPop to current application's NSPopUpButton's alloc's initWithFrame:(current application's NSMakeRect(128, yTop, 210, 24)) pullsDown:false
	famPop's addItemsWithTitles:families
	famPop's selectItemWithTitle:family
	theView's addSubview:famPop

	set yTop2 to totalH - (2 * rowH)
	my addLabel(theView, "Color / action:", 0, yTop2 + 2, 120)
	set slotPop to current application's NSPopUpButton's alloc's initWithFrame:(current application's NSMakeRect(128, yTop2, 210, 24)) pullsDown:false
	slotPop's addItemsWithTitles:slotItems
	theView's addSubview:slotPop

	repeat with i from 1 to count of hexes
		set yy to totalH - ((2 + i) * rowH)
		my addLabel(theView, "Color " & i, 0, yy + 2, 80)
		my addSwatch(theView, (item i of hexes), 90, yy + 2)
		my addLabel(theView, (item i of hexes), 160, yy + 2, 120)
	end repeat

	set alert to current application's NSAlert's alloc's init()
	alert's setMessageText:titleText
	alert's setInformativeText:"Pick a color or action, then Apply Change. Done writes the palette."
	alert's setAccessoryView:theView
	set tokens to {}
	alert's addButtonWithTitle:"Apply Change"
	set end of tokens to "CHANGE"
	alert's addButtonWithTitle:"Done"
	set end of tokens to "DONE"
	alert's addButtonWithTitle:"Cancel"
	set end of tokens to "CANCEL"

	set tok to my runAlert(alert, tokens)
	if tok is "CANCEL" then return "CANCEL"

	set famSel to (famPop's titleOfSelectedItem()) as text
	set slotSel to (slotPop's titleOfSelectedItem()) as text
	set AppleScript's text item delimiters to linefeed
	set res to tok & linefeed & "FAMILY=" & famSel & linefeed & "SLOT=" & slotSel
	set AppleScript's text item delimiters to ""
	return res
end editColors

-- ---------- shared UI helpers ----------

on runAlert(alert, tokens)
	try
		(current application's NSApplication's sharedApplication())'s activateIgnoringOtherApps:true
	end try
	try
		(alert's |window|())'s setLevel:3
	end try
	set code to (alert's runModal()) as integer
	set idx to code - 999
	if idx < 1 or idx > (count of tokens) then return "CANCEL"
	return item idx of tokens
end runAlert

on parseSpec(param)
	set recs to {}
	set titleText to "Chart settings"
	set infoText to ""
	set hasColors to false
	repeat with rawLn in (paragraphs of param)
		set ln to rawLn as text
		if ln is not "" then
			set AppleScript's text item delimiters to tab
			set f to text items of ln
			set AppleScript's text item delimiters to ""
			set tg to item 1 of f
			if tg is "#" then
				if (count of f) ≥ 2 then set titleText to item 2 of f
				if (count of f) ≥ 3 then set infoText to item 3 of f
			else
				set kk to ""
				set lbl to ""
				set vl to ""
				set mn to 0
				set mx to 100
				set opts to {}
				if (count of f) ≥ 2 then set kk to item 2 of f
				if (count of f) ≥ 3 then set lbl to item 3 of f
				if (count of f) ≥ 4 then set vl to item 4 of f
				if (count of f) ≥ 5 then
					try
						set mn to (item 5 of f) as number
					end try
				end if
				if (count of f) ≥ 6 then
					try
						set mx to (item 6 of f) as number
					end try
				end if
				if (count of f) ≥ 7 then
					set AppleScript's text item delimiters to "|"
					set opts to text items of (item 7 of f)
					set AppleScript's text item delimiters to ""
				end if
				if tg is "swatch" then set hasColors to true
				set end of recs to {ctrl:tg, k:kk, lbl:lbl, vl:vl, mn:mn, mx:mx, opts:opts}
			end if
		end if
	end repeat
	return {recs, titleText, infoText, hasColors}
end parseSpec

on addControls(theView, recs, totalH)
	set readbacks to {}
	set n to count of recs
	repeat with i from 1 to n
		set r to item i of recs
		set yTop to totalH - (i * 32)
		set c to ctrl of r
		if c is "num" then
			my addLabel(theView, (lbl of r), 0, yTop + 4, 160)
			set sld to current application's NSSlider's alloc's initWithFrame:(current application's NSMakeRect(168, yTop + 2, 220, 20))
			sld's setMinValue:(mn of r)
			sld's setMaxValue:(mx of r)
			try
				sld's setDoubleValue:((vl of r) as number)
			end try
			theView's addSubview:sld
			set ro to current application's NSTextField's labelWithString:"0"
			ro's setFrame:(current application's NSMakeRect(394, yTop + 4, 80, 18))
			set fmt to current application's NSNumberFormatter's alloc's init()
			fmt's setMaximumFractionDigits:0
			ro's setFormatter:fmt
			ro's bind:"value" toObject:sld withKeyPath:"doubleValue" options:(missing value)
			theView's addSubview:ro
			set end of readbacks to {ctrl:"num", k:(k of r), ctl:sld}
		else if c is "check" then
			set cb to current application's NSButton's alloc's initWithFrame:(current application's NSMakeRect(0, yTop + 2, 470, 20))
			cb's setButtonType:3
			cb's setTitle:(lbl of r)
			if (vl of r) is "1" then
				cb's setState:1
			else
				cb's setState:0
			end if
			theView's addSubview:cb
			set end of readbacks to {ctrl:"check", k:(k of r), ctl:cb}
		else if c is "popup" then
			my addLabel(theView, (lbl of r), 0, yTop + 4, 160)
			set pop to current application's NSPopUpButton's alloc's initWithFrame:(current application's NSMakeRect(168, yTop + 2, 220, 24)) pullsDown:false
			pop's addItemsWithTitles:(opts of r)
			pop's selectItemWithTitle:(vl of r)
			theView's addSubview:pop
			set end of readbacks to {ctrl:"popup", k:(k of r), ctl:pop}
		else if c is "swatch" then
			my addLabel(theView, (lbl of r), 0, yTop + 4, 160)
			my addSwatch(theView, (vl of r), 168, yTop + 4)
			my addLabel(theView, (vl of r), 236, yTop + 4, 120)
		end if
	end repeat
	return readbacks
end addControls

on readbackLines(readbacks)
	set out to {}
	repeat with rb in readbacks
		set c to ctrl of rb
		set kk to k of rb
		set ctl to ctl of rb
		if c is "num" then
			set end of out to kk & "=" & ((round (ctl's doubleValue())) as integer)
		else if c is "check" then
			set end of out to kk & "=" & ((ctl's state()) as integer)
		else if c is "popup" then
			set end of out to kk & "=" & ((ctl's titleOfSelectedItem()) as text)
		end if
	end repeat
	return out
end readbackLines

on addLabel(theView, t, x, y, w)
	set lf to current application's NSTextField's labelWithString:t
	lf's setFrame:(current application's NSMakeRect(x, y, w, 18))
	theView's addSubview:lf
	return lf
end addLabel

-- NSBox's setFillColor:/setBorderColor: take an NSColor directly, so we
-- avoid the CGColor bridging that a layer-backed NSView chokes on
-- ("value must be missing value or reference" -> swatch draws empty).
on addSwatch(theView, hx, x, y)
	set sw to current application's NSBox's alloc's initWithFrame:(current application's NSMakeRect(x, y, 60, 18))
	sw's setBoxType:4        -- NSBoxCustom
	sw's setTitlePosition:0  -- NSNoTitle
	sw's setCornerRadius:2.0
	sw's setBorderWidth:0.5
	sw's setBorderColor:(current application's NSColor's grayColor())
	sw's setFillColor:(my hexToNSColor(hx))
	theView's addSubview:sw
	return sw
end addSwatch

-- ---------- hex helpers ----------

on hexToRGBList(hx)
	try
		set u to (current application's NSString's stringWithString:hx)'s uppercaseString() as text
		if (count of u) < 6 then error "short"
		return {my hex2(text 1 thru 2 of u), my hex2(text 3 thru 4 of u), my hex2(text 5 thru 6 of u)}
	on error
		return {128, 128, 128}
	end try
end hexToRGBList

on hexToNSColor(hx)
	set rgb to my hexToRGBList(hx)
	return current application's NSColor's colorWithCalibratedRed:((item 1 of rgb) / 255) green:((item 2 of rgb) / 255) blue:((item 3 of rgb) / 255) alpha:1.0
end hexToNSColor

on hex2(twoChars)
	set digits to "0123456789ABCDEF"
	set c1 to (offset of (character 1 of twoChars) in digits) - 1
	set c2 to (offset of (character 2 of twoChars) in digits) - 1
	if c1 < 0 then set c1 to 0
	if c2 < 0 then set c2 to 0
	return c1 * 16 + c2
end hex2

on rgbToHex(r, g, b)
	return my toHex2(r) & my toHex2(g) & my toHex2(b)
end rgbToHex

on toHex2(n)
	set digits to "0123456789ABCDEF"
	set n to n as integer
	if n < 0 then set n to 0
	if n > 255 then set n to 255
	return (character ((n div 16) + 1) of digits) & (character ((n mod 16) + 1) of digits)
end toHex2

-- Headless self-test: build the accessory + read back initial values
-- WITHOUT showing the modal (used by the repo's compile/verify step).
on chartSettingsSelfTest(param)
	set parsed to my parseSpec(param)
	set recs to item 1 of parsed
	set n to count of recs
	set totalH to (n * 32) + 6
	if totalH < 40 then set totalH to 40
	set theView to current application's NSView's alloc's initWithFrame:(current application's NSMakeRect(0, 0, 480, totalH))
	set readbacks to my addControls(theView, recs, totalH)
	set outLines to my readbackLines(readbacks)
	set AppleScript's text item delimiters to linefeed
	set res to "OK" & linefeed & (outLines as text)
	set AppleScript's text item delimiters to ""
	return res
end chartSettingsSelfTest
