# Slide Aid roadmap — future feature ideas

Ranked by value ÷ effort on Mac VBA.

## Done (v3)

Implemented: Advanced Format Painter, Select Similar Shapes, clean-up tools (notes/animations/unused designs), slide summary to clipboard, Paste on Selected Slides, Master-objects toggle, proofing language (FI/EN-US/EN-UK/SV/DE), My Elements library, My Formats, Agenda from sections, and the menu-bar shortcut system (macOS App Shortcuts bindable, user-chosen keys per tool).

## Next — medium effort

**Check presentation (corporate design).** Rule loop over all shapes: font ∉ whitelist, color ∉ theme+palette, text overflowing frame, double spaces. Output a report slide listing violations with slide numbers, or fix-on-click via a "fix all" pass. Start with 3–4 rules.

**Navigation history + Go to Slide.** A class module with `WithEvents App As Application` catching `SlideSelectionChanged` pushes slide indexes onto a stack; Back/Forward buttons pop it. Go-to = InputBox + `ActiveWindow.View.GotoSlide n`.

**Agenda Wizard extensions.** Time slots and responsible columns (store per-section metadata in `Presentation.Tags`), auto page numbers, backup-section handling.

## Tier 3 — feasible but significant effort

**Adjust Pentagon Headers.** Given a pentagon + header box pair, set header width = pentagon width minus arrowhead inset computed from `Adjustments(1)`. Easy code, but only worth it once an Element library with conclusion boxes exists.

**Decompose Tables.** For each cell create a rectangle at the cell's geometry, copy fill/borders/text. The cell-boundary math already exists in `modTableSnap`.

**Insert Selected Slides as Pictures.** `slide.Export path, "PNG"` then `Shapes.AddPicture` in a grid (reuse the Matrix math). Mac sandbox may prompt for file access — export to `~/Library/Containers/com.microsoft.Powerpoint/Data/` to avoid prompts.

**Gantt / project plan.** A full project-plan editor in Mac VBA would be InputBox-driven and painful. If needed: define the plan in Excel/CSV and have a macro that draws bars from it. Recommend deferring.

## Researched options (2026-07)

**Right-click menus DO work on Mac** — not via CommandBars (absent) but via RibbonX `<contextMenus>` in customUI14.xml, per Ron de Bruin's Mac Office documentation (implemented: Slide Aid on shape right-click, Chart Aid on group right-click). Also learned: `imageMso` never renders on Mac (removed), and `CommandBars.ExecuteMso "..."` can trigger built-in ribbon commands from VBA. **Native dialogs from VBA**: the `SlideAidUI` AppleScript helper (via `AppleScriptTask`) hosts rich native UI — `choose color` (Recolor Series) and ASObjC `NSAlert` panels with sliders, checkboxes, and popups (Chart Aid's **Chart Settings** and **Edit Colors**). Learned the hard way: custom Cocoa target/action callbacks do NOT dispatch reliably in that runtime, so every action is an `NSAlert` button and the caller re-invokes the handler for multi-step edits (pick a color, add a swatch); control state is read back after `runModal`. **Office.js keyboard shortcuts now support PowerPoint** (KeyboardShortcuts 1.1 + SharedRuntime) — only relevant if the tools are ever rewritten as a web add-in, since JS actions can't call VBA. **Hammerspoon** could host a real settings panel (hs.dialog.color, webviews) writing our config files — now largely superseded by the native ASObjC Chart Settings / Edit Colors panels, but still an option if a persistent floating panel is ever wanted.

## Not achievable in Mac VBA — don't attempt

Floating Color Bar and screen-wide color picker (no floating windows / no screen access), per-button keyboard shortcuts (PowerPoint VBA has no OnKey), in-canvas drag handles (no VBA mouse events), Office web task panes in VBA (Windows .NET only), automatic update distribution. The nearest Mac equivalents are already built in: ribbon menus and galleries, native ASObjC dialog panels (Chart Settings / Edit Colors), pick-from-Master, and macOS Digital Color Meter.

## Suggested order

1. Advanced Format Painter, Select Similar, clean-up tools (quick wins, daily use)
2. My Elements library (biggest workflow gain)
3. My Formats, Agenda Wizard
4. Check presentation rules
