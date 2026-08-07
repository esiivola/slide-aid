CHART AID — CONTEXT & GOAL

GOAL
Make Chart Aid's recoloring, color picking, chart-parameter changing, and the
rebuild-after-change flow simple, with genuinely good UI/UX. Parameter changing
is especially hacky today. Benchmark good tools, think about the ideal
interaction, and decide the approach yourself.

WHAT CHART AID IS
A VBA add-in tab in PowerPoint (Mac-first) that builds editable, SHAPE-BASED
charts from a PowerPoint table (column, bar, stacked, 100%, waterfall, Mekko,
line, area, pie, doughnut, scatter/bubble, Gantt). Charts are native shapes (not
embedded Office charts) and carry their own data so they can be rebuilt/restyled.
Keep it native and offline — ships in one .ppam, no live server.

HOW THE RELEVANT FLOWS WORK NOW (and what's hacky)
- Build: put data in a PowerPoint table, click a chart type.
- Change data: "Edit Data" drops a data table by the chart; edit numbers; select
  table+chart; "Rebuild".
- Change parameters (the hacky part): "Edit Settings (Table)" inserts a table of
  cryptic keys with current values (ClusterFill=0.72, StackFill=0.65,
  LabelSizePt=9, Decimals=auto, ValueLabels=1, TotalLabels=1, Legend=1,
  MekkoGapPt=2, MarkerSizePt=5, PlotWidthCm=12, ...); you type magic numbers,
  select, "Apply from Selection", eyeball, repeat. It litters the slide and has
  no preview and no direct manipulation.
- Recolor / color picking: a per-family palette; "Edit Palettes (Swatches)" drops
  recolorable swatch rows you recolor then Apply; "Recolor Series" opens the real
  macOS color panel via an AppleScript helper; "Color Themes" is a ribbon gallery
  (the one nice control today).
- After any change you must Rebuild/Restyle to see the result.

PLATFORM REALITY (design within these; verify, but they're known)
- Mac PowerPoint VBA: no reliable UserForms (can't pop a custom dialog with
  sliders) and no mouse/drag events (can't do drag-the-bar handles) — that's why
  today's UI is on-slide tables.
- Ribbon (customUI14.xml): buttons/toggles/menus/galleries work on Mac; free-text
  and numeric spin boxes are unreliable on Mac.
- Native macOS dialogs via AppleScript (tools/SlideAidUI.scpt) work — already used
  for the color panel.
- A web task-pane add-in can host rich HTML (sliders, pickers, live preview), but
  Office.js has no chart/freeform API (drawing stays in VBA) and sideloaded web
  add-ins don't load reliably on this Mac without M365 central deployment — so a
  web panel would need a VBA handshake and a deployment story.
- Rebuild is cheap (data travels with the chart), but each rebuild is many undo
  steps.

WHAT GOOD LOOKS LIKE
Recolor, pick colors, tweak parameters, and rebuild should feel simple and
discoverable, ideally with quick/live feedback, and without typing magic numbers
into a helper table — while staying native, offline, and shareable via the .ppam.

FILE MAP (repo root = slide-aid)
 apps/powerpoint/ribbon/customUI14.xml   Chart Aid ribbon
 apps/powerpoint/src/modChartStyle.bas   params/settings + palette + Apply — START HERE
 apps/powerpoint/src/modChartCore.bas    build orchestration + shared drawing
 apps/powerpoint/src/modChartBars.bas / modChartLines.bas / modChartGantt.bas
 apps/powerpoint/src/modChartEdit.bas    Edit Data / Rebuild
 apps/powerpoint/src/modChartAnno.bas    annotations
 apps/powerpoint/src/modRibbon.bas       ribbon tag -> sub dispatch
 apps/powerpoint/tools/SlideAidUI.applescript   native macOS helper (color panel)
 docs/CHARTS.md, docs/POWERPOINT_UI_REFERENCE.md   layouts + UI reference
 apps/powerpoint-iconaid/                web task-pane precedent

Benchmark tools you consider best-in-class, read the code (start at
modChartStyle.bas), and decide the UI/UX yourself.
