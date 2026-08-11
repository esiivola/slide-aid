# Slide Aid for Google Slides

This folder contains the Google Slides companion to Slide Aid. The PowerPoint add-in remains the main and canonical version. Names, order, icons, defaults, and chart semantics come from the [PowerPoint UI reference](../../docs/POWERPOINT_UI_REFERENCE.md), the [chart layout reference](../../docs/CHARTS.md), the [RibbonX definition](../powerpoint/ribbon/customUI14.xml), and the original [PowerPoint README](../../README.md).

This is a TypeScript Google Workspace editor add-on built with Apps Script and `clasp`. It opens an HTML sidebar, operates on the current Slides selection, and deliberately separates personal user settings from shared presentation state.

## Sidebar structure

The sidebar mirrors the ribbon rather than inventing its own arrangement. Its
four sections are the two PowerPoint ribbon tabs plus the two areas that have no
ribbon equivalent:

| Sidebar section | Mirrors | Groups, in ribbon order |
|---|---|---|
| **Slide Aid** | Slide Aid tab | Wizards, Position, Dock, Arrange, Place on Slide, Size, Stretch, Fill Gap, Shape, Color, Text, View & Expert |
| **Chart Aid** | Chart Aid tab | Charts, Data, Style, Annotations, Elements |
| **Icons** | Insert → Insert Icons | Icons |
| **Deck** | *(Google Slides only)* | Named layouts, Matrix preview, My Elements library, Deck QA |

Control names, defaults, and group order come from the ribbon. Where PowerPoint
uses a dropdown menu, the sidebar uses a small button row or a select; where it
uses a supertip, the sidebar uses a tooltip with the same explanation. The pinned
reference and the selection inspector sit in the header so they stay visible from
every section, because the reference is the concept every position tool depends on.

## Implemented

- An explicit **Set reference** workflow to replace PowerPoint's last-selected Master convention, with pinned/slide/selection-bounds modes.
- Align, dock, stretch, and fill gap in all four directions; distribute H and V; match width/height/both; Golden Canon; Swap with corner anchoring and optional sizes.
- Stack, exact spacing, one-click near-square Matrix and custom Matrix, all twelve place-on-slide regions, slice, multiply, and a Magic Resizer that scales font sizes with the geometry.
- Format Painter, My Formats, Select Similar (type / fill / both), Process Chain, Snap to Table, and an Agenda generator.
- RGB, theme-linked, and the eight named palette colors for fill, border, and text; theme/RGB conversion; Pick from Master; Color Info; optional palette-to-theme-accent updates.
- Fit to Text, Split at Cursor, Merge Boxes, case conversion, double-space cleanup, and text swap.
- Hide Objects / Unhide All, Paste on Slides, Remove Speaker Notes, and Copy Slide Summary.
- All Chart Aid chart families: column, bar, stacked column, stacked bar, 100%, Waterfall, Mekko, line, area, pie, doughnut, scatter, bubble, and Gantt — with value labels, stacked totals, and a series legend on by default, matching PowerPoint's defaults.
- **Chart Settings** and **Edit Colors** as sidebar panels: the same sixteen parameters with the same per-chart-type override rule (`COL.ClusterFill` beats `ClusterFill`), and the same three palette families (Bars / Lines / Pies). See [`shared/specs/chart-style.json`](../../shared/specs/chart-style.json).
- Chart annotations — Difference, % Difference, CAGR, Value Line, Average Line — computed from each bar's stored datum, never from its pixel size. Plus Harvey balls, checkboxes, and Cycle State.
- **Data Layouts** help and **Sample Slides**: fourteen example slides built by the real chart code, so the examples cannot drift from the builders.
- Atomic shape-native chart construction through the Advanced Slides service, with revision checks to avoid partial charts.
- Accessible chart alt text plus chunked shared metadata for Edit Data and Rebuild; old v0.1 alt-text JSON migrates automatically when a chart is selected or rebuilt.
- Nine canonical shared deck palettes, restyle selected/all, and persistent per-series recoloring. Any style change that leaves charts stale offers a deck-wide restyle, as PowerPoint does.
- Read-only Google Sheets connections: build from an explicit spreadsheet URL, tab, and range, then refresh the selected chart from current values.
- A live selection inspector, command search across every button, matrix preview, and named layouts shared with deck collaborators.
- Shared element libraries backed by one explicitly configured Slides presentation. Slide Aid never searches the user's Drive.
- An **Icons** panel carrying the same 54,250-icon library as PowerPoint plus the reviewed consulting set, with the task pane's concept search, set and category filters, color selection, and incremental loading. Icons insert as pictures and **Make Editable** converts them to native shapes, matching PowerPoint's flow.
- Deck QA reporting for off-slide objects, tiny text, missing alt text, low contrast, fixed RGB fills, stale/broken Sheet sources, orphan datasheets, and irregular horizontal spacing. Safe mechanical issues have targeted fixes.
- The PowerPoint PNG icons are embedded into the built sidebar at build time, so no public image host is required.

## Deliberate differences and current gaps

Google Slides does not expose ordinary click/shift-click selection order. Reference-based tools therefore use a pinned reference, and order-dependent tools use deterministic spatial order: horizontal operations sort left-to-right, vertical operations sort top-to-bottom, and Matrix sorts left-to-right. The selection API behavior is documented in Google's [Slides selection guide](https://developers.google.com/apps-script/guides/slides/selecting).

Google Workspace add-ons use menus, dialogs, and HTML sidebars rather than Office RibbonX or shape context menus. Slide Aid therefore appears under **Extensions → Slide Aid → Open Slide Aid** and runs from the sidebar. Google documents this model in [Dialogs and sidebars](https://developers.google.com/apps-script/guides/dialogs).

Area, pie, and doughnut charts are server-rendered with the Apps Script Charts service and inserted as static images. They retain Slide Aid metadata and remain editable through **Edit data → Rebuild**, but individual segments are not native Slides shapes. Google documents server-side image rendering in the [Charts service reference](https://developers.google.com/apps-script/reference/charts/).

The shared Slides library replaces the local PowerPoint My Elements store, but it is intentionally URL-based rather than a Drive-wide browser. Google Slides still lacks reliable click-order tracking and add-on keyboard-shortcut hooks.

The Icons panel carries the same library as PowerPoint — 54,250 icons from 14
permissively licensed libraries, plus the 70-icon reviewed consulting set — and
follows PowerPoint's two-step model: clicking inserts a picture, and **Make
Editable** turns it into real shapes. Inserting every vector segment directly on
each click would make complex icons unwieldy, so the picture remains the default.

Both steps are exact. The sidebar rasterises the icon from the same normalized
path data the add-in uses, so the picture matches the thumbnail; **Make Editable**
then flattens those paths into native Slides geometry occupying the picture's own
box. Stroke icons become line runs. Solid icons — all of Bootstrap, plus any
`-solid`/`-mini` variant — are scan-converted into filled slices under SVG's
even-odd rule, because Slides cannot fill an arbitrary outline; holes stay open,
so a Bootstrap "0 circle" converts to a ring rather than a disc. The consulting
icons carry their own primitives and go straight in as shapes, skipping the
picture step.

The delivery split is a Slides constraint worth knowing about: an Apps Script
sidebar has no HTTP cache to lean on the way the Office task pane does, so
shipping 4.6 MB of path data inside it would re-download on every open. Metadata
(~1 MB, ~130 KB gzipped) is embedded so search stays instant, and path data ships
as numbered project files the server reads on demand. Icons are sorted by id and
sharded into contiguous ranges, so either side resolves an id to its shard with a
binary search over a small boundary table. `scripts/build_iconaid_web.py` emits
all of it — the Office task pane catalog, the add-in's `icons.dat`, and the Slides
index and shards — from one set of normalized paths, so the three cannot drift.

Google Slides still does not accept SVG through its image API, and has no
freeform-path API; the flattening above is what stands in for both.

**Chart Settings** and **Edit Colors** are HTML panels in the sidebar rather than
the native macOS panels PowerPoint drives through its `SlideAidUI` AppleScript
helper. They read and write the same parameters with the same semantics; the
sidebar simply needs no helper install and no on-slide table fallback.

**Value Line** and **Average Line** work on the chart kinds drawn against one linear
value axis: column, bar, stacked, stacked bar, 100%, waterfall, and line. Mekko
normalizes per column, scatter and bubble have two axes, Gantt is a time band, and
area/pie/doughnut are images — none of them has a single scale for a line to sit on,
and the sidebar says so rather than drawing something meaningless. Difference,
% Difference, CAGR and Average Line work wherever bars carry data, which is every
shape-native kind. A chart built before this version needs one **Rebuild** so its
value scale is recorded.

Harvey balls are built by fanning thin rotated slivers out of the circle's centre,
because the Slides API exposes neither shape adjustment handles nor freeform paths,
so an adjustable pie wedge cannot be made from a preset shape. Any percentage
renders correctly; the arc is a 24-sided approximation.

Hidden objects are parked far off-canvas with their real position recorded in alt
text, since Slides has no per-object visibility flag. Position and layer order
survive, and **Unhide All** restores them.

The following PowerPoint features remain unavailable, each blocked by a missing
Slides API rather than by effort: **Set Margins** and **Wrap Text** (no per-side
text insets or wrap flag), **Block Arrows** and **Rounded Rect.** (no shape
adjustment handles), **Master Objects** (no background-graphics toggle), **Remove
All Animations** and **Delete Unused Designs** (no animation or design API),
**Language** (no proofing-language property), and **Shortcuts** (add-ons cannot
bind keys). Split at Cursor and Merge Boxes are implemented but move plain text:
Slides has no formatting-preserving text-range copy. See the [PowerPoint UI
reference](../../docs/POWERPOINT_UI_REFERENCE.md) for the canonical behavior of each.

## Permissions and stored data

The add-on requests:

- `presentations`: required for atomic Slides API updates and for copying elements to or from an explicitly configured library presentation.
- `spreadsheets.readonly`: required only when building or refreshing a linked-data chart.
- `script.container.ui`: required for the editor menu and sidebar.

It does **not** request a Drive scope and cannot list or search Drive files. Users must paste the exact Google Slides library URL and Google Sheets source URL.

Personal values such as the pinned reference and preferred spacing live in User Properties. Deck palettes, chart style parameters, per-family palettes, named layouts, saved formats, library configuration, and chart payloads live in Document Properties shared by collaborators. Chart data is split into Unicode-safe chunks below the Apps Script per-property size limit.

## Install

Everything below runs from `apps/google-slides`. You need **Node.js 20 or newer**
and a Google account allowed to create Apps Script projects. Nothing is installed
system-wide, and no Drive access is requested at any point.

### 1. Build

```bash
cd apps/google-slides
npm install
npm run build
```

`dist/` is the whole deployment — nothing else is uploaded:

```text
dist/Code.js            the add-on
dist/Sidebar.html       the sidebar, with the icon search index embedded
dist/appsscript.json    manifest and OAuth scopes
dist/IconShards.html    icon shard boundary table
dist/IconPaths00..22    icon path data, read on demand
```

To check your changes first: `npm run check` (types) and `npm test` (71 tests).

### 2. Connect an Apps Script project

Authenticate once:

```bash
npx clasp login
```

Then either create a project:

```bash
npx clasp create --type standalone --title "Slide Aid for Google Slides" --rootDir dist
```

…or point at an existing one by copying `.clasp.example.json` to `.clasp.json`,
replacing `YOUR_SCRIPT_ID`, and keeping `rootDir` as `dist`.

### 3. Push

```bash
npm run push
```

This rebuilds and uploads. It is ~6 MB, most of it icon data, so the first push
takes a moment.

### 4. Install it into Google Slides

In the Apps Script editor (`npm run open`):

1. Confirm **Services** lists **Google Slides API** (`Slides`, v1). `appsscript.json`
   declares it; if the project uses a custom Google Cloud project, enable the
   Google Slides API there too. Charts and icons both need it — without it the
   sidebar loads but every build fails.
2. **Deploy → Test deployments → Install**, choosing **Editor add-on** with Google
   Slides as the host.
3. Open or reload a Google Slides presentation.
4. **Extensions → Slide Aid → Open Slide Aid**, then authorize presentation,
   read-only spreadsheet, and container-UI access.

### 5. Check it works

Worth doing once, since these three paths exercise most of the add-on:

- **Position** — draw two boxes, select one, **Set reference**, select the other, **Left**.
- **Chart Aid** — insert a 3x3 table (blank top-left, categories across, series down),
  select it, click **Column**. You should get bars with value labels and a legend.
- **Icons** — search `strategy`, click an icon, then **Make Editable**.

### Updating

`npm run push` again, then reload the presentation. A test deployment tracks the
latest push, so there is nothing to reinstall. For organization-wide distribution,
create a versioned deployment and configure a private Google Workspace Marketplace
SDK listing.

### If something fails

| Symptom | Cause |
|---|---|
| "The Advanced Slides service is not enabled for this deployment." | Step 4.1 — the Slides API service is not added to the project. |
| "The icon catalog is missing from this deployment." | `dist` was pushed without the `IconPaths*`/`IconShards` files. Re-run `npm run build`, then `npm run push`. |
| Sidebar opens but every command errors | Authorization was declined. Remove the test deployment and reinstall to be re-prompted. |
| Charts build but look unstyled | A **Color Theme** change needs **Restyle All** to reach existing charts. |

## Use

For Master-based tools, select one object and click **Set reference**. Then select the targets and run a command. The reference is keyed to both presentation and slide, and Slide Aid refuses to use it on a different slide.

For charts, create and select a Google Slides table in one of the formats defined in [Chart layouts](../../docs/CHARTS.md), then choose a chart type. **Data Layouts** shows the expected table for each type in the sidebar, and **Sample Slides** appends one working example per type. Select a generated chart or one of its elements to edit, rebuild, restyle, or recolor it.

To style charts, open **Chart Settings**. With a chart selected you edit that chart's own parameters and **Apply** rebuilds it in place; with nothing selected you edit the defaults new charts use. **Edit Colors** edits one palette family at a time — Bars, Lines, or Pies — and **Restyle All** applies it across the deck. Picking a **Color Theme** sets all three families at once and offers to restyle every existing chart.

For annotations, click into a chart to enter the group, select one bar, then shift-click a second, and choose **Difference**, **% Difference**, or **CAGR**. **Average Line** uses the selected bars; **Value Line** asks for a value and draws on the chart's own scale. Every number comes from the bar's stored datum, not from its size on the slide. A chart built before this version needs one **Rebuild** so Slide Aid records its value scale.

For a linked chart, open **Charts → Linked Google Sheets data**, paste the exact spreadsheet URL, enter its tab and A1 range, and choose a chart type. **Refresh selected** reads the range again and rebuilds the selected chart. Rebuild also refreshes linked data automatically.

Named layouts store normalized object positions and sizes in the presentation. Save a selection once, then select the same number of objects elsewhere and apply the layout.

For a shared library, create a normal Google Slides presentation containing one reusable component per slide. Connect its exact URL under **Deck → My Elements**. Library inserts copy every element except Slide Aid's metadata marker and retain a source-slide link. Saving the same item name updates that source slide; select an inserted component and use **Refresh selected** to pull the latest version while keeping its current position and size. **Add current selection** requires edit access to the library presentation.

Under **Icons**, search by concept, synonym, category, or consulting terminology — typing a concept such as "AI" or "KPI" surfaces the icons that express it, even when none is named that — then narrow by set or category, choose a color, and click an icon to insert it at the center of the slide. Results render in pages of 120. An inserted icon is a picture; select it (or leave nothing selected to take every icon on the slide) and click **Make Editable** to turn it into shapes you can recolor, ungroup and reshape.

Deck QA is non-destructive until the user clicks a specific **Fix** button. Accessibility and contrast findings that require editorial judgment are report-only.

### Upgrading from v0.1

No manual migration is required. Existing charts still decode their embedded `SLIDE_AID_CHART_V1` metadata. Selecting, rebuilding, recoloring, or restyling a chart moves its payload into shared Document Properties and replaces raw JSON alt text with a human-readable description.

## Development

```bash
npm run check       # TypeScript type-check
npm test            # compile and run pure geometry/data/metadata tests
npm run build       # bundle Apps Script and embed canonical icons
npm run push        # build and clasp push
npm run open        # open the Apps Script project
```

Architecture:

```text
src/core/          pure geometry, chart-data, chart-style and icon-path logic
src/slides/        selection, pinned reference, atomic shape batch, icon insertion
src/storage/       personal preferences and shared chunked document state
src/commands/      Slide Aid geometry, object and deck commands
src/charts/        atomic Chart Aid builders, edit/rebuild/restyle, annotations, samples
src/integrations/  read-only Google Sheets data adapter
src/layouts/       shared named layouts
src/library/       explicit-URL shared element libraries
src/qa/            deck analysis and targeted safe fixes
src/ui/            HTML sidebar
src/entrypoints/   Apps Script public functions
scripts/           local build/test helpers
tests/             Node tests for platform-independent logic
```

Shape-native charts and compatible geometry operations use atomic Slides API batches. Rotated or unsupported geometry falls back to `SlidesApp`. Area, pie, and doughnut still use the server-side Charts renderer and therefore remain image-based.
