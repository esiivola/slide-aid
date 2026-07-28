# Slide Aid for Google Slides

This folder contains the Google Slides companion to Slide Aid. The PowerPoint add-in remains the main and canonical version. Names, order, icons, defaults, and chart semantics come from the [PowerPoint UI reference](../docs/POWERPOINT_UI_REFERENCE.md), the [chart layout reference](../docs/CHARTS.md), the [RibbonX definition](../ribbon/customUI14.xml), and the original [PowerPoint README](../README.md).

This is a TypeScript Google Workspace editor add-on built with Apps Script and `clasp`. It opens an HTML sidebar, operates on the current Slides selection, and deliberately separates personal user settings from shared presentation state.

## Implemented in v0.2

- An explicit **Set reference** workflow to replace PowerPoint's last-selected Master convention.
- Align to pinned reference, slide, or selection bounds; dock; stretch; fill gap; match dimensions; Golden Canon.
- Stack, exact spacing, distribute, matrix, scale, place-on-slide regions, slice, and multiply.
- RGB and theme-linked fill, border, and text colors; theme/RGB conversion; live deck-theme swatches; optional palette-to-theme-accent updates.
- Rotation matching, case conversion, double-space cleanup, and text swap.
- All Chart Aid chart families: column, bar, stacked column, stacked bar, 100%, Waterfall, Mekko, line, area, pie, doughnut, scatter, bubble, and Gantt.
- Atomic shape-native chart construction through the Advanced Slides service, with revision checks to avoid partial charts.
- Accessible chart alt text plus chunked shared metadata for Edit Data and Rebuild; old v0.1 alt-text JSON migrates automatically when a chart is selected or rebuilt.
- Nine canonical shared deck palettes, restyle selected/all, and persistent per-series recoloring.
- Read-only Google Sheets connections: build from an explicit spreadsheet URL, tab, and range, then refresh the selected chart from current values.
- A live selection inspector, command search, matrix preview, and named layouts shared with deck collaborators.
- Shared element libraries backed by one explicitly configured Slides presentation. Slide Aid never searches the user's Drive.
- Deck QA reporting for off-slide objects, tiny text, missing alt text, low contrast, fixed RGB fills, stale/broken Sheet sources, orphan datasheets, and irregular horizontal spacing. Safe mechanical issues have targeted fixes.
- The PowerPoint PNG icons are embedded into the built sidebar at build time, so no public image host is required.

## Deliberate differences and current gaps

Google Slides does not expose ordinary click/shift-click selection order. Reference-based tools therefore use a pinned reference, and order-dependent tools use deterministic spatial order: horizontal operations sort left-to-right, vertical operations sort top-to-bottom, and Matrix sorts left-to-right. The selection API behavior is documented in Google's [Slides selection guide](https://developers.google.com/apps-script/guides/slides/selecting).

Google Workspace add-ons use menus, dialogs, and HTML sidebars rather than Office RibbonX or shape context menus. Slide Aid therefore appears under **Extensions → Slide Aid → Open Slide Aid** and runs from the sidebar. Google documents this model in [Dialogs and sidebars](https://developers.google.com/apps-script/guides/dialogs).

Area, pie, and doughnut charts are server-rendered with the Apps Script Charts service and inserted as static images. They retain Slide Aid metadata and remain editable through **Edit data → Rebuild**, but individual segments are not native Slides shapes. Google documents server-side image rendering in the [Charts service reference](https://developers.google.com/apps-script/reference/charts/).

The shared Slides library replaces the local PowerPoint My Elements store, but it is intentionally URL-based rather than a Drive-wide browser. Google Slides still lacks reliable click-order tracking and add-on keyboard-shortcut hooks.

The following PowerPoint features are not yet implemented here: My Formats, Agenda, advanced Format Painter, Select Similar, Process Chain, block-arrow and rounded-rectangle adjustment handles, Snap to Table, split/merge text with formatting preservation, hide/cleanup/language tools, chart annotations, Harvey balls, checkboxes, and sample slides. See the [PowerPoint UI reference](../docs/POWERPOINT_UI_REFERENCE.md) for their canonical behavior.

## Permissions and stored data

The add-on requests:

- `presentations`: required for atomic Slides API updates and for copying elements to or from an explicitly configured library presentation.
- `spreadsheets.readonly`: required only when building or refreshing a linked-data chart.
- `script.container.ui`: required for the editor menu and sidebar.

It does **not** request a Drive scope and cannot list or search Drive files. Users must paste the exact Google Slides library URL and Google Sheets source URL.

Personal values such as the pinned reference and preferred spacing live in User Properties. Deck palettes, named layouts, library configuration, and chart payloads live in Document Properties shared by collaborators. Chart data is split into Unicode-safe chunks below the Apps Script per-property size limit.

## Local setup

Requirements: Node.js 20 or newer and a Google account allowed to create Apps Script projects.

```bash
cd google-slides
npm install
npm run check
npm test
npm run build
```

The build produces only the Apps Script deployment files under `dist/`:

```text
dist/Code.js
dist/Sidebar.html
dist/appsscript.json
```

Source files, tests, and `node_modules` are excluded from deployment.

## Create and push an Apps Script project

Authenticate once:

```bash
npx clasp login
```

Create a standalone Apps Script project, or create one in the Apps Script UI and copy its script ID. For a new project from the command line:

```bash
npx clasp create --type standalone --title "Slide Aid for Google Slides" --rootDir dist
```

If using an existing project, copy `.clasp.example.json` to `.clasp.json`, replace `YOUR_SCRIPT_ID`, and keep `rootDir` set to `dist`. Then push:

```bash
npm run push
```

In the Apps Script editor:

1. Open **Deploy → Test deployments**.
2. Choose **Editor add-on** and Google Slides as the host application.
3. Install the test deployment for your account.
4. Open or reload a Google Slides presentation.
5. Confirm that **Services** contains **Google Slides API** (`Slides`, v1). It is declared by `appsscript.json`; if the project uses a custom Google Cloud project, also enable the Google Slides API there.
6. Use **Extensions → Slide Aid → Open Slide Aid** and authorize presentation, read-only spreadsheet, and container-UI access.

For organization-wide distribution, create a versioned deployment and configure a private Google Workspace Marketplace SDK listing. Keep PowerPoint wording and screenshots canonical; document only Google-specific differences in this folder.

## Use

For Master-based tools, select one object and click **Set reference**. Then select the targets and run a command. The reference is keyed to both presentation and slide, and Slide Aid refuses to use it on a different slide.

For charts, create and select a Google Slides table in one of the formats defined in [Chart layouts](../docs/CHARTS.md), then choose a chart type. Select a generated chart or one of its elements to edit, rebuild, restyle, or recolor it.

For a linked chart, open **Charts → Linked Google Sheets data**, paste the exact spreadsheet URL, enter its tab and A1 range, and choose a chart type. **Refresh selected** reads the range again and rebuilds the selected chart. Rebuild also refreshes linked data automatically.

Named layouts store normalized object positions and sizes in the presentation. Save a selection once, then select the same number of objects elsewhere and apply the layout.

For a shared library, create a normal Google Slides presentation containing one reusable component per slide. Connect its exact URL under **Library**. Library inserts copy every element except Slide Aid's metadata marker and retain a source-slide link. Saving the same item name updates that source slide; select an inserted component and use **Refresh selected** to pull the latest version while keeping its current position and size. **Add current selection** requires edit access to the library presentation.

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
src/core/          pure geometry and chart-data logic
src/slides/        selection and pinned-reference adapter
src/storage/       personal preferences and shared chunked document state
src/commands/      Slide Aid object commands
src/charts/        atomic Chart Aid builders, edit/rebuild/restyle
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
