# Shared Specs

This folder holds platform-neutral Slide Aid constants that should stay aligned
between the PowerPoint VBA add-in and the Google Slides Apps Script companion.

- `palettes.json`: canonical Chart Aid palette names and six-color values.
- `chart-kinds.json`: canonical chart kind identifiers and chart metadata markers.
- `chart-style.json`: the sixteen Chart Aid style parameters, their defaults, and
  the three palette families each chart kind belongs to.

The current implementations still define these constants locally for runtime
simplicity. Treat these files as the reviewable shared contract and migrate
platform code to generate from them when the build pipeline is ready.

Both sides are covered by tests that would fail on drift: the PowerPoint keys
live in `KeyDefs()` in `apps/powerpoint/src/modChartStyle.bas`, and the Google
Slides copy in `apps/google-slides/src/core/chart-style.ts` is asserted against
its own panel definitions by `apps/google-slides/tests/core.test.ts`.
