# Shared Specs

This folder holds platform-neutral Slide Aid constants that should stay aligned
between the PowerPoint VBA add-in and the Google Slides Apps Script companion.

- `palettes.json`: canonical Chart Aid palette names and six-color values.
- `chart-kinds.json`: canonical chart kind identifiers and chart metadata prefixes.

The current implementations still define these constants locally for runtime
simplicity. Treat these files as the reviewable shared contract and migrate
platform code to generate from them when the build pipeline is ready.
