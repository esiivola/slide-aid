# IconAid Host-App Verification

Date: 2026-07-31

## PowerPoint

- Updated `apps/powerpoint-iconaid/manifest.xml` to version `1.1.0.0` and cache-busted both task-pane URLs to `v=1.1.0-schema3-pilot`.
- Refreshed the local sideload manifest at `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/IconAid.xml` so PowerPoint no longer points at the old `v=1.0.2` task pane URL.
- Verified the currently running localhost HTTPS server returns `200 OK` for:
  - `https://127.0.0.1:3000/apps/powerpoint-iconaid/taskpane.html?v=1.1.0-schema3-pilot`
  - `https://127.0.0.1:3000/shared/iconaid/catalog.json`
- Confirmed the current localhost server is broad and still serves non-IconAid repository files. The repo `dev_server.py` has therefore been narrowed to IconAid task-pane assets for future use, but the already-running process was not killed.
- Created a new PowerPoint test presentation without closing existing user presentations.
- Could not complete automated click-through insertion from the real task pane because PowerPoint did not expose a usable window through macOS accessibility automation in this session.

## Google Slides

- No `apps/google-slides/.clasp.json` is present, so there is no local Apps Script project binding for a live `clasp push` or test deployment from this checkout.
- Local verification remains the Apps Script build and unit test path until a script ID/test deployment is configured.

## Remaining Host Gates

- Open the IconAid task pane inside PowerPoint after restarting PowerPoint, search for pilot icons, insert multiple icons in light and dark colors, and confirm a single selectable vector group is inserted on the current slide without adding slides.
- Push/install the Google Slides test deployment, open the sidebar in a real Slides presentation, search both the consulting set and the library, insert in multiple colors, then run **Make Editable** on a stroke icon and on a solid Bootstrap icon and confirm both convert to usable shapes at the picture's own size.
