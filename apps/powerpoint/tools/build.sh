#!/bin/bash
# =====================================================================
# Slide Aid - one-command build & reload.
#
#   cd apps/powerpoint
#   ./tools/build.sh            build dist/Slide Aid.ppam, offer to restart PP
#   ./tools/build.sh -r         build and restart PowerPoint silently
#
# How it works: refreshes the sandbox source cache + icons.dat, starts
# PowerPoint if needed (it auto-loads the Slide Aid add-in, which hosts
# BuildSlideAid), tells it to import src/*.bas into a fresh presentation,
# save a .pptm and run the ribbon injector -> dist/Slide Aid.ppam. It then
# syncs that build to whatever .ppam PowerPoint actually has loaded (the dev
# load path can differ from the build output) and restarts PowerPoint.
#
# Works from a fully closed PowerPoint. Click OK if a 'rebuilt' dialog appears.
#
# One-time prerequisites (see README): AccessVBOM enabled, and the
# SlideAidUI.scpt helper compiled/installed - it hosts buildPpam for this
# build plus Chart Aid's Chart Settings / Edit Colors dialogs. Recompile it
# whenever tools/SlideAidUI.applescript changes.
# =====================================================================
set -euo pipefail
cd "$(dirname "$0")/.."

RESTART="${1:-}"

# The build macro runs from the loaded add-in and cannot infer this
# checkout's location. Persist the current app root in PowerPoint's
# sandbox without embedding a developer-specific path in the add-in.
STORE="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/SlideAid"
SOURCE_CACHE="$STORE/build/src"
mkdir -p "$SOURCE_CACHE"
pwd > "$STORE/repo_path.txt"
find "$SOURCE_CACHE" -type f -name '*.bas' -delete
cp src/*.bas "$SOURCE_CACHE/"
cp tools/import_helper.bas "$SOURCE_CACHE/modImportHelper.bas"

# IconAid loads its 2,500 icons at runtime from icons.dat in the SlideAid
# folder (too much data to live inside the VBA project). Environ("HOME") in
# sandboxed PowerPoint maps here, so the add-in finds it at SlideAid/icons.dat.
if [ ! -f data/icons.dat ]; then
  echo "ERROR: data/icons.dat is missing - run scripts/build_iconaid_web.py (from the repo root) first." >&2
  exit 1
fi
cp data/icons.dat "$STORE/icons.dat"

ADDIN="dist/Slide Aid.ppam"

# PowerPoint must run so its auto-loaded add-in can host BuildSlideAid.
# From a cold start, wait until it is scriptable and the add-in has loaded.
if ! pgrep -xq "Microsoft PowerPoint"; then
  echo "Starting PowerPoint (it auto-loads the Slide Aid add-in, which hosts BuildSlideAid)..."
  open -a "Microsoft PowerPoint"
  for _ in $(seq 1 20); do
    if osascript -e 'tell application "Microsoft PowerPoint" to version' >/dev/null 2>&1; then break; fi
    sleep 2
  done
  sleep 4   # let the add-in finish loading
fi

before=$(stat -f %m "$ADDIN" 2>/dev/null || echo 0)

echo "Running BuildSlideAid inside PowerPoint (click OK if a 'rebuilt' dialog appears)..."
# The macro ends with a dialog, so this AppleEvent may 'time out' while it waits
# for the click - harmless; success is verified by the rebuilt file below.
osascript -e 'tell application "Microsoft PowerPoint" to run VB macro macro name "BuildSlideAid"' >/dev/null 2>&1 || true

# Wait (up to ~2 min) for the build output to update.
for _ in $(seq 1 60); do
  [ "$(stat -f %m "$ADDIN" 2>/dev/null || echo 0)" != "$before" ] && break
  sleep 2
done
if [ "$(stat -f %m "$ADDIN" 2>/dev/null || echo 0)" = "$before" ]; then
  cat >&2 <<'EOF'

'dist/Slide Aid.ppam' did not update. Run the macro by hand:
  PowerPoint > Tools > Macro > Macros... > (Macros in: Slide Aid.ppam) > BuildSlideAid > Run
If it is not listed / errors, open the VB Editor (Option+F11) >
Debug > Compile VBAProject to see the compile error.
EOF
  exit 1
fi

# Dev machines often load the add-in from a different copy than the build
# output (e.g. the installer-staging 'dist/Slide Aid/'). Sync the fresh build
# to whatever .ppam PowerPoint actually has open so the restart loads it.
pid=$(pgrep -x "Microsoft PowerPoint" | head -1 || true)
loaded=$(lsof -p "${pid:-0}" -F n 2>/dev/null | sed -n 's/^n//p' | grep -i '\.ppam$' | grep -i 'slide aid' | head -1 || true)
if [ -n "$loaded" ] && [ "$loaded" != "$PWD/$ADDIN" ]; then
  cp "$ADDIN" "$loaded"
  echo "Synced new build to the loaded add-in: $loaded"
fi

# Restart so PowerPoint loads the new build.
if [ "$RESTART" = "-r" ]; then
  ans=y
else
  read -r -p "Restart PowerPoint now to load the new build? [y/N] " ans
fi
case "${ans:-n}" in
  [Yy]*)
    osascript -e 'tell application "Microsoft PowerPoint" to quit' >/dev/null 2>&1 || true
    sleep 4
    open -a "Microsoft PowerPoint"
    echo "PowerPoint restarted - new build loaded."
    ;;
  *) echo "Remember: the new build loads on the next PowerPoint start." ;;
esac
