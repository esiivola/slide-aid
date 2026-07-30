#!/bin/bash
# =====================================================================
# Slide Aid - one-command build & reload.
#
#   cd apps/powerpoint
#   ./tools/build.sh            build dist/Slide Aid.ppam, offer to restart PP
#   ./tools/build.sh -r         build and restart PowerPoint silently
#
# How it works: tells the RUNNING PowerPoint (whose loaded Slide Aid
# add-in contains BuildSlideAid) to import src/*.bas into a fresh
# presentation, save it as .pptm and run the ribbon injector. Restart
# PowerPoint afterwards so it loads the new .ppam.
#
# One-time prerequisites (see README): AccessVBOM enabled, and the
# SlideAidUI.scpt helper compiled AFTER buildPpam was added to it.
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

# PowerPoint must run so the loaded add-in can host the build macro.
if ! pgrep -xq "Microsoft PowerPoint"; then
  echo "Starting PowerPoint (the loaded add-in hosts the build macro)..."
  open -a "Microsoft PowerPoint"
  sleep 6
fi

echo "Running BuildSlideAid inside PowerPoint..."
if ! osascript -e 'tell application "Microsoft PowerPoint" to run VB macro macro name "BuildSlideAid"'; then
  cat >&2 <<'EOF'

Could not trigger the macro via AppleScript. Run it manually instead:
  PowerPoint > Tools > Macro > Macros... > type: BuildSlideAid > Run
(then rerun this script with -r to restart PowerPoint, or restart it
yourself). If PowerPoint said the macro does not exist, the loaded
add-in predates BuildSlideAid - do one manual rebuild (README steps
1-6) to bootstrap it; every build after that is one command.
EOF
  exit 1
fi

# BuildSlideAid shows a completion dialog; once it returns,
# dist/Slide Aid.ppam is the new build. PowerPoint still has the OLD
# add-in loaded.
if [ "$RESTART" = "-r" ]; then
  ans=y
else
  read -r -p "Restart PowerPoint now to load the new build? [y/N] " ans
fi
case "${ans:-n}" in
  [Yy]*)
    osascript -e 'tell application "Microsoft PowerPoint" to quit'
    sleep 3
    open -a "Microsoft PowerPoint"
    echo "PowerPoint restarted - new build loaded."
    ;;
  *) echo "Remember: the new build loads on the next PowerPoint start." ;;
esac
