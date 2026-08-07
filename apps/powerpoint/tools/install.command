#!/bin/bash
# =====================================================================
# Slide Aid - installer for colleagues (per-user, no admin rights).
# Double-click to run. If macOS blocks it (downloaded file), right-
# click -> Open -> Open once.
#
# Installs:
#   1. SlideAidUI.scpt   -> native macOS dialogs for Chart Aid (settings,
#                           colors, color picker)
#   2. Slide Aid.ppam    -> the add-in, into Office's Add-Ins folder
#   3. slideaid.lua      -> keyboard shortcuts via Hammerspoon (asked)
# =====================================================================
set -euo pipefail
cd "$(dirname "$0")"

bold() { printf '\033[1m%s\033[0m\n' "$*"; }
ok()   { printf '  \033[32m✓\033[0m %s\n' "$*"; }
warn() { printf '  \033[33m!\033[0m %s\n' "$*"; }

bold "Slide Aid installer"
echo

# --- sanity -----------------------------------------------------------
if [ ! -d "/Applications/Microsoft PowerPoint.app" ]; then
  warn "Microsoft PowerPoint was not found in /Applications."
  warn "Install PowerPoint first, then run this installer again."
  exit 1
fi
for f in "SlideAidUI.scpt" "Slide Aid.ppam" "slideaid.lua" "icons.dat" "manifest.xml"; do
  if [ ! -f "$f" ]; then
    warn "Missing '$f' next to this installer - unzip the whole folder first."
    exit 1
  fi
done

# --- 1. native dialogs helper -----------------------------------------
SCRIPTS_DIR="$HOME/Library/Application Scripts/com.microsoft.Powerpoint"
mkdir -p "$SCRIPTS_DIR"
cp -f "SlideAidUI.scpt" "$SCRIPTS_DIR/"
ok "Native dialogs helper installed (Chart Aid settings, colors, and the color picker)"

# --- 2. the add-in -----------------------------------------------------
OFFICE_UC="$HOME/Library/Group Containers/UBF8T346G9.Office"
ADDINS=""
for c in "User Content.localized/Add-Ins.localized" \
         "User Content.localized/Add-Ins" \
         "User Content/Add-Ins"; do
  if [ -d "$OFFICE_UC/$c" ]; then ADDINS="$OFFICE_UC/$c"; break; fi
done
if [ -z "$ADDINS" ]; then
  ADDINS="$OFFICE_UC/User Content/Add-Ins"
  mkdir -p "$ADDINS"
fi
cp -f "Slide Aid.ppam" "$ADDINS/"
# strip the download-quarantine flag so Office doesn't hard-block macros
xattr -d com.apple.quarantine "$ADDINS/Slide Aid.ppam" 2>/dev/null || true
ok "Add-in copied to Office's Add-Ins folder"

# --- 2b. icon library (used by Chart Aid > Make Editable) --------------
# PowerPoint's sandbox maps VBA's Environ("HOME") to this container dir, so
# the add-in reads the library from SlideAid/icons.dat when converting the
# sidebar's inserted pictures into editable freeforms.
ICON_STORE="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/SlideAid"
mkdir -p "$ICON_STORE"
cp -f "icons.dat" "$ICON_STORE/icons.dat"
ok "Icon library installed (for Chart Aid > Make Editable)"

# --- 2c. Icons sidebar (Office task pane) - sideload the manifest ----
# Dropping the manifest in PowerPoint's 'wef' folder registers the add-in.
# Its content loads from GitHub Pages, so the sidebar needs internet. It
# adds an "Insert Icons" button on PowerPoint's Insert tab.
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF"
cp -f "manifest.xml" "$WEF/iconaid-manifest.xml"
ok "Icons sidebar registered (Insert tab > Insert Icons; needs internet)"

# --- 3. keyboard shortcuts (optional) ----------------------------------
echo
read -r -p "Install keyboard shortcuts via Hammerspoon (free, open-source)? [y/N] " ans
case "${ans:-n}" in
  [Yy]*)
    if [ ! -d "/Applications/Hammerspoon.app" ]; then
      if command -v brew >/dev/null 2>&1; then
        read -r -p "  Hammerspoon is not installed. Install it with Homebrew now? [y/N] " b
        case "${b:-n}" in
          [Yy]*) brew install --cask hammerspoon || \
                   warn "Homebrew install failed - get Hammerspoon from https://www.hammerspoon.org" ;;
          *) warn "Skipping Hammerspoon app install - get it from https://www.hammerspoon.org" ;;
        esac
      else
        warn "Hammerspoon is not installed and Homebrew is not available."
        warn "Download it from https://www.hammerspoon.org (drag to Applications),"
        warn "then run this installer again. Installing the config anyway."
        open "https://www.hammerspoon.org" || true
      fi
    fi
    mkdir -p "$HOME/.hammerspoon"
    cp -f "slideaid.lua" "$HOME/.hammerspoon/"
    touch "$HOME/.hammerspoon/init.lua"
    grep -qs 'require("slideaid")' "$HOME/.hammerspoon/init.lua" || \
      printf '\nrequire("slideaid")\n' >> "$HOME/.hammerspoon/init.lua"
    ok "Shortcut config installed (~/.hammerspoon/slideaid.lua)"
    if [ -d "/Applications/Hammerspoon.app" ]; then
      open -a Hammerspoon || true
      warn "If macOS asks: grant Hammerspoon 'Accessibility' permission"
      warn "(System Settings > Privacy & Security > Accessibility), then"
      warn "use its menu-bar icon -> Reload Config."
    fi
    ;;
  *) echo "  Skipped - you can rerun this installer later to add shortcuts." ;;
esac

# --- final manual step --------------------------------------------------
echo
bold "One manual step remains (macOS does not allow automating it):"
echo "  1. Open PowerPoint"
echo "  2. Tools > PowerPoint Add-ins..."
echo "  3. 'Slide Aid' should be listed - tick it. If it is not listed,"
echo "     click + and choose:"
echo "       $ADDINS/Slide Aid.ppam"
echo "  4. Click 'Enable Macros' when asked, then restart PowerPoint."
echo
bold "Done. You should then see the 'Slide Aid' and 'Chart Aid' ribbon tabs."
