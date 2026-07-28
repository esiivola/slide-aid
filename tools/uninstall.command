#!/bin/bash
# =====================================================================
# Slide Aid - uninstaller. Removes everything install.command created.
# =====================================================================
set -uo pipefail

ok()   { printf '  \033[32m✓\033[0m %s\n' "$*"; }
warn() { printf '  \033[33m!\033[0m %s\n' "$*"; }

printf '\033[1mSlide Aid uninstaller\033[0m\n\n'

# color picker helper
rm -f "$HOME/Library/Application Scripts/com.microsoft.Powerpoint/SlideAidUI.scpt" \
  && ok "Color picker helper removed"

# add-in (all known Add-Ins locations)
OFFICE_UC="$HOME/Library/Group Containers/UBF8T346G9.Office"
for c in "User Content.localized/Add-Ins.localized" \
         "User Content.localized/Add-Ins" \
         "User Content/Add-Ins"; do
  if [ -f "$OFFICE_UC/$c/Slide Aid.ppam" ]; then
    rm -f "$OFFICE_UC/$c/Slide Aid.ppam"
    ok "Add-in removed from $c"
  fi
done

# hammerspoon config
if [ -f "$HOME/.hammerspoon/slideaid.lua" ]; then
  rm -f "$HOME/.hammerspoon/slideaid.lua"
  [ -f "$HOME/.hammerspoon/init.lua" ] && \
    sed -i '' '/require("slideaid")/d' "$HOME/.hammerspoon/init.lua"
  ok "Hammerspoon shortcut config removed (Hammerspoon itself was kept)"
fi

# per-user data (palettes, settings, element library) - ask first
STORE="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/SlideAid"
if [ -d "$STORE" ]; then
  read -r -p "Also delete your Slide Aid data (palettes, settings, My Elements)? [y/N] " a
  case "${a:-n}" in
    [Yy]*) rm -rf "$STORE"; ok "User data deleted" ;;
    *)     echo "  Kept: $STORE" ;;
  esac
fi

echo
warn "Finally: open PowerPoint > Tools > PowerPoint Add-ins... and remove"
warn "the 'Slide Aid' entry from the list (select it, click -)."
