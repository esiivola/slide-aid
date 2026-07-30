-- =====================================================================
-- Slide Aid - keyboard shortcuts via Hammerspoon
--
-- Binds system hotkeys that press Slide Aid ribbon buttons in
-- PowerPoint through the macOS accessibility API (the same channel
-- VoiceOver uses). Implemented outside PowerPoint because Mac PowerPoint
-- offers no in-process keyboard hooks.
--
-- Install:
--   brew install --cask hammerspoon
--   cp apps/powerpoint/hammerspoon/slideaid.lua ~/.hammerspoon/
--   echo 'require("slideaid")' >> ~/.hammerspoon/init.lua
--   Open Hammerspoon, grant Accessibility permission, "Reload Config".
--
-- Edit the BINDINGS table to taste, then Hammerspoon menu -> Reload
-- Config. `button` = the ribbon button's label. Optional `item` =
-- a menu entry to click after the button opens a menu.
-- =====================================================================

local BINDINGS = {
  -- Align to Master
  { mods = {"ctrl", "alt"}, key = "l", button = "Left" },
  { mods = {"ctrl", "alt"}, key = "r", button = "Right" },
  { mods = {"ctrl", "alt"}, key = "t", button = "Top" },
  { mods = {"ctrl", "alt"}, key = "b", button = "Bottom" },
  { mods = {"ctrl", "alt"}, key = "c", button = "Center" },
  { mods = {"ctrl", "alt"}, key = "m", button = "Middle" },
  -- Distribute
  { mods = {"ctrl", "alt"}, key = "h", button = "Distribute H" },
  { mods = {"ctrl", "alt"}, key = "v", button = "Distribute V" },
  -- Size to Master
  { mods = {"ctrl", "alt"}, key = "1", button = "Width" },
  { mods = {"ctrl", "alt"}, key = "2", button = "Height" },
  { mods = {"ctrl", "alt"}, key = "3", button = "Width + Height" },
  -- One-click grid / golden canon / painter
  { mods = {"ctrl", "alt"}, key = "x", button = "Matrix" },
  { mods = {"ctrl", "alt"}, key = "g", button = "Golden Canon" },
  { mods = {"ctrl", "alt"}, key = "p", button = "Format Painter" },
  -- Menu tools: press the menu button, then click the item
  { mods = {"ctrl", "alt"}, key = "s", button = "Swap",  item = "At Centers" },
  { mods = {"ctrl", "alt"}, key = "k", button = "Stack", item = "Horizontally" },
  { mods = {"ctrl", "alt", "shift"}, key = "k", button = "Stack", item = "Vertically" },
}

local RIBBON_TAB = "Slide Aid"
local BUNDLE = "com.microsoft.Powerpoint"

local ax = require("hs.axuielement")

-- Breadth-first search of the accessibility tree.
local function findElement(root, matchFn, maxDepth)
  if not root then return nil end
  local queue = { { root, 0 } }
  while #queue > 0 do
    local el, depth = queue[1][1], queue[1][2]
    table.remove(queue, 1)
    local okMatch, matched = pcall(matchFn, el)
    if okMatch and matched then return el end
    if depth < maxDepth then
      local kids = el:attributeValue("AXChildren")
      if kids then
        for _, k in ipairs(kids) do
          queue[#queue + 1] = { k, depth + 1 }
        end
      end
    end
  end
  return nil
end

local function titled(el, title)
  return el:attributeValue("AXTitle") == title
      or el:attributeValue("AXDescription") == title
end

local function isPressable(el)
  local role = el:attributeValue("AXRole")
  return role == "AXButton" or role == "AXMenuButton"
      or role == "AXRadioButton" or role == "AXCheckBox"
      or role == "AXTab"
end

local function press(el)
  return pcall(function() el:performAction("AXPress") end)
end

local function pptWindow()
  local app = hs.application.get(BUNDLE)
  if not app then
    hs.alert.show("PowerPoint is not running")
    return nil
  end
  app:activate()
  local axApp = ax.applicationElement(app)
  return axApp:attributeValue("AXMainWindow")
      or axApp:attributeValue("AXFocusedWindow"), axApp
end

local function findButton(win, title)
  return findElement(win, function(el)
    return isPressable(el) and titled(el, title)
  end, 14)
end

-- Click `item` in the menu that just opened (menus live at app level).
local function clickMenuItem(axApp, item, attempt)
  attempt = attempt or 1
  local el = findElement(axApp, function(e)
    local role = e:attributeValue("AXRole")
    return (role == "AXMenuItem" or role == "AXButton") and titled(e, item)
  end, 16)
  if el then
    press(el)
  elseif attempt < 4 then
    hs.timer.doAfter(0.15, function() clickMenuItem(axApp, item, attempt + 1) end)
  else
    hs.alert.show("Slide Aid: menu item '" .. item .. "' not found")
  end
end

local function trigger(binding)
  local win, axApp = pptWindow()
  if not win then return end

  local function pressAndFollow()
    local btn = findButton(win, binding.button)
    if not btn then
      hs.alert.show("Slide Aid: button '" .. binding.button .. "' not found")
      return
    end
    press(btn)
    if binding.item then
      hs.timer.doAfter(0.2, function() clickMenuItem(axApp, binding.item) end)
    end
  end

  -- Generic labels ("Left", "Width", "Height", ...) also occur on
  -- other ribbon tabs, so make sure the Slide Aid tab is the ACTIVE
  -- one before searching for the button - otherwise a same-named
  -- button of the current tab could be pressed instead.
  local tab = findElement(win, function(el)
    return isPressable(el) and titled(el, RIBBON_TAB)
  end, 14)
  if not tab then
    hs.alert.show("Slide Aid tab not found - is the add-in loaded?")
    return
  end
  local sel = tab:attributeValue("AXValue")
  local tabActive
  if sel == nil then
    -- AXValue not exposed: fall back to "button visible = tab active"
    tabActive = findButton(win, binding.button) ~= nil
  else
    tabActive = (sel == 1 or sel == true)
  end
  if tabActive then
    pressAndFollow()
  else
    press(tab)
    hs.timer.doAfter(0.3, pressAndFollow)
  end
end

-- Hotkeys are registered but only ENABLED while PowerPoint is the
-- frontmost app, so the ctrl+alt combos keep working normally in
-- every other application.
local hotkeys = {}
for _, b in ipairs(BINDINGS) do
  hotkeys[#hotkeys + 1] = hs.hotkey.new(b.mods, b.key, function() trigger(b) end)
end

local function setHotkeys(on)
  for _, hk in ipairs(hotkeys) do
    if on then hk:enable() else hk:disable() end
  end
end

-- Global (not local): keeps the watcher alive across garbage collection.
slideAidAppWatcher = hs.application.watcher.new(function(_, event, app)
  if app and app:bundleID() == BUNDLE then
    if event == hs.application.watcher.activated then
      setHotkeys(true)
    elseif event == hs.application.watcher.deactivated
        or event == hs.application.watcher.terminated then
      setHotkeys(false)
    end
  end
end)
slideAidAppWatcher:start()

local front = hs.application.frontmostApplication()
setHotkeys(front ~= nil and front:bundleID() == BUNDLE)

-- Auto-reload: saving any .lua file in ~/.hammerspoon applies it
-- immediately - no manual "Reload Config" needed.
if not slideAidWatcher then
  slideAidWatcher = hs.pathwatcher.new(os.getenv("HOME") .. "/.hammerspoon/",
    function(files)
      for _, f in pairs(files) do
        if f:sub(-4) == ".lua" then
          hs.reload()
          return
        end
      end
    end):start()
end

hs.alert.show("Slide Aid shortcuts loaded (" .. #BINDINGS .. ")")
