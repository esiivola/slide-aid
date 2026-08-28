Attribute VB_Name = "modShortcuts"
' =====================================================================
' Slide Aid - Keyboard shortcuts on Mac
'
' PowerPoint VBA has no OnKey, and Mac PowerPoint cannot assign
' shortcuts to macros directly. The reliable Mac route:
'
'   1. This module adds a "Slide Aid" menu to the Mac MENU BAR
'      (via CommandBars) with one named item per tool.
'   2. macOS System Settings > Keyboard > Keyboard Shortcuts >
'      App Shortcuts > "+" > App: Microsoft PowerPoint,
'      Menu title: the EXACT item name (e.g. "Align Left to Master"),
'      then press the key combo YOU want.
'
' Every menu item below can get its own user-chosen shortcut this
' way, and you can change them at any time in System Settings.
' The menu is (re)built automatically when the add-in loads.
' =====================================================================
Option Explicit

Private Const MENU_NAME As String = "Slide Aid"

' Runs automatically when the add-in loads.
Public Sub Auto_Open()
    On Error Resume Next
    BuildShortcutMenu False
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------
' Ribbon "Shortcuts" button:
'  * Hammerspoon installed -> open the shortcut config for editing
'    (it auto-reloads on save, so edits apply immediately)
'  * otherwise -> explain how to get shortcuts, and try the legacy
'    menu-bar route for PowerPoint builds that still allow it
' ---------------------------------------------------------------
Public Sub ShortcutsButton()
    If Dir("/Applications/Hammerspoon.app", vbDirectory) <> "" Then
        OpenShortcutConfig
    Else
        MsgBox "Keyboard shortcuts use Hammerspoon (free)." & vbCr & vbCr & _
               "Install:  brew install --cask hammerspoon" & vbCr & _
               "Then copy apps/powerpoint/hammerspoon/slideaid.lua from the Slide Aid repo " & _
               "to ~/.hammerspoon/ and add  require(""slideaid"")  to " & _
               "~/.hammerspoon/init.lua (see README).", _
               vbInformation, MENU_NAME
        BuildShortcutMenu False   ' bonus on builds that allow CommandBars
    End If
End Sub

Private Sub OpenShortcutConfig()
    Dim cfg As String
    cfg = RealHome() & "/.hammerspoon/slideaid.lua"

    If Dir(cfg) = "" Then
        MsgBox "Hammerspoon is installed, but the Slide Aid config is missing." & vbCr & vbCr & _
               "Copy apps/powerpoint/hammerspoon/slideaid.lua from the Slide Aid repo to " & _
               "~/.hammerspoon/ and add  require(""slideaid"")  to init.lua.", _
               vbExclamation, MENU_NAME
        Exit Sub
    End If

    ' Open the config; edits apply on save (the config watches itself).
    On Error GoTo TryFolder
    ActivePresentation.FollowHyperlink "file://" & cfg
    Exit Sub
TryFolder:
    On Error GoTo GiveUp
    ActivePresentation.FollowHyperlink "file://" & RealHome() & "/.hammerspoon/"
    Exit Sub
GiveUp:
    MsgBox "Edit this file (each shortcut is one line; saves apply instantly):" & _
           vbCr & cfg, vbInformation, MENU_NAME
End Sub

Public Sub RemoveShortcutMenu()
    On Error Resume Next
    Application.CommandBars("Menu Bar").Controls(MENU_NAME).Delete
    On Error GoTo 0
End Sub

Public Sub BuildShortcutMenu(Optional ByVal verbose As Boolean = True)
    On Error GoTo Unsupported
    RemoveShortcutMenu

    Dim menu As CommandBarPopup
    Set menu = Application.CommandBars("Menu Bar").Controls.Add(msoControlPopup, , , , True)
    menu.Caption = MENU_NAME

    ' --- Position ---
    AddItem menu, "Align Left to Master", "SC_AlignL"
    AddItem menu, "Align Right to Master", "SC_AlignR"
    AddItem menu, "Align Top to Master", "SC_AlignT"
    AddItem menu, "Align Bottom to Master", "SC_AlignB"
    AddItem menu, "Center on Master", "SC_AlignCH"
    AddItem menu, "Middle on Master", "SC_AlignCV"
    AddItem menu, "Dock Left", "SC_DockL"
    AddItem menu, "Dock Right", "SC_DockR"
    AddItem menu, "Dock Up", "SC_DockT"
    AddItem menu, "Dock Down", "SC_DockB"
    AddItem menu, "Distribute Horizontally", "SC_DistH"
    AddItem menu, "Distribute Vertically", "SC_DistV"
    AddItem menu, "Swap Objects", "SC_Swap", True
    AddItem menu, "Stack Horizontally", "SC_StackH"
    AddItem menu, "Stack Vertically", "SC_StackV"
    AddItem menu, "Golden Canon", "SC_Golden"

    ' --- Size ---
    AddItem menu, "Same Width as Master", "SC_SizeW", True
    AddItem menu, "Same Height as Master", "SC_SizeH"
    AddItem menu, "Same Size as Master", "SC_SizeWH"
    AddItem menu, "Same Width as Master (keep ratio)", "SC_SizeWR"
    AddItem menu, "Same Height as Master (keep ratio)", "SC_SizeHR"
    AddItem menu, "Magic Resizer", "SC_Magic"
    AddItem menu, "Stretch Right to Master", "SC_StretchR"
    AddItem menu, "Stretch Left to Master", "SC_StretchL"

    ' --- Tools ---
    AddItem menu, "Advanced Format Painter", "SC_Painter", True
    AddItem menu, "Select Similar Shapes", "SC_Similar"
    AddItem menu, "Split Text Box at Cursor", "SC_Split"
    AddItem menu, "Merge Text Boxes", "SC_Merge"
    AddItem menu, "Hide Selected Objects", "SC_Hide"
    AddItem menu, "Unhide All Objects", "SC_Unhide"

    If verbose Then
        MsgBox "The '" & MENU_NAME & "' menu is now in the menu bar." & vbCr & vbCr & _
               "Assign your own shortcut to any item:" & vbCr & _
               "System Settings > Keyboard > Keyboard Shortcuts > App Shortcuts > +" & vbCr & _
               "App: Microsoft PowerPoint - Menu title: exact item name - your keys.", _
               vbInformation, "Slide Aid"
    End If
    Exit Sub
Unsupported:
    If verbose Then
        MsgBox "This PowerPoint version does not allow menu-bar customization " & _
               "(CommandBars). Alternative: a macro runner utility such as " & _
               "Keyboard Maestro can trigger ribbon buttons by name.", _
               vbExclamation, "Slide Aid"
    End If
End Sub

Private Sub AddItem(menu As CommandBarPopup, ByVal caption As String, _
                    ByVal action As String, Optional ByVal group As Boolean = False)
    Dim b As CommandBarButton
    Set b = menu.Controls.Add(msoControlButton)
    b.Caption = caption
    b.OnAction = action
    b.BeginGroup = group
End Sub

' ---- thin wrappers (OnAction cannot pass arguments) ----
Public Sub SC_AlignL():   AlignToMaster "L":        End Sub
Public Sub SC_AlignR():   AlignToMaster "R":        End Sub
Public Sub SC_AlignT():   AlignToMaster "T":        End Sub
Public Sub SC_AlignB():   AlignToMaster "B":        End Sub
Public Sub SC_AlignCH():  AlignToMaster "CH":       End Sub
Public Sub SC_AlignCV():  AlignToMaster "CV":       End Sub
Public Sub SC_DockL():    DockToMaster "L":         End Sub
Public Sub SC_DockR():    DockToMaster "R":         End Sub
Public Sub SC_DockT():    DockToMaster "T":         End Sub
Public Sub SC_DockB():    DockToMaster "B":         End Sub
Public Sub SC_DistH():    DistributeObjects "H":    End Sub
Public Sub SC_DistV():    DistributeObjects "V":    End Sub
Public Sub SC_Swap():     SwapObjects "PZ", "C":    End Sub
Public Sub SC_StackH():   StackObjects "H", False:  End Sub
Public Sub SC_StackV():   StackObjects "V", False:  End Sub
Public Sub SC_Golden():   GoldenCanon:              End Sub
Public Sub SC_SizeW():    MatchSizeToMaster "W":    End Sub
Public Sub SC_SizeH():    MatchSizeToMaster "H":    End Sub
Public Sub SC_SizeWH():   MatchSizeToMaster "WH":   End Sub
Public Sub SC_SizeWR():   MatchSizeToMaster "WR":   End Sub
Public Sub SC_SizeHR():   MatchSizeToMaster "HR":   End Sub
Public Sub SC_Magic():    MagicResizer:             End Sub
Public Sub SC_StretchR(): StretchToMaster "R":      End Sub
Public Sub SC_StretchL(): StretchToMaster "L":      End Sub
Public Sub SC_Painter():  AdvancedFormatPainter:    End Sub
Public Sub SC_Similar():  SelectSimilarShapes "TF": End Sub
Public Sub SC_Split():    SplitTextBox:             End Sub
Public Sub SC_Merge():    MergeTextBoxes:           End Sub
Public Sub SC_Hide():     HideSelectedObjects:      End Sub
Public Sub SC_Unhide():   UnhideHiddenObjects:      End Sub
