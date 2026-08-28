Attribute VB_Name = "modRibbon"
' =====================================================================
' Slide Aid - Ribbon callbacks
' All buttons in customUI14.xml call RB_Dispatch with a Tag of the
' form "Command:Arg[:Arg2]".
' =====================================================================
Option Explicit

' Captured at ribbon load; needed to refresh cached dynamicMenus
' (My Elements / My Formats) after their contents change.
Public gRibbon As IRibbonUI

Public Sub RB_OnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

' Call after adding elements/formats so the menus rebuild on next open.
Public Sub RefreshDynamicMenus()
    On Error Resume Next
    If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "saLib"
        gRibbon.InvalidateControl "saFmt"
    End If
    On Error GoTo 0
End Sub

Public Sub RB_Dispatch(control As IRibbonControl)
    RB_DispatchTag control.Tag
End Sub

' Separated from the IRibbonControl signature so a stub add-in can
' forward clicks here via Application.Run (Mac PowerPoint only fires
' ribbon callbacks from loaded add-ins, not from open documents).
Public Sub RB_DispatchTag(ByVal tagStr As String)
    Dim parts() As String
    parts = Split(tagStr, ":")

    On Error GoTo Failed
    Select Case parts(0)
        ' --- Align / dock / stretch / fill ---
        Case "AlignM":  AlignToMaster parts(1), False
        Case "AlignS":  AlignToMaster parts(1), True
        Case "Dock":    DockToMaster parts(1)
        Case "Stretch": StretchToMaster parts(1)
        Case "FillGap": FillGapToMaster parts(1)

        ' --- Arrange ---
        Case "Stack":    StackObjects parts(1), False
        Case "StackGap": StackObjects parts(1), True
        Case "Place":    PlaceOnSlide parts(1)
        Case "Matrix"
            If parts(1) = "O" Then ArrangeMatrix Else ArrangeMatrixQuick
        Case "Spacing":  SetSpacing parts(1)
        Case "Distrib":  DistributeObjects parts(1)
        Case "Swap":     SwapObjects parts(1), parts(2)
        Case "Golden":   GoldenCanon
        Case "Slice":    SliceShape
        Case "Multiply": MultiplyShape

        ' --- Size / shape / table ---
        Case "Size":     MatchSizeToMaster parts(1)
        Case "Magic":    MagicResizer
        Case "Angles":   AlignAnglesToMaster
        Case "ProcChain": AlignProcessChain
        Case "BlockArr": AlignBlockArrows
        Case "RoundRect": AlignRoundedRectangles
        Case "TblSnap":  SnapToTableCells parts(1)

        ' --- Colors ---
        Case "ThemeCol": ApplyThemeColor parts(1), CLng(parts(2))
        Case "PalCol":   ApplyPaletteColor parts(1), CLng(parts(2))
        Case "Theme2RGB": ThemeToRGB
        Case "RGB2Theme": RGBToTheme
        Case "PickCol":  PickColorsFromMaster parts(1)
        Case "ColInfo":  ShowColorInfo

        ' --- Text ---
        Case "TxtSplit":  SplitTextBox
        Case "TxtMerge":  MergeTextBoxes
        Case "TxtMargin": SetTextMargins
        Case "TxtFit":    FitFormToText
        Case "TxtWrap":   ToggleWrapText
        Case "TxtCase":   ChangeTextCase parts(1)
        Case "TxtTidy":   TidyText
        Case "TxtSwap":   SwapText

        ' --- View ---
        Case "HideObj":   HideSelectedObjects
        Case "UnhideObj": UnhideHiddenObjects
        Case "MasterObj": ToggleMasterShapes

        ' --- Review markup ---
        Case "Note":       AddReviewNote parts(1)
        Case "Stamp":      AddStatusStamp parts(1)
        Case "Callout":    AddReviewCallout
        Case "ReviewDel":  RemoveReviewMarkup
        Case "ReviewInit": SetReviewInitials

        ' --- Wizards / productivity ---
        Case "Painter":  AdvancedFormatPainter
        Case "SelSim":   SelectSimilarShapes parts(1)
        Case "Agenda":   GenerateAgenda
        Case "AgendaDel": RemoveAgenda
        Case "LibIns":   InsertLibraryElement CLng(parts(1))
        Case "LibAdd":   AddSelectionToLibrary
        Case "LibOpen":  OpenLibraryForEditing
        Case "FmtApply": ApplySavedFormat CLng(parts(1))
        Case "FmtSave":  SaveFormatFromMaster
        Case "FmtOpen":  ShowFormatsFile

        ' --- Expert / clean-up ---
        Case "Notes":    RemoveAllNotes
        Case "Anim":     RemoveAllAnimations
        Case "Designs":  DeleteUnusedDesigns
        Case "Summary":  CopySummaryToClipboard
        Case "PasteAll": PasteOnSelectedSlides
        Case "Extract":  ExtractSelectedSlides
        Case "Lang":     SetSpellLanguage parts(1), parts(2)

        ' --- Shortcuts ---
        Case "ShortMenu": ShortcutsButton

        ' --- Icon Aid ---  (browsing is the web task pane; VBA just materializes)
        Case "IconMakeEditable": MakeIconsEditable

        ' --- Chart Aid ---
        Case "Ch":       BuildChart parts(1)
        Case "ChRebuild": RebuildChart
        Case "ChEdit":   EditChartData
        Case "ChHelp":   ChartDataHelp
        Case "ChSamples": InsertChartSamples
        Case "ChSamplesDel": RemoveSampleSlides
        Case "ChDiff":   DifferenceArrow parts(1)
        Case "ChCAGR":   CagrArrow
        Case "ChVLine":  ValueLine
        Case "ChAvg":    AverageLine
        Case "ChColors": ChartColorsFile
        Case "ChSettings": ChartSettingsDialog
        Case "ChColors2": EditColorsDialog
        Case "ChStylePal": EditPaletteSwatches
        Case "ChStyleSet": InsertStyleTable
        Case "ChStyleSetK": InsertStyleTableForSelected
        Case "ChTheme":  ApplyColorTheme CLng(parts(1))
        Case "ChStyleApply": ApplyStyleFromSelection
        Case "ChStyleAll": RestyleAllCharts
        Case "ChReOne":  RestyleSelectedChart
        Case "ChStyleReset": ResetStyle
        Case "ChRecolor": RecolorSeries
        Case "ChHarvey": InsertHarveyBall
        Case "ChCheck":  InsertCheckbox
        Case "ChCycle":  CycleCheckbox
    End Select
    Exit Sub
Failed:
    MsgBox "Slide Aid: " & Err.Description, vbExclamation, "Slide Aid"
End Sub

' ---- color theme gallery ----
Public Sub RB_PalGallery(control As IRibbonControl, ByVal id As String, _
                         ByVal index As Integer)
    ApplyColorTheme index + 1          ' gallery index is 0-based
End Sub

' ---- dynamicMenu callbacks (My Elements / My Formats) ----
Public Sub RB_GetLibraryMenu(control As IRibbonControl, ByRef returnedVal)
    returnedVal = LibraryMenuXML()
End Sub

Public Sub RB_GetFormatsMenu(control As IRibbonControl, ByRef returnedVal)
    returnedVal = FormatsMenuXML()
End Sub
