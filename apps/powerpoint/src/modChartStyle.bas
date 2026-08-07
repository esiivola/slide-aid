Attribute VB_Name = "modChartStyle"
' =====================================================================
' Chart Aid - visual style system.
'
' COLORS: 'Edit Palette' inserts the current palette as swatch
' squares. Recolor them with PowerPoint's own fill tools (color
' wheel, theme colors, eyedropper), reorder/add/delete squares,
' keep them selected -> 'Apply from Selection'. Their fills become
' the chart palette (chartcolors.txt is written automatically).
'
' PARAMETERS: 'Edit Settings' inserts a table with every tunable
' (bar widths, gaps, label size, decimals, toggles, waterfall
' colors, plot size) pre-filled with current values. Edit values,
' select the table -> 'Apply from Selection'.
'
' 'Apply from Selection' is smart: swatches set the palette, a
' settings table sets the parameters - both at once also works.
' Existing charts pick up changes when rebuilt (Edit Data).
' =====================================================================
Option Explicit

Private sKeys() As String
Private sVals() As String
Private sN As Long
Private curKind As String       ' chart type being built; makes every
                                ' StyleNum/Str lookup check "<KIND>.Key"
                                ' before the global "Key" automatically

' ---- native Chart Settings dialog: live session state ----
' One row per control shown in the native panel. dlgVal holds the value
' in its STORED form (ratio for a %, hex or "theme" for a color, the raw
' string otherwise); the spec builder/return parser convert at the edges.
Private dlgN As Long
Private dlgCtrl() As String     ' num | pct | check | popup | color
Private dlgBase() As String     ' key without the "<KIND>." prefix
Private dlgKeys() As String     ' key as WRITTEN ("COL.ClusterFill" or global)
Private dlgLabel() As String
Private dlgVal() As String
Private dlgMin() As Double
Private dlgMax() As Double
Private dlgOpts() As String     ' pipe-joined popup options
Private dlgScope As String      ' "GLOBAL" or a chart KIND
Private dlgKind As String

' ---- native Edit Colors dialog: per-family palette session state ----
' Family index 1 = BARS, 2 = LINES, 3 = PIES. Families are lazy-loaded
' and dirty-tracked so Cancel discards and only edited families are
' written (an untouched Done never freezes a theme-following palette).
Private colFam As String
Private colSlot As String
Private colHex(1 To 3, 1 To 60) As String
Private colN(1 To 3) As Long
Private colLoaded(1 To 3) As Boolean
Private colDirty(1 To 3) As Boolean

' key, default, description - single source of truth
Private Function KeyDefs() As Variant
    KeyDefs = Array( _
        Array("PlotWidthCm", "12", "Default chart width (cm)"), _
        Array("PlotHeightCm", "8", "Default chart height (cm)"), _
        Array("ClusterFill", "0.72", "Bar group width as share of category slot (clustered)"), _
        Array("StackFill", "0.65", "Bar width as share of slot (stacked / 100%)"), _
        Array("WaterfallFill", "0.62", "Bar width as share of slot (waterfall)"), _
        Array("MekkoGapPt", "2", "Gap between Mekko columns (pt)"), _
        Array("LabelSizePt", "9", "Font size of chart labels (pt)"), _
        Array("Decimals", "auto", "Value decimals: auto, 0, 1 or 2"), _
        Array("ValueLabels", "1", "Show value labels (1 = yes, 0 = no)"), _
        Array("TotalLabels", "1", "Show totals on stacked columns (1/0)"), _
        Array("Legend", "1", "Show the series legend (1/0)"), _
        Array("MarkerSizePt", "5", "Line-chart marker size (pt)"), _
        Array("WaterfallUp", "9BBB59", "Waterfall: positive segments (hex)"), _
        Array("WaterfallDown", "C0504D", "Waterfall: negative segments (hex)"), _
        Array("WaterfallTotal", "BFBFBF", "Waterfall: subtotal bars (hex)"), _
        Array("GanttBarColor", "theme", "Gantt bars: hex color, or 'theme'"))
End Function

Public Function StylePath() As String
    StylePath = StoreDir() & "/chartstyle.txt"
End Function

' ---------------------------------------------------------------
Public Sub LoadStyle()
    sN = 0
    ReDim sKeys(1 To 50)
    ReDim sVals(1 To 50)
    If Dir(StylePath()) = "" Then Exit Sub
    Dim f As Integer, ln As String, p As Long
    f = FreeFile
    On Error GoTo Done                ' Open failed: nothing to close
    Open StylePath() For Input As #f
    On Error GoTo CloseIt             ' from here on the file IS open
    Do While Not EOF(f)
        Line Input #f, ln
        ln = Trim$(ln)
        p = InStr(ln, "=")
        If p > 1 And Left$(ln, 1) <> "#" And sN < 50 Then
            sN = sN + 1
            sKeys(sN) = Trim$(Left$(ln, p - 1))
            sVals(sN) = Trim$(Mid$(ln, p + 1))
        End If
    Loop
CloseIt:
    Close #f
Done:
    On Error GoTo 0
End Sub

Public Sub SetStyleKind(ByVal kind As String)
    curKind = kind
End Sub

Private Function RawLookup(ByVal key As String, ByRef found As Boolean) As String
    Dim i As Long
    found = False
    For i = 1 To sN
        If LCase$(sKeys(i)) = LCase$(key) Then
            RawLookup = sVals(i)
            found = True
            Exit Function
        End If
    Next i
End Function

' Per-type override ("COL.ClusterFill") wins over the global key.
Public Function StyleStr(ByVal key As String, ByVal dflt As String) As String
    Dim found As Boolean, v As String
    If Len(curKind) > 0 Then
        v = RawLookup(curKind & "." & key, found)
        If found Then
            StyleStr = v
            Exit Function
        End If
    End If
    v = RawLookup(key, found)
    If found Then StyleStr = v Else StyleStr = dflt
End Function

Public Function StyleNum(ByVal key As String, ByVal dflt As Double) As Double
    Dim s As String
    s = StyleStr(key, "")
    If Len(s) = 0 Then
        StyleNum = dflt
    Else
        StyleNum = Val(Replace(s, ",", "."))
    End If
End Function

Public Function StyleColor(ByVal key As String, ByVal dfltRGB As Long) As Long
    Dim ok As Boolean, v As Long
    v = ParseColorText(StyleStr(key, ""), ok)
    If ok Then StyleColor = v Else StyleColor = dfltRGB
End Function

' Chart label size (used throughout the builders)
Public Function LSz() As Single
    LSz = StyleNum("LabelSizePt", 9)
End Function

' ---------------------------------------------------------------
' COLOR THEMES: curated global palettes, shown visually in the
' ribbon gallery (strip images) and as menu fallback entries.
' Applying a theme writes all three group palettes and offers to
' restyle every existing chart. (Must match THEMES in make_icons.py.)
' ---------------------------------------------------------------
Private Function ThemeDef(ByVal i As Long) As Variant
    Select Case i
        Case 1: ThemeDef = Array("Office", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47")
        Case 2: ThemeDef = Array("Nordic Blue", "1F4E79", "2E75B6", "9DC3E6", "BDD7EE", "636363", "D9D9D9")
        Case 3: ThemeDef = Array("Fjord", "264653", "2A9D8F", "E9C46A", "F4A261", "E76F51", "8AB17D")
        Case 4: ThemeDef = Array("Forest", "1B4332", "2D6A4F", "40916C", "74C69D", "B7E4C7", "95D5B2")
        Case 5: ThemeDef = Array("Sunset", "073B4C", "118AB2", "06D6A0", "FFD166", "EF476F", "26547C")
        Case 6: ThemeDef = Array("Berry", "4A1942", "893168", "C05299", "E29ACD", "6F6F6F", "CFCFCF")
        Case 7: ThemeDef = Array("Greyscale", "212529", "495057", "6C757D", "ADB5BD", "CED4DA", "DEE2E6")
        Case 8: ThemeDef = Array("Financial", "00304D", "006BA6", "FFB81C", "97999B", "DA291C", "63666A")
        Case 9: ThemeDef = Array("Vivid", "3D348B", "7678ED", "F7B801", "F18701", "F35B04", "5F0F40")
    End Select
End Function

Public Sub ApplyColorTheme(ByVal i As Long)
    If i < 1 Or i > 9 Then Exit Sub
    Dim th As Variant
    th = ThemeDef(i)

    EnsureStore
    Dim grp As Variant, g As Long, f As Integer, c As Long
    grp = Array("BARS", "LINES", "PIES")
    For g = 0 To 2
        f = FreeFile
        Open GroupPalettePath(CStr(grp(g))) For Output As #f
        Print #f, "# Chart Aid palette - color theme '" & th(0) & "'"
        For c = 1 To 6
            Print #f, th(c)
        Next c
        Close #f
    Next g

    If MsgBox("Color theme '" & th(0) & "' applied to new charts." & vbCr & vbCr & _
              "Restyle all existing charts in this presentation now?", _
              vbYesNo + vbQuestion, "Chart Aid") = vbYes Then
        RestyleAllCharts False
    End If
End Sub

' ---------------------------------------------------------------
' NATIVE COLOR PICKER: opens the macOS color panel (wheel, sliders,
' eyedropper) via the SlideAidUI AppleScript helper. Falls back to
' a hex/R,G,B InputBox when the helper isn't installed.
' ---------------------------------------------------------------
Public Function NativePickColor(ByVal defaultRGB As Long, ByRef ok As Boolean) As Long
    ok = False
    Dim param As String, res As String
    param = (defaultRGB And &HFF) & "," & _
            ((defaultRGB \ &H100) And &HFF) & "," & _
            ((defaultRGB \ &H10000) And &HFF)

    On Error GoTo Fallback
    res = AppleScriptTask("SlideAidUI.scpt", "chooseColor", param)
    On Error GoTo 0
    If Len(res) = 0 Then Exit Function          ' user cancelled

    ' Defensive: an OUTDATED compiled helper returns raw 16-bit
    ' components (0-65535). Without this check VBA's RGB() clamps
    ' them to 255 each - every pick would come back as white.
    Dim q() As String
    q = Split(res, ",")
    If UBound(q) = 2 Then
        If Val(q(0)) > 255 Or Val(q(1)) > 255 Or Val(q(2)) > 255 Then
            res = CLng(Val(q(0)) / 257) & "," & _
                  CLng(Val(q(1)) / 257) & "," & _
                  CLng(Val(q(2)) / 257)
        End If
    End If
    NativePickColor = ParseColorText(res, ok)
    Exit Function

Fallback:
    ' helper not installed - text entry instead
    On Error GoTo 0
    Dim s As String
    s = InputBox("Color as hex (e.g. 1F497D) or R,G,B:" & vbCr & vbCr & _
                 "(Tip: install the native color picker - see README, " & _
                 "'SlideAidUI' - to get the macOS color wheel with " & _
                 "eyedropper here instead.)", "Chart Aid", "1F497D")
    NativePickColor = ParseColorText(s, ok)
End Function

' ---------------------------------------------------------------
' EDIT PALETTE: insert one recolorable swatch row per chart family
' (BARS / LINES / PIES). Each family = one color group: applying a
' row recolors ALL charts of that family (after Restyle All).
' ---------------------------------------------------------------
Public Sub EditPaletteSwatches()
    Dim groups As Variant, labels As Variant
    groups = Array("BARS", "LINES", "PIES")
    labels = Array("Bars (column, bar, stacked, 100%, Mekko, waterfall, Gantt)", _
                   "Lines (line, area, scatter, bubble)", _
                   "Pies (pie, doughnut)")

    Dim names(1 To 40) As String, nn As Long
    Dim g As Long, i As Long, s As Shape
    Dim y As Single
    y = SlideH() / 2 - 70

    For g = 0 To 2
        LoadChartPalette IIf(groups(g) = "BARS", "COL", IIf(groups(g) = "LINES", "LINE", "PIE"))
        Dim lab As Shape
        Set lab = CurrentSlide().Shapes.AddTextbox(msoTextOrientationHorizontal, _
                  SlideW() / 2 - 210, y + g * 52, 420, 14)
        With lab.TextFrame.TextRange
            .Text = labels(g)
            .Font.Size = 10
            .Font.Color.RGB = RGB(100, 100, 100)
        End With
        nn = nn + 1
        names(nn) = lab.Name
        For i = 1 To 6
            Set s = CurrentSlide().Shapes.AddShape(msoShapeRectangle, _
                    SlideW() / 2 - 210 + (i - 1) * 34, y + g * 52 + 16, 28, 28)
            s.Fill.ForeColor.RGB = ChartColorRGB(i)
            s.Line.ForeColor.RGB = RGB(160, 160, 160)
            s.Line.Weight = 0.75
            s.Tags.Add "SACH_SWGRP", CStr(groups(g))
            nn = nn + 1
            names(nn) = s.Name
        Next i
    Next g

    Dim arr() As String
    ReDim arr(1 To nn)
    For i = 1 To nn
        arr(i) = names(i)
    Next i
    CurrentSlide().Shapes.Range(arr).Select

    MsgBox "One swatch row per chart family (left to right = series 1, 2, 3...)." & vbCr & vbCr & _
           "Recolor with PowerPoint's fill tools - theme colors and the eyedropper " & _
           "work. Add, delete or reorder squares within a row." & vbCr & vbCr & _
           "Then: select the swatches (any or all rows) > Style > More > Apply " & _
           "from Selection, and Style > Restyle All to update existing charts." & vbCr & vbCr & _
           "(Tip: Style > Edit Colors does this in a native palette editor - no slide clutter.)", _
           vbInformation, "Chart Aid"
End Sub

' ---------------------------------------------------------------
' EDIT SETTINGS: insert the parameter table with current values.
' ---------------------------------------------------------------
Public Sub InsertStyleTable()
    LoadStyle
    Dim defs As Variant
    defs = KeyDefs()

    Dim nR As Long
    nR = UBound(defs) + 2                          ' + header row
    Dim tbl As Shape
    Set tbl = CurrentSlide().Shapes.AddTable(nR, 3, 30, 60, 460, nR * 17)

    SetCell tbl, 1, 1, "Parameter", True
    SetCell tbl, 1, 2, "Value", True
    SetCell tbl, 1, 3, "Meaning", True

    Dim i As Long
    For i = 0 To UBound(defs)
        SetCell tbl, i + 2, 1, defs(i)(0), False
        SetCell tbl, i + 2, 2, StyleStr(CStr(defs(i)(0)), CStr(defs(i)(1))), False
        SetCell tbl, i + 2, 3, defs(i)(2), False
    Next i

    tbl.Select
    MsgBox "Edit the Value column, keep the table selected, then click " & _
           "Style > More > Apply from Selection. Existing charts pick up the " & _
           "new style when rebuilt (Edit Data). Delete the table when done." & vbCr & vbCr & _
           "(Tip: Style > Chart Settings does this in a native panel - no magic numbers.)", _
           vbInformation, "Chart Aid"
End Sub

Private Sub SetCell(tbl As Shape, ByVal r As Long, ByVal c As Long, _
                    ByVal txt As String, ByVal bold As Boolean)
    On Error Resume Next
    Dim oCell As Object
    Set oCell = tbl.Table.Cell(r, c)
    With oCell.Shape.TextFrame.TextRange
        .Text = txt
        .Font.Size = 10
        .Font.Bold = bold
    End With
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------
' APPLY FROM SELECTION: swatches -> palette, settings table ->
' parameters. Both in one selection also works.
' ---------------------------------------------------------------
Public Sub ApplyStyleFromSelection()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim didPal As Boolean, didSet As Boolean
    Dim i As Long

    ' --- settings table? ---
    For i = 1 To sr.Count
        If sr(i).HasTable Then
            If ApplySettingsTable(sr(i)) Then didSet = True
        End If
    Next i

    ' --- swatches? Grouped by chart family tag; untagged filled
    '     shapes go to the legacy global palette. ---
    Dim groups As Variant
    groups = Array("BARS", "LINES", "PIES", "")
    Dim g As Long, n As Long, nTotal As Long
    For g = 0 To 3
        Dim idx() As Long
        ReDim idx(1 To sr.Count)
        n = 0
        For i = 1 To sr.Count
            If Not sr(i).HasTable Then
                On Error Resume Next
                If sr(i).Fill.Visible = msoTrue And _
                   sr(i).Tags("SACH_SWGRP") = CStr(groups(g)) Then
                    n = n + 1
                    idx(n) = i
                End If
                On Error GoTo 0
            End If
        Next i
        If n >= 2 Then
            Dim j As Long, t As Long
            For i = 1 To n - 1
                For j = 1 To n - i
                    If sr(idx(j)).Left > sr(idx(j + 1)).Left Then
                        t = idx(j): idx(j) = idx(j + 1): idx(j + 1) = t
                    End If
                Next j
            Next i
            EnsureStore
            Dim f As Integer, v As Long, path As String
            If Len(CStr(groups(g))) > 0 Then
                path = GroupPalettePath(CStr(groups(g)))
            Else
                path = ChartPalettePath()
            End If
            f = FreeFile
            Open path For Output As #f
            Print #f, "# Chart Aid palette - written by 'Apply from Selection'."
            For i = 1 To n
                v = sr(idx(i)).Fill.ForeColor.RGB
                Print #f, Right$("0" & Hex(v And &HFF), 2) & _
                          Right$("0" & Hex((v \ &H100) And &HFF), 2) & _
                          Right$("0" & Hex((v \ &H10000) And &HFF), 2)
            Next i
            Close #f
            nTotal = nTotal + n
            didPal = True
        End If
    Next g
    n = nTotal

    If didPal Or didSet Then
        Dim what As String
        If didPal And didSet Then
            what = "Palette and settings applied."
        ElseIf didPal Then
            what = "Palette applied (" & n & " colors, left to right)."
        Else
            what = "Settings applied."
        End If
        If MsgBox(what & vbCr & vbCr & _
                  "Restyle all existing charts in this presentation now?", _
                  vbYesNo + vbQuestion, "Chart Aid") = vbYes Then
            RestyleAllCharts False
        End If
    Else
        MsgBox "Select palette swatches (2+ filled shapes) and/or a settings " & _
               "table first - see Style > More. (Or use Style > Chart Settings / " & _
               "Edit Colors for the native panels.)", _
               vbExclamation, "Chart Aid"
    End If
End Sub

' Accepts global keys ("LabelSizePt") and per-type keys
' ("WF.WaterfallFill"). MERGES into the existing settings, so a
' per-type table doesn't wipe global values and vice versa.
Private Function ApplySettingsTable(tbl As Shape) As Boolean
    ApplySettingsTable = False
    Dim d As ChartData
    If Not ReadTable(tbl, d) Then Exit Function

    Dim r As Long, hits As Long, key As String
    LoadStyle                                     ' existing values -> sKeys/sVals
    For r = 0 To UBound(d.cells, 1)
        key = Trim$(d.cells(r, 0))
        If IsKnownKey(key) Then
            UpsertStyle key, Trim$(d.cells(r, 1))
            hits = hits + 1
        End If
    Next r
    If hits = 0 Then Exit Function

    SaveStyleFile
    ApplySettingsTable = True
End Function

' Rewrite chartstyle.txt from the in-memory sKeys/sVals (populated by
' LoadStyle + UpsertStyle). Shared by 'Apply from Selection' and the
' native Chart Settings dialog so neither clobbers unrelated keys.
Public Sub SaveStyleFile()
    EnsureStore
    Dim f As Integer, i As Long
    f = FreeFile
    Open StylePath() For Output As #f
    Print #f, "# Chart Aid settings - written by Chart Aid."
    For i = 1 To sN
        Print #f, sKeys(i) & "=" & sVals(i)
    Next i
    Close #f
End Sub

' "Key" or "KIND.Key" where Key is defined and KIND is a chart type.
Private Function IsKnownKey(ByVal key As String) As Boolean
    Dim base As String, p As Long
    p = InStr(key, ".")
    base = key
    If p > 1 Then
        Dim pre As String
        pre = UCase$(Left$(key, p - 1))
        If InStr(",COL,BAR,STK,SBR,PCT,MEK,WF,LINE,AREA,PIE,DON,SCAT,BUB,GANTT,", _
                 "," & pre & ",") = 0 Then Exit Function
        base = Mid$(key, p + 1)
    End If
    Dim defs As Variant, i As Long
    defs = KeyDefs()
    For i = 0 To UBound(defs)
        If LCase$(base) = LCase$(defs(i)(0)) Then
            IsKnownKey = True
            Exit Function
        End If
    Next i
End Function

Private Sub UpsertStyle(ByVal key As String, ByVal val As String)
    Dim i As Long
    For i = 1 To sN
        If LCase$(sKeys(i)) = LCase$(key) Then
            sVals(i) = val
            Exit Sub
        End If
    Next i
    If sN < 50 Then
        sN = sN + 1
        sKeys(sN) = key
        sVals(sN) = val
    End If
End Sub

' ---------------------------------------------------------------
' PER-TYPE SETTINGS: select a chart, get a table with only that
' chart type's relevant parameters (prefixed "KIND.Param"), edit,
' Apply from Selection. Overrides apply to that type only.
' ---------------------------------------------------------------
Public Sub InsertStyleTableForSelected()
    Dim g As Shape
    Set g = FindOldChart()
    If g Is Nothing Then
        MsgBox "Select a Chart Aid chart first (or use Edit Settings for the " & _
               "global table).", vbExclamation, "Chart Aid"
        Exit Sub
    End If

    Dim kind As String
    kind = g.Tags(TAG_TYPE)
    LoadStyle
    SetStyleKind kind

    Dim keys As Variant
    keys = TypeKeys(kind)
    If IsEmpty(keys) Then
        MsgBox "This chart type has no type-specific parameters.", _
               vbInformation, "Chart Aid"
        Exit Sub
    End If

    Dim nR As Long
    nR = UBound(keys) + 2
    Dim tbl As Shape
    Set tbl = CurrentSlide().Shapes.AddTable(nR, 3, 30, 60, 470, nR * 17)
    SetCell tbl, 1, 1, "Parameter (" & kind & " only)", True
    SetCell tbl, 1, 2, "Value", True
    SetCell tbl, 1, 3, "Meaning", True

    Dim i As Long, base As String
    For i = 0 To UBound(keys)
        base = keys(i)(0)
        SetCell tbl, i + 2, 1, kind & "." & base, False
        SetCell tbl, i + 2, 2, StyleStr(base, CStr(keys(i)(1))), False
        SetCell tbl, i + 2, 3, keys(i)(2), False
    Next i
    SetStyleKind ""

    tbl.Select
    MsgBox "These values will apply to " & kind & " charts only. Edit, keep " & _
           "the table selected, Style > More > Apply from Selection, then Restyle All." & vbCr & vbCr & _
           "(Tip: Style > Chart Settings edits this chart type in a native panel.)", _
           vbInformation, "Chart Aid"
End Sub

' Relevant parameter subset per chart type (key, default, meaning).
Private Function TypeKeys(ByVal kind As String) As Variant
    Select Case kind
        Case "COL", "BAR"
            TypeKeys = Array( _
                Array("ClusterFill", "0.72", "Bar group width as share of slot"), _
                Array("ValueLabels", "1", "Show value labels (1/0)"), _
                Array("Legend", "1", "Show series legend (1/0)"), _
                Array("LabelSizePt", "9", "Label font size (pt)"), _
                Array("Decimals", "auto", "auto, 0, 1 or 2"))
        Case "STK", "SBR", "PCT"
            TypeKeys = Array( _
                Array("StackFill", "0.65", "Bar width as share of slot"), _
                Array("ValueLabels", "1", "Show segment labels (1/0)"), _
                Array("TotalLabels", "1", "Show totals (1/0)"), _
                Array("Legend", "1", "Show series legend (1/0)"), _
                Array("LabelSizePt", "9", "Label font size (pt)"), _
                Array("Decimals", "auto", "auto, 0, 1 or 2"))
        Case "WF"
            TypeKeys = Array( _
                Array("WaterfallFill", "0.62", "Bar width as share of slot"), _
                Array("WaterfallUp", "9BBB59", "Positive segments (hex)"), _
                Array("WaterfallDown", "C0504D", "Negative segments (hex)"), _
                Array("WaterfallTotal", "BFBFBF", "Subtotal bars (hex)"), _
                Array("LabelSizePt", "9", "Label font size (pt)"), _
                Array("Decimals", "auto", "auto, 0, 1 or 2"))
        Case "MEK"
            TypeKeys = Array( _
                Array("MekkoGapPt", "2", "Gap between columns (pt)"), _
                Array("LabelSizePt", "9", "Label font size (pt)"), _
                Array("Decimals", "auto", "auto, 0, 1 or 2"))
        Case "LINE", "AREA"
            TypeKeys = Array( _
                Array("MarkerSizePt", "5", "Marker size (pt, LINE)"), _
                Array("ValueLabels", "1", "Show value labels (1/0)"), _
                Array("LabelSizePt", "9", "Label font size (pt)"), _
                Array("Decimals", "auto", "auto, 0, 1 or 2"))
        Case "GANTT"
            TypeKeys = Array( _
                Array("GanttBarColor", "theme", "Bar color: hex or 'theme'"), _
                Array("LabelSizePt", "9", "Label font size (pt)"))
        Case "PIE", "DON", "SCAT", "BUB"
            TypeKeys = Array( _
                Array("LabelSizePt", "9", "Label font size (pt)"), _
                Array("Decimals", "auto", "auto, 0, 1 or 2"))
    End Select
End Function

' ---------------------------------------------------------------
Public Sub ResetStyle()
    If MsgBox("Reset chart colors and settings to defaults?", _
              vbYesNo + vbQuestion, "Chart Aid") <> vbYes Then Exit Sub
    On Error Resume Next
    Kill ChartPalettePath()
    Kill GroupPalettePath("BARS")
    Kill GroupPalettePath("LINES")
    Kill GroupPalettePath("PIES")
    Kill StylePath()
    On Error GoTo 0
    sN = 0
    MsgBox "Style reset: theme colors and default parameters. " & _
           "Rebuild existing charts via Edit Data.", vbInformation, "Chart Aid"
End Sub

' =====================================================================
' NATIVE CHART SETTINGS  (modChartStyle: main new UI)
'
' Select a chart -> a native macOS panel of sliders / checkboxes /
' popups pre-filled with that chart type's parameters. Apply rebuilds
' it in place and re-opens (nudge-and-see); OK writes and closes. With
' nothing selected the panel edits the defaults for NEW charts.
'
' Falls back to the on-slide parameter table (InsertStyleTable /
' InsertStyleTableForSelected) if the SlideAidUI helper isn't installed,
' so the add-in still works standalone.
' =====================================================================
Public Sub ChartSettingsDialog()
    Dim g As Shape
    Set g = FindOldChart()
    ClearEditingTag                       ' drop any stale edit marker

    If g Is Nothing Then
        dlgScope = "GLOBAL": dlgKind = ""
    ElseIf Not KnownChartKind(g.Tags(TAG_TYPE)) Then
        dlgScope = "GLOBAL": dlgKind = ""     ' unknown -> global defaults
    Else
        dlgKind = g.Tags(TAG_TYPE)
        dlgScope = dlgKind
        g.Tags.Add "SACH_EDITING", "1"
    End If

    BuildDlgControls dlgScope, dlgKind

    Dim res As String, action As String
    Do
        On Error GoTo Fallback
        res = AppleScriptTask("SlideAidUI.scpt", "chartSettings", BuildSettingsSpec())
        On Error GoTo 0

        action = ParseSettingsReturn(res)     ' updates dlgVal; returns token
        Select Case action
            Case "CANCEL"
                Exit Do
            Case "COLORS"
                DoDialogColorPick                 ' one color, then re-open
            Case "APPLY"
                WriteDlgSettings
                If dlgScope <> "GLOBAL" Then RebuildEditingChart
            Case "OK"
                WriteDlgSettings
                If dlgScope = "GLOBAL" Then
                    OfferRestyleAll "Defaults saved (used by new charts)."
                Else
                    RebuildEditingChart
                End If
                Exit Do
            Case Else
                Exit Do
        End Select
    Loop
    ClearEditingTag
    Exit Sub

Fallback:
    On Error GoTo 0
    ClearEditingTag
    If dlgScope = "GLOBAL" Then InsertStyleTable Else InsertStyleTableForSelected
End Sub

Private Sub OfferRestyleAll(ByVal msg As String)
    If MsgBox(msg & vbCr & vbCr & _
              "Restyle all existing charts in this presentation now?", _
              vbYesNo + vbQuestion, "Chart Aid") = vbYes Then RestyleAllCharts False
End Sub

' ---- build the control set + current values for the active scope ----
Private Sub BuildDlgControls(ByVal scope As String, ByVal kind As String)
    Dim defs As Variant
    If scope = "GLOBAL" Then defs = GlobalControlDefs() Else defs = KindControlDefs(kind)

    dlgN = 0
    ReDim dlgCtrl(1 To 40)
    ReDim dlgBase(1 To 40)
    ReDim dlgKeys(1 To 40)
    ReDim dlgLabel(1 To 40)
    ReDim dlgVal(1 To 40)
    ReDim dlgMin(1 To 40)
    ReDim dlgMax(1 To 40)
    ReDim dlgOpts(1 To 40)

    LoadStyle
    If scope <> "GLOBAL" Then SetStyleKind kind

    Dim i As Long, base As String, ctrl As String, dflt As String
    For i = 0 To UBound(defs)
        dlgN = dlgN + 1
        ctrl = defs(i)(0): base = defs(i)(1)
        dlgCtrl(dlgN) = ctrl
        dlgBase(dlgN) = base
        dlgLabel(dlgN) = defs(i)(2)
        dlgMin(dlgN) = defs(i)(3)
        dlgMax(dlgN) = defs(i)(4)
        dlgOpts(dlgN) = defs(i)(5)
        dflt = defs(i)(6)

        If Left$(base, 2) = "__" Then
            dlgKeys(dlgN) = base                       ' synthetic, never written
        ElseIf scope = "GLOBAL" Then
            dlgKeys(dlgN) = base
        Else
            dlgKeys(dlgN) = kind & "." & base
        End If

        Select Case base
            Case "__GanttTheme"
                dlgVal(dlgN) = IIf(LCase$(StyleStr("GanttBarColor", "theme")) = "theme", "1", "0")
            Case "GanttBarColor"
                Dim gv As String
                gv = StyleStr("GanttBarColor", "theme")
                If IsHex6(gv) Then dlgVal(dlgN) = gv Else dlgVal(dlgN) = RGBToHex6(ChartColorRGB(1))
            Case Else
                dlgVal(dlgN) = StyleStr(base, dflt)
        End Select
    Next i

    If scope <> "GLOBAL" Then SetStyleKind ""
    dlgScope = scope: dlgKind = kind
End Sub

' ---- serialize the controls to the helper's request format ----
Private Function BuildSettingsSpec() As String
    Dim title As String, info As String
    If dlgScope = "GLOBAL" Then
        title = "Chart defaults (new charts)"
        info = "Applied to charts you build next. Select a chart to edit its own settings."
    Else
        title = KindDisplayName(dlgKind) & " chart settings"
        info = "Apply rebuilds this chart in place and keeps the panel open. OK closes."
    End If

    Dim s As String, i As Long
    s = "#" & vbTab & title & vbTab & info
    For i = 1 To dlgN
        s = s & vbLf & SpecLine(i)
    Next i
    BuildSettingsSpec = s
End Function

Private Function SpecLine(ByVal i As Long) As String
    Dim k As String, lbl As String
    k = dlgKeys(i): lbl = dlgLabel(i)
    Select Case dlgCtrl(i)
        Case "num"
            SpecLine = "num" & vbTab & k & vbTab & lbl & vbTab & _
                       CStr(CLng(Val(Replace(dlgVal(i), ",", ".")))) & vbTab & _
                       CStr(CLng(dlgMin(i))) & vbTab & CStr(CLng(dlgMax(i)))
        Case "pct"
            SpecLine = "num" & vbTab & k & vbTab & lbl & vbTab & _
                       CStr(CLng(Val(Replace(dlgVal(i), ",", ".")) * 100)) & vbTab & _
                       CStr(CLng(dlgMin(i))) & vbTab & CStr(CLng(dlgMax(i)))
        Case "check"
            SpecLine = "check" & vbTab & k & vbTab & lbl & vbTab & IIf(dlgVal(i) = "1", "1", "0")
        Case "popup"
            SpecLine = "popup" & vbTab & k & vbTab & lbl & vbTab & dlgVal(i) & _
                       vbTab & vbTab & vbTab & dlgOpts(i)
        Case "color"
            SpecLine = "swatch" & vbTab & k & vbTab & lbl & vbTab & ColorHexForDisplay(i)
    End Select
End Function

Private Function ColorHexForDisplay(ByVal i As Long) As String
    If IsHex6(dlgVal(i)) Then
        ColorHexForDisplay = dlgVal(i)
    Else
        ColorHexForDisplay = RGBToHex6(ChartColorRGB(1))   ' "theme" -> a real swatch
    End If
End Function

' ---- read the helper's reply back into dlgVal; return the action ----
Private Function ParseSettingsReturn(ByVal res As String) As String
    Dim lines() As String, j As Long, ln As String, p As Long, k As String, v As String
    lines = Split(Replace(res, vbCr, vbLf), vbLf)
    If UBound(lines) < 0 Then ParseSettingsReturn = "CANCEL": Exit Function
    ParseSettingsReturn = Trim$(lines(0))
    If ParseSettingsReturn = "CANCEL" Or Len(ParseSettingsReturn) = 0 Then _
        ParseSettingsReturn = "CANCEL": Exit Function

    For j = 1 To UBound(lines)
        ln = lines(j)
        p = InStr(ln, "=")
        If p > 1 Then
            k = Left$(ln, p - 1)
            v = Mid$(ln, p + 1)
            Dim i As Long
            For i = 1 To dlgN
                If dlgKeys(i) = k Then
                    If dlgCtrl(i) = "pct" Then
                        dlgVal(i) = RatioStr(Val(Replace(v, ",", ".")))
                    Else
                        dlgVal(i) = Trim$(v)
                    End If
                    Exit For
                End If
            Next i
        End If
    Next j
End Function

' ---- "Colors..." button: pick ONE color via the native panel ----
Private Sub DoDialogColorPick()
    Dim param As String, i As Long
    param = "Which color to change?"
    For i = 1 To dlgN
        If dlgCtrl(i) = "color" Then _
            param = param & vbLf & dlgLabel(i) & vbTab & ColorHexForDisplay(i)
    Next i

    Dim res As String
    On Error Resume Next
    res = AppleScriptTask("SlideAidUI.scpt", "chooseChartColor", param)
    On Error GoTo 0
    If Len(res) = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(Replace(res, vbCr, ""), vbTab)
    If UBound(parts) < 1 Then Exit Sub
    Dim pickedLabel As String, pickedHex As String
    pickedLabel = parts(0): pickedHex = Trim$(parts(1))
    If Not IsHex6(pickedHex) Then Exit Sub

    For i = 1 To dlgN
        If dlgCtrl(i) = "color" And dlgLabel(i) = pickedLabel Then
            dlgVal(i) = pickedHex
            SetDlgVal "__GanttTheme", "0"      ' a custom pick overrides "theme"
            Exit For
        End If
    Next i
End Sub

' ---- persist the edited values (merge; never clobber other keys) ----
Private Sub WriteDlgSettings()
    LoadStyle                                  ' current file -> sKeys/sVals
    Dim themeOn As Boolean
    themeOn = (GetDlgVal("__GanttTheme") = "1")

    Dim i As Long, wv As String
    For i = 1 To dlgN
        If Left$(dlgKeys(i), 2) <> "__" Then
            If dlgBase(i) = "GanttBarColor" And themeOn Then
                wv = "theme"
            Else
                wv = dlgVal(i)
            End If
            UpsertStyle dlgKeys(i), wv
        End If
    Next i
    SaveStyleFile
End Sub

' ---- tiny dlg-state helpers ----
Private Function GetDlgVal(ByVal base As String) As String
    Dim i As Long
    For i = 1 To dlgN
        If dlgBase(i) = base Then GetDlgVal = dlgVal(i): Exit Function
    Next i
End Function

Private Sub SetDlgVal(ByVal base As String, ByVal v As String)
    Dim i As Long
    For i = 1 To dlgN
        If dlgBase(i) = base Then dlgVal(i) = v: Exit Sub
    Next i
End Sub

Private Function IsHex6(ByVal s As String) As Boolean
    Dim ok As Boolean
    ParseColorText s, ok
    IsHex6 = ok
End Function

' ratio (0..1) -> a period-decimal string for chartstyle.txt
Private Function RatioStr(ByVal pct As Double) As String
    RatioStr = Replace(Format$(pct / 100, "0.00"), ",", ".")
End Function

Private Function KindDisplayName(ByVal kind As String) As String
    Select Case kind
        Case "COL": KindDisplayName = "Column"
        Case "BAR": KindDisplayName = "Bar"
        Case "STK": KindDisplayName = "Stacked column"
        Case "SBR": KindDisplayName = "Stacked bar"
        Case "PCT": KindDisplayName = "100% column"
        Case "WF": KindDisplayName = "Waterfall"
        Case "MEK": KindDisplayName = "Mekko"
        Case "LINE": KindDisplayName = "Line"
        Case "AREA": KindDisplayName = "Area"
        Case "PIE": KindDisplayName = "Pie"
        Case "DON": KindDisplayName = "Doughnut"
        Case "SCAT", "BUB": KindDisplayName = "Scatter / bubble"
        Case "GANTT": KindDisplayName = "Gantt"
        Case Else: KindDisplayName = "Chart"
    End Select
End Function

' ---- control definitions: {ctrl, base-key, label, min, max, opts, default} ----
Private Function GlobalControlDefs() As Variant
    GlobalControlDefs = Array( _
        Array("num", "PlotWidthCm", "Default width (cm)", 4, 30, "", "12"), _
        Array("num", "PlotHeightCm", "Default height (cm)", 3, 20, "", "8"), _
        Array("num", "LabelSizePt", "Label size (pt)", 5, 24, "", "9"), _
        Array("popup", "Decimals", "Decimals", 0, 0, "auto|0|1|2", "auto"), _
        Array("check", "ValueLabels", "Show value labels", 0, 0, "", "1"), _
        Array("check", "TotalLabels", "Show totals on stacked columns", 0, 0, "", "1"), _
        Array("check", "Legend", "Show legend (charts with 2+ series)", 0, 0, "", "1"))
End Function

' Curated to the parameters each builder actually reads (e.g. AREA reads
' only the label size; LINE also uses markers / value labels / decimals).
Private Function KindControlDefs(ByVal kind As String) As Variant
    Select Case kind
        Case "COL", "BAR"
            KindControlDefs = Array( _
                Array("pct", "ClusterFill", "Bar width (%)", 30, 100, "", "0.72"), _
                Array("check", "ValueLabels", "Show value labels", 0, 0, "", "1"), _
                Array("check", "Legend", "Show legend (2+ series)", 0, 0, "", "1"), _
                Array("num", "LabelSizePt", "Label size (pt)", 5, 24, "", "9"), _
                Array("popup", "Decimals", "Decimals", 0, 0, "auto|0|1|2", "auto"))
        Case "STK", "SBR", "PCT"
            KindControlDefs = Array( _
                Array("pct", "StackFill", "Bar width (%)", 30, 100, "", "0.65"), _
                Array("check", "ValueLabels", "Show segment labels", 0, 0, "", "1"), _
                Array("check", "TotalLabels", "Show totals", 0, 0, "", "1"), _
                Array("check", "Legend", "Show legend (2+ series)", 0, 0, "", "1"), _
                Array("num", "LabelSizePt", "Label size (pt)", 5, 24, "", "9"), _
                Array("popup", "Decimals", "Decimals", 0, 0, "auto|0|1|2", "auto"))
        Case "WF"
            KindControlDefs = Array( _
                Array("pct", "WaterfallFill", "Bar width (%)", 30, 100, "", "0.62"), _
                Array("color", "WaterfallUp", "Positive color", 0, 0, "", "9BBB59"), _
                Array("color", "WaterfallDown", "Negative color", 0, 0, "", "C0504D"), _
                Array("color", "WaterfallTotal", "Subtotal color", 0, 0, "", "BFBFBF"), _
                Array("num", "LabelSizePt", "Label size (pt)", 5, 24, "", "9"), _
                Array("popup", "Decimals", "Decimals", 0, 0, "auto|0|1|2", "auto"))
        Case "MEK"
            KindControlDefs = Array( _
                Array("num", "MekkoGapPt", "Column gap (pt)", 0, 20, "", "2"), _
                Array("num", "LabelSizePt", "Label size (pt)", 5, 24, "", "9"), _
                Array("popup", "Decimals", "Decimals", 0, 0, "auto|0|1|2", "auto"))
        Case "LINE"
            KindControlDefs = Array( _
                Array("num", "MarkerSizePt", "Marker size (pt)", 0, 12, "", "5"), _
                Array("check", "ValueLabels", "Show value labels", 0, 0, "", "1"), _
                Array("num", "LabelSizePt", "Label size (pt)", 5, 24, "", "9"), _
                Array("popup", "Decimals", "Decimals", 0, 0, "auto|0|1|2", "auto"))
        Case "AREA"
            KindControlDefs = Array( _
                Array("num", "LabelSizePt", "Label size (pt)", 5, 24, "", "9"))
        Case "GANTT"
            KindControlDefs = Array( _
                Array("check", "__GanttTheme", "Use theme color for bars", 0, 0, "", "1"), _
                Array("color", "GanttBarColor", "Bar color", 0, 0, "", "4472C4"), _
                Array("num", "LabelSizePt", "Label size (pt)", 5, 24, "", "9"))
        Case Else                                  ' PIE, DON, SCAT, BUB
            KindControlDefs = Array( _
                Array("num", "LabelSizePt", "Label size (pt)", 5, 24, "", "9"), _
                Array("popup", "Decimals", "Decimals", 0, 0, "auto|0|1|2", "auto"))
    End Select
End Function

' =====================================================================
' NATIVE EDIT COLORS  (replaces the drop-swatches-on-the-slide flow)
'
' A native palette editor: pick a chart family, then change / add /
' remove colors as real swatches. Done writes the changed family
' palette file(s) and offers Restyle All. Falls back to the on-slide
' swatch rows (EditPaletteSwatches) if the helper isn't installed.
' =====================================================================
Public Sub EditColorsDialog()
    Dim g As Shape, fam As String
    Set g = FindOldChart()
    If g Is Nothing Then
        fam = "BARS"
    ElseIf Not KnownChartKind(g.Tags(TAG_TYPE)) Then
        fam = "BARS"
    Else
        fam = PaletteGroupOf(g.Tags(TAG_TYPE))
    End If

    Dim k As Long
    For k = 1 To 3
        colLoaded(k) = False: colDirty(k) = False: colN(k) = 0
    Next k
    colFam = fam
    LoadColFamily FamIdx(fam)

    Dim res As String, action As String
    Do
        On Error GoTo Fallback
        res = AppleScriptTask("SlideAidUI.scpt", "editColors", BuildColorsSpec())
        On Error GoTo 0

        action = ParseColorsReturn(res)       ' may switch colFam; sets colSlot
        Select Case action
            Case "CANCEL": Exit Do
            Case "SWITCH": LoadColFamily FamIdx(colFam)
            Case "CHANGE": DoPaletteAction
            Case "DONE": WriteDirtyFamilies: Exit Do
            Case Else: Exit Do
        End Select
    Loop
    Exit Sub

Fallback:
    On Error GoTo 0
    EditPaletteSwatches
End Sub

Private Function BuildColorsSpec() As String
    Dim idx As Long, i As Long, s As String
    idx = FamIdx(colFam)
    s = "#" & vbTab & ("Chart colors - " & FamLabel(colFam)) & vbTab & colFam & vbTab & "BARS|LINES|PIES"
    For i = 1 To colN(idx)
        s = s & vbLf & "swatch" & vbTab & CStr(i) & vbTab & colHex(idx, i)
    Next i
    BuildColorsSpec = s
End Function

Private Function ParseColorsReturn(ByVal res As String) As String
    Dim lines() As String, j As Long, ln As String, p As Long, k As String, v As String
    Dim tok As String, fam As String, slot As String
    lines = Split(Replace(res, vbCr, vbLf), vbLf)
    If UBound(lines) < 0 Then ParseColorsReturn = "CANCEL": Exit Function
    tok = Trim$(lines(0))
    If Len(tok) = 0 Or tok = "CANCEL" Then ParseColorsReturn = "CANCEL": Exit Function

    fam = colFam: slot = ""
    For j = 1 To UBound(lines)
        ln = lines(j)
        p = InStr(ln, "=")
        If p > 1 Then
            k = Left$(ln, p - 1): v = Mid$(ln, p + 1)
            If k = "FAMILY" Then fam = Trim$(v)
            If k = "SLOT" Then slot = Trim$(v)
        End If
    Next j
    colSlot = slot

    If fam <> colFam And (fam = "BARS" Or fam = "LINES" Or fam = "PIES") Then
        colFam = fam
        ParseColorsReturn = "SWITCH"
    Else
        ParseColorsReturn = tok
    End If
End Function

Private Sub DoPaletteAction()
    Dim idx As Long: idx = FamIdx(colFam)
    Select Case colSlot
        Case "+ Add a color"
            If colN(idx) < 60 Then
                colN(idx) = colN(idx) + 1
                colHex(idx, colN(idx)) = RGBToHex6(AccentColorRGB(colN(idx)))
                colDirty(idx) = True
            End If
        Case "- Remove last color"
            If colN(idx) > 1 Then colN(idx) = colN(idx) - 1: colDirty(idx) = True
        Case "* Reset to theme colors"
            Dim i As Long
            For i = 1 To 6
                colHex(idx, i) = RGBToHex6(AccentColorRGB(i))
            Next i
            colN(idx) = 6: colDirty(idx) = True
        Case Else
            If Left$(colSlot, 6) = "Color " Then
                Dim n As Long: n = CLng(Val(Mid$(colSlot, 7)))
                If n >= 1 And n <= colN(idx) Then
                    Dim ok As Boolean, seed As Long, newC As Long
                    seed = ParseColorText(colHex(idx, n), ok)
                    If Not ok Then seed = RGB(79, 129, 189)
                    newC = NativePickColor(seed, ok)
                    If ok Then colHex(idx, n) = RGBToHex6(newC): colDirty(idx) = True
                End If
            End If
    End Select
End Sub

Private Sub WriteDirtyFamilies()
    Dim k As Long, i As Long, f As Integer, any As Boolean
    For k = 1 To 3
        If colLoaded(k) And colDirty(k) And colN(k) >= 1 Then
            EnsureStore
            f = FreeFile
            Open GroupPalettePath(FamName(k)) For Output As #f
            Print #f, "# Chart Aid palette - written by Edit Colors."
            For i = 1 To colN(k)
                Print #f, colHex(k, i)
            Next i
            Close #f
            any = True
        End If
    Next k
    If any Then OfferRestyleAll "Palette updated."
End Sub

' Load a family's palette (hex strings). Seeds from the theme accents
' when no palette file exists yet, so the editor always shows colors.
Private Sub LoadColFamily(ByVal idx As Long)
    If colLoaded(idx) Then Exit Sub
    Dim path As String, f As Integer, ln As String, v As Long, ok As Boolean
    colN(idx) = 0
    path = GroupPalettePath(FamName(idx))
    If Dir(path) <> "" Then
        f = FreeFile
        On Error GoTo Seed
        Open path For Input As #f
        On Error GoTo CloseIt
        Do While Not EOF(f) And colN(idx) < 60
            Line Input #f, ln
            v = ParseColorText(ln, ok)
            If ok Then
                colN(idx) = colN(idx) + 1
                colHex(idx, colN(idx)) = RGBToHex6(v)
            End If
        Loop
CloseIt:
        Close #f
    End If
Seed:
    On Error GoTo 0
    If colN(idx) = 0 Then
        Dim i As Long
        For i = 1 To 6
            colHex(idx, i) = RGBToHex6(AccentColorRGB(i))
        Next i
        colN(idx) = 6
    End If
    colLoaded(idx) = True
End Sub

Private Function FamIdx(ByVal fam As String) As Long
    Select Case fam
        Case "LINES": FamIdx = 2
        Case "PIES": FamIdx = 3
        Case Else: FamIdx = 1        ' BARS
    End Select
End Function

Private Function FamName(ByVal idx As Long) As String
    Select Case idx
        Case 2: FamName = "LINES"
        Case 3: FamName = "PIES"
        Case Else: FamName = "BARS"
    End Select
End Function

Private Function FamLabel(ByVal fam As String) As String
    Select Case fam
        Case "LINES": FamLabel = "Lines (line, area, scatter, bubble)"
        Case "PIES": FamLabel = "Pies (pie, doughnut)"
        Case Else: FamLabel = "Bars (column, bar, stacked, Mekko, waterfall, Gantt)"
    End Select
End Function
