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
           "Then: select the swatches (any or all rows) > Style > Apply from " & _
           "Selection, and Style > Restyle All Charts to update existing charts.", _
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
           "Style > Apply from Selection. Existing charts pick up the new " & _
           "style when rebuilt (Edit Data). Delete the table when done.", _
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
               "table first - see Style > Edit Palette / Edit Settings.", _
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

    EnsureStore
    Dim f As Integer, i As Long
    f = FreeFile
    Open StylePath() For Output As #f
    Print #f, "# Chart Aid settings - written by 'Apply from Selection'."
    For i = 1 To sN
        Print #f, sKeys(i) & "=" & sVals(i)
    Next i
    Close #f
    ApplySettingsTable = True
End Function

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
           "the table selected, Style > Apply from Selection, then Restyle.", _
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
