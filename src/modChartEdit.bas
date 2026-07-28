Attribute VB_Name = "modChartEdit"
' =====================================================================
' Chart Aid - the edit round-trip and data-layout help.
' Charts carry their data in tags. "Edit Data" recreates the data
' table next to the chart; edit it, select table + chart, and click
' the chart button again to rebuild in place.
' =====================================================================
Option Explicit

Public Sub EditChartData()
    ' NB: VBA's Or does not short-circuit, so the Nothing test and the
    ' tag test must stay in separate Ifs.
    Dim g As Shape
    Set g = FindOldChart()
    If g Is Nothing Then
        MsgBox "Select a chart created by Chart Aid first.", vbExclamation, "Chart Aid"
        Exit Sub
    End If
    If Len(g.Tags(TAG_DATA)) = 0 Then
        MsgBox "Select a chart created by Chart Aid first.", vbExclamation, "Chart Aid"
        Exit Sub
    End If

    Dim d As ChartData
    DeserializeData g.Tags(TAG_DATA), d

    Dim nR As Long, nC As Long
    nR = UBound(d.cells, 1) + 1
    nC = UBound(d.cells, 2) + 1

    Dim tl As Single
    tl = g.Left - (nC * 55) - 15
    If tl < 5 Then tl = 5

    Dim tbl As Shape
    Set tbl = CurrentSlide().Shapes.AddTable(nR, nC, tl, g.Top, nC * 55, nR * 18)
    tbl.Tags.Add "SACH_DATASHEET", "1"   ' auto-removed after the rebuild

    ' Cell.Shape is missing from the Mac object library - late-bind
    Dim r As Long, c As Long, oCell As Object
    For r = 1 To nR
        For c = 1 To nC
            On Error Resume Next
            Set oCell = tbl.Table.Cell(r, c)
            With oCell.Shape.TextFrame.TextRange
                .Text = d.cells(r - 1, c - 1)
                .Font.Size = 10
            End With
            On Error GoTo 0
        Next c
    Next r

    tbl.Select
    MsgBox "Edit the values, then select the table AND the chart and click " & _
           "Rebuild (or any chart button) - the chart is rebuilt in place " & _
           "and this table is removed automatically.", vbInformation, "Chart Aid"
End Sub

' ---------------------------------------------------------------
' RESTYLE THIS CHART: rebuild the selected chart with the current
' style (data, position and size are kept).
' ---------------------------------------------------------------
Public Sub RestyleSelectedChart()
    Dim g As Shape
    Set g = FindOldChart()
    If g Is Nothing Then
        MsgBox "Select a chart created by Chart Aid first.", vbExclamation, "Chart Aid"
        Exit Sub
    End If
    If Len(g.Tags(TAG_DATA)) = 0 Then
        MsgBox "Select a chart created by Chart Aid first.", vbExclamation, "Chart Aid"
        Exit Sub
    End If

    Dim kind As String, dataS As String
    Dim l As Single, t As Single, w As Single, h As Single
    kind = g.Tags(TAG_TYPE)
    dataS = g.Tags(TAG_DATA)
    If Not KnownChartKind(kind) Then          ' never delete what we
        MsgBox "Unknown chart type '" & kind & "' - cannot rebuild.", _
               vbExclamation, "Chart Aid"     ' cannot rebuild
        Exit Sub
    End If
    Dim ovr As String
    ovr = g.Tags(TAG_COLOVR)
    RectFromOldChart g, l, t, w, h      ' original build rect, not the
    g.Delete                            ' bbox - avoids size drift

    Dim d As ChartData
    DeserializeData dataS, d
    If BuildChartFromData(kind, d, l, t, w, h) Then ApplyOverridesToNewChart ovr
End Sub

' ---------------------------------------------------------------
' RESTYLE ALL: rebuild every Chart Aid chart in the presentation
' from its stored data, so palette and settings changes propagate
' with one click (charts keep their position and size).
' ---------------------------------------------------------------
Public Sub RestyleAllCharts(Optional ByVal confirm As Boolean = True)
    Dim p As Presentation
    Set p = ActivePresentation

    ' collect first - rebuilding modifies the shape collections
    Dim slideIdx(1 To 500) As Long, shpName(1 To 500) As String, n As Long
    Dim sl As Slide, s As Shape
    For Each sl In p.Slides
        For Each s In sl.Shapes
            If Len(s.Tags(TAG_TYPE)) > 0 And n < 500 Then
                n = n + 1
                slideIdx(n) = sl.SlideIndex
                shpName(n) = s.Name
            End If
        Next s
    Next sl

    If n = 0 Then
        MsgBox "No Chart Aid charts found in this presentation.", _
               vbInformation, "Chart Aid"
        Exit Sub
    End If
    If confirm Then
        If MsgBox("Rebuild " & n & " chart(s) with the current style? " & _
                  "Charts keep their data, position and size.", _
                  vbYesNo + vbQuestion, "Chart Aid") <> vbYes Then Exit Sub
    End If

    Dim i As Long, done As Long
    For i = 1 To n
        On Error Resume Next
        Set sl = p.Slides(slideIdx(i))
        Set s = sl.Shapes(shpName(i))
        If Not s Is Nothing Then
            Dim kind As String, dataS As String
            Dim l As Single, t As Single, w As Single, h As Single
            kind = s.Tags(TAG_TYPE)
            dataS = s.Tags(TAG_DATA)
            RectFromOldChart s, l, t, w, h   ' original rect - no size drift
            If Len(kind) > 0 And Len(dataS) > 0 And KnownChartKind(kind) Then
                Dim ovr As String
                ovr = s.Tags(TAG_COLOVR)
                ActiveWindow.View.GotoSlide sl.SlideIndex
                s.Delete
                Dim d As ChartData
                DeserializeData dataS, d
                Dim builtOK As Boolean
                builtOK = BuildChartFromData(kind, d, l, t, w, h)
                If builtOK Then ApplyOverridesToNewChart ovr
                If builtOK And Err.Number = 0 Then done = done + 1
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next i

    MsgBox done & " of " & n & " chart(s) rebuilt with the current style.", _
           vbInformation, "Chart Aid"
End Sub

' ---------------------------------------------------------------
' CHART COLORS: open (create if needed) the custom chart palette.
' One color per line, hex or R,G,B; # for comments. If the file
' has colors, new charts use them instead of the theme accents.
' ---------------------------------------------------------------
Public Sub ChartColorsFile()
    EnsureStore
    Dim p As String
    p = ChartPalettePath()

    If Dir(p) = "" Then
        ' seed with the current theme's accents so editing = tweaking
        Dim f As Integer, i As Long, v As Long
        f = FreeFile
        Open p For Output As #f
        Print #f, "# Chart Aid palette - one color per line (hex or R,G,B)."
        Print #f, "# Delete this file to fall back to the theme accents."
        For i = 1 To 6
            v = AccentColorRGB(i)
            Print #f, Right$("0" & Hex(v And &HFF), 2) & _
                      Right$("0" & Hex((v \ &H100) And &HFF), 2) & _
                      Right$("0" & Hex((v \ &H10000) And &HFF), 2)
        Next i
        Close #f
    End If

    On Error GoTo ShowPath
    ActivePresentation.FollowHyperlink "file://" & p
    Exit Sub
ShowPath:
    MsgBox "Edit this file (colors apply to the NEXT chart build; " & _
           "rebuild existing charts via Edit Data):" & vbCr & p, _
           vbInformation, "Chart Aid"
End Sub

' ---------------------------------------------------------------
' SAMPLE SLIDES: append one example slide per chart type, each with
' a correctly formatted data table AND the chart built from it by
' the real chart code. Living documentation - copy any table as a
' starting point for your own chart.
' ---------------------------------------------------------------
Public Sub InsertChartSamples()
    If MsgBox("Insert 14 sample slides (one per chart type) at the end " & _
              "of this presentation?" & vbCr & vbCr & _
              "(Remove them again anytime: Slide Aid > Clean-up > " & _
              "Remove Chart Sample Slides.)", _
              vbYesNo + vbQuestion, "Chart Aid") <> vbYes Then Exit Sub

    Const GRID_NOTE As String = "Table: row 1 = categories (top-left free), " & _
                                "column 1 = series names, body = numbers. "

    AddSample "Column - Revenue by region (EUR m)", _
        GRID_NOTE & "Use to COMPARE a few series across periods - here: Asia is the growth engine.", _
        Array(Array("", "2023", "2024", "2025", "2026"), _
              Array("Europe", "42", "48", "55", "61"), _
              Array("Americas", "35", "39", "46", "58"), _
              Array("Asia", "18", "26", "37", "52")), "COL"

    AddSample "Bar - Units sold by product (k)", _
        GRID_NOTE & "Use for RANKINGS: sort the categories by value and the winner reads top-down.", _
        Array(Array("", "Alpha", "Bravo", "Charlie", "Delta", "Echo"), _
              Array("Units", "84", "66", "45", "31", "12")), "BAR"

    AddSample "Stacked Column - Revenue by product line (EUR m)", _
        GRID_NOTE & "Use to show TOTAL GROWTH and its COMPOSITION at once - here: services drive the growth.", _
        Array(Array("", "2024", "2025", "2026"), _
              Array("Hardware", "50", "48", "45"), _
              Array("Software", "25", "32", "41"), _
              Array("Services", "12", "20", "31")), "STK"

    AddSample "Stacked Bar - Headcount by office", _
        GRID_NOTE & "Stacked bars fit long category names and many categories better than columns.", _
        Array(Array("", "Helsinki", "Stockholm", "Berlin"), _
              Array("Engineering", "38", "22", "31"), _
              Array("Sales", "12", "18", "9"), _
              Array("Operations", "9", "7", "11")), "SBR"

    AddSample "100% Column - Sales channel mix", _
        GRID_NOTE & "Use for MIX SHIFTS over time - here: online overtakes retail. Absolute totals are hidden on purpose.", _
        Array(Array("", "2022", "2024", "2026"), _
              Array("Online", "20", "38", "57"), _
              Array("Retail", "65", "48", "31"), _
              Array("Partner", "15", "14", "12")), "PCT"

    AddSample "Line - Customer satisfaction (NPS)", _
        GRID_NOTE & "Use for TRENDS over many periods - lines make the overtake in Q1 '26 obvious.", _
        Array(Array("", "Q1 25", "Q2 25", "Q3 25", "Q4 25", "Q1 26", "Q2 26"), _
              Array("Us", "42", "45", "49", "55", "62", "71"), _
              Array("Competitor", "58", "57", "55", "54", "52", "51")), "LINE"

    AddSample "Area - Installed base by product generation (k units)", _
        GRID_NOTE & "Use for CUMULATIVE volumes and generational replacement - Gen 3 grows while Gen 1 phases out, total keeps rising.", _
        Array(Array("", "2022", "2023", "2024", "2025", "2026"), _
              Array("Gen 1", "40", "32", "22", "12", "5"), _
              Array("Gen 2", "8", "25", "38", "42", "40"), _
              Array("Gen 3", "0", "3", "12", "28", "47")), "AREA"

    AddSample "Mekko - Market by region and segment (EUR m)", _
        GRID_NOTE & "Use to show TWO dimensions at once: column width = market size per region, segments = product mix within it.", _
        Array(Array("", "Europe", "Americas", "Asia"), _
              Array("Premium", "25", "20", "10"), _
              Array("Standard", "20", "35", "15"), _
              Array("Budget", "10", "25", "20")), "MEK"

    AddSample "Waterfall - EBIT bridge (EUR m)", _
        "Table: rows of label | value. 'e' or '=' creates a COMPUTED subtotal that always shows the " & _
        "running total - the classic P&L bridge.", _
        Array(Array("Revenue", "120"), Array("COGS", "-45"), _
              Array("Gross profit", "="), Array("Opex", "-32"), _
              Array("EBITDA", "="), Array("D&A", "-12"), _
              Array("EBIT", "=")), "WF"

    AddSample "Pie - Cost structure", _
        "Table: rows of label | value. Use with FEW slices and one clear message - here: half the cost is people.", _
        Array(Array("Personnel", "48"), Array("Facilities", "21"), _
              Array("Marketing", "17"), Array("Other", "14")), "PIE"

    AddSample "Doughnut - Revenue mix", _
        "Table: rows of label | value. Like a pie, with room in the middle for a headline number.", _
        Array(Array("Subscriptions", "55"), Array("Licenses", "25"), _
              Array("Services", "20")), "DON"

    AddSample "Scatter - Price vs. satisfaction", _
        "Table: rows of label | x | y. Use to show RELATIONSHIPS - the trend is clear and product Foxtrot " & _
        "is the overpriced outlier.", _
        Array(Array("Alpha", "35", "6.2"), Array("Bravo", "48", "7.1"), _
              Array("Charlie", "61", "7.9"), Array("Delta", "72", "8.4"), _
              Array("Echo", "90", "8.7"), Array("Foxtrot", "55", "5.1")), "SCAT"

    AddSample "Bubble - Portfolio: share vs. growth vs. revenue", _
        "Table: rows of label | x | y | size. The growth-share matrix: x = market share %, " & _
        "y = market growth %, bubble size = revenue.", _
        Array(Array("Alpha", "32", "4", "120"), Array("Bravo", "18", "12", "80"), _
              Array("Charlie", "9", "22", "40"), Array("Delta", "4", "28", "15"), _
              Array("Echo", "25", "-2", "95")), "SCAT"

    AddSample "Gantt - Product launch plan", _
        "Table: rows of activity | start | end (dates or numbers). Overlapping phases show " & _
        "dependencies; start = end makes a milestone (Launch).", _
        Array(Array("Discovery", "1.9.2026", "19.9.2026"), _
              Array("Design", "15.9.2026", "10.10.2026"), _
              Array("Build", "6.10.2026", "14.11.2026"), _
              Array("Testing", "9.11.2026", "28.11.2026"), _
              Array("Launch", "1.12.2026", "1.12.2026")), "GANTT"
End Sub

' Remove all inserted sample slides again (they are tagged).
Public Sub RemoveSampleSlides()
    Dim p As Presentation
    Set p = ActivePresentation
    Dim i As Long, removed As Long
    For i = p.Slides.Count To 1 Step -1
        If p.Slides(i).Tags("SLIDEAID") = "SAMPLE" Then
            p.Slides(i).Delete
            removed = removed + 1
        End If
    Next i
    MsgBox removed & " sample slide(s) removed.", vbInformation, "Chart Aid"
End Sub

Private Sub AddSample(ByVal title As String, ByVal note As String, _
                      dat As Variant, ByVal kind As String)
    Dim p As Presentation
    Set p = ActivePresentation
    Dim sl As Slide
    Set sl = p.Slides.Add(p.Slides.Count + 1, ppLayoutBlank)
    sl.Tags.Add "SLIDEAID", "SAMPLE"            ' -> RemoveSampleSlides
    ActiveWindow.View.GotoSlide sl.SlideIndex   ' builders draw on the current slide

    ' title + note
    Dim tb As Shape
    Set tb = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, 20, 14, SlideW() - 40, 26)
    With tb.TextFrame.TextRange
        .Text = "Chart Aid sample: " & title
        .Font.Size = 20
        .Font.Bold = msoTrue
    End With
    Set tb = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, 20, 44, SlideW() - 40, 30)
    With tb.TextFrame.TextRange
        .Text = note
        .Font.Size = 11
        .Font.Color.RGB = RGB(100, 100, 100)
    End With

    ' the data table
    Dim nR As Long, nC As Long
    nR = UBound(dat) + 1
    nC = UBound(dat(0)) + 1
    Dim tbl As Shape
    Set tbl = sl.Shapes.AddTable(nR, nC, 20, 100, nC * 62, nR * 20)
    Dim r As Long, c As Long, oCell As Object
    For r = 1 To nR
        For c = 1 To nC
            On Error Resume Next
            Set oCell = tbl.Table.Cell(r, c)
            With oCell.Shape.TextFrame.TextRange
                .Text = CStr(dat(r - 1)(c - 1))
                .Font.Size = 10
            End With
            On Error GoTo 0
        Next c
    Next r

    ' the chart, built by the real chart code
    Dim d As ChartData
    If ReadTable(tbl, d) Then
        BuildChartFromData kind, d, ShpRight(tbl) + 30, 100, PLOT_W, PLOT_H
    End If
End Sub

Public Sub ChartDataHelp()
    MsgBox "Chart Aid reads its data from a PowerPoint table you select:" & vbCr & vbCr & _
           "Column / Bar / Stacked / 100% / Mekko / Line / Area:" & vbCr & _
           "  row 1 = category names (top-left cell free)," & vbCr & _
           "  column 1 = series names, body = numbers." & vbCr & vbCr & _
           "Waterfall:  rows of  label | value" & vbCr & _
           "  (value 'e' or '=' creates a computed subtotal bar)." & vbCr & vbCr & _
           "Pie / Doughnut:  rows of  label | value" & vbCr & vbCr & _
           "Scatter / Bubble:  rows of  label | x | y  (| size)" & vbCr & vbCr & _
           "Gantt:  rows of  activity | start | end  (numbers or dates;" & vbCr & _
           "  start = end makes a milestone)." & vbCr & vbCr & _
           "Select the table, click a chart button. To update: select the" & vbCr & _
           "chart, click Edit Data, change values, select table + chart," & vbCr & _
           "click the chart button again.", vbInformation, "Chart Aid - data layouts"
End Sub
