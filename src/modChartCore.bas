Attribute VB_Name = "modChartCore"
' =====================================================================
' Chart Aid - core: data model, scaling, tagging, shared drawing.
'
' table-driven charts drawn from plain shapes. Data comes from a
' PowerPoint TABLE on the slide (the "datasheet"):
'   * Series charts: row 1 = category names (cell 1,1 free),
'     column 1 = series names, body = numbers.
'   * Waterfall / Pie / Doughnut: rows of  label | value
'     (waterfall: value "e" or "=" -> computed subtotal, tc-style).
'   * Scatter/Bubble: rows of  label | x | y [| size].
'   * Gantt: rows of  activity | start | end  (numbers or dates).
'
' Select the table, click a chart button. The chart is created next
' to the table, grouped, and tagged with its data so it can be
' re-edited (Edit Data) and rebuilt: select table + old chart and
' click a chart button to replace it in place.
' =====================================================================
Option Explicit

' Tags on the chart group
Public Const TAG_TYPE As String = "SACH_TYPE"
Public Const TAG_DATA As String = "SACH_DATA"
Public Const TAG_Y0 As String = "SACH_Y0"     ' y of value 0 (baseline)
Public Const TAG_PPU As String = "SACH_PPU"   ' points per unit
Public Const TAG_TOP As String = "SACH_TOP"   ' group top at build time
Public Const TAG_LEFT As String = "SACH_LEFT" ' group left at build time
                                              ' (to compensate for moves)
Public Const TAG_RECT As String = "SACH_RECT" ' ORIGINAL build rect "l|t|w|h".
Public Const TAG_GW As String = "SACH_GW"     ' group width at build time
Public Const TAG_GH As String = "SACH_GH"     ' group height at build time
' The group bbox is NOT the build rect (labels overhang the plot), so
' rebuilds must reuse TAG_RECT - reusing the bbox would shrink or grow
' the chart a little on every rebuild.
' Tags on individual bars/points
Public Const TAG_VAL As String = "SACH_VAL"   ' data value (arrows, averages)
Public Const TAG_SER As String = "SACH_SER"   ' series index (recoloring)

' Manual series recolors, kept across rebuilds: "ser=RRGGBB|ser=..."
Public Const TAG_COLOVR As String = "SACH_COLOVR"

Public Const ROW_SEP As String = "\n|"        ' serialization separators
Public Const CELL_SEP As String = "|;"

' Default plot size
Public Const PLOT_W As Single = 340           ' ~12 cm
Public Const PLOT_H As Single = 227           ' ~8 cm

' Custom chart palettes (loaded per build; see LoadChartPalette).
' A GROUP palette (per chart family: BARS / LINES / PIES) wins over
' the legacy global palette, which wins over the theme accents.
Private palRGB() As Long              ' global palette
Private palN As Long
Private grpRGB() As Long              ' group palette
Private grpN As Long
Private curGroup As String

' ---------- data holder ----------
' Grid(0,c) = category names, Grid(r,0) = series names (r>=1)
Public Type ChartData
    rows As Long                ' series count (grid) or data rows (list)
    cols As Long                ' category count (grid) or columns (list)
    cells() As String           ' (0..rows, 0..cols) raw text
End Type

' ---------- selection helpers ----------

Public Function FindTableShape() As Shape
    Dim sr As ShapeRange
    Set sr = GetSelection(1, False)
    If Not sr Is Nothing Then
        Dim i As Long
        For i = 1 To sr.Count
            If sr(i).HasTable Then
                Set FindTableShape = sr(i)
                Exit Function
            End If
        Next i
    End If
    Set FindTableShape = Nothing
End Function

' A previously built chart group in the selection (for replace-in-place)
Public Function FindOldChart() As Shape
    Dim sr As ShapeRange
    Set sr = GetSelection(1, False)
    If Not sr Is Nothing Then
        Dim i As Long
        For i = 1 To sr.Count
            If Len(sr(i).Tags(TAG_TYPE)) > 0 Then
                Set FindOldChart = sr(i)
                Exit Function
            End If
        Next i
    End If
    Set FindOldChart = Nothing
End Function

' ---------- table -> data ----------

Public Function ReadTable(tbl As Shape, ByRef d As ChartData) As Boolean
    ReadTable = False
    If tbl Is Nothing Then Exit Function
    Dim t As Table
    Set t = tbl.Table
    If t.rows.Count < 1 Or t.Columns.Count < 1 Then Exit Function

    d.rows = t.rows.Count - 1                 ' excluding header row
    d.cols = t.Columns.Count - 1              ' excluding series-name col
    ReDim d.cells(0 To t.rows.Count - 1, 0 To t.Columns.Count - 1)

    ' Cell.Shape is missing from the Mac object library - late-bind
    Dim r As Long, c As Long, oCell As Object
    For r = 1 To t.rows.Count
        For c = 1 To t.Columns.Count
            On Error Resume Next
            Set oCell = t.Cell(r, c)
            d.cells(r - 1, c - 1) = Trim$(oCell.Shape.TextFrame.TextRange.Text)
            On Error GoTo 0
        Next c
    Next r
    ReadTable = True
End Function

Public Function CellNum(ByRef d As ChartData, ByVal r As Long, ByVal c As Long) As Double
    Dim s As String
    s = Replace(d.cells(r, c), ",", ".")
    s = Replace(s, " ", "")
    s = Replace(s, "%", "")
    CellNum = Val(s)
End Function

Public Function CellIsSubtotal(ByRef d As ChartData, ByVal r As Long, ByVal c As Long) As Boolean
    Dim s As String
    s = LCase$(Trim$(d.cells(r, c)))
    CellIsSubtotal = (s = "e" Or s = "=")
End Function

' ---------- serialization (for Edit Data / rebuild) ----------

Public Function SerializeData(ByRef d As ChartData) As String
    Dim r As Long, c As Long, rowS As String, out As String
    For r = 0 To UBound(d.cells, 1)
        rowS = ""
        For c = 0 To UBound(d.cells, 2)
            If c > 0 Then rowS = rowS & CELL_SEP
            rowS = rowS & d.cells(r, c)
        Next c
        If r > 0 Then out = out & ROW_SEP
        out = out & rowS
    Next r
    SerializeData = out
End Function

Public Sub DeserializeData(ByVal s As String, ByRef d As ChartData)
    Dim rws() As String, cls() As String, r As Long, c As Long
    rws = Split(s, ROW_SEP)
    cls = Split(rws(0), CELL_SEP)
    ReDim d.cells(0 To UBound(rws), 0 To UBound(cls))
    For r = 0 To UBound(rws)
        cls = Split(rws(r), CELL_SEP)
        For c = 0 To UBound(cls)
            If c <= UBound(d.cells, 2) Then d.cells(r, c) = cls(c)
        Next c
    Next r
    d.rows = UBound(d.cells, 1)
    d.cols = UBound(d.cells, 2)
End Sub

' ---------- chart canvas ----------

' The rect a rebuild should use: the ORIGINAL build rect, compensated
' for any moving/resizing the user did to the group since. Falls back
' to the group's bounding box for charts built before TAG_RECT existed.
Public Sub RectFromOldChart(oldChart As Shape, _
                            ByRef l As Single, ByRef t As Single, _
                            ByRef w As Single, ByRef h As Single)
    l = oldChart.Left: t = oldChart.Top
    w = oldChart.Width: h = oldChart.Height

    Dim rectS As String
    rectS = oldChart.Tags(TAG_RECT)
    If Len(rectS) = 0 Then Exit Sub
    Dim p() As String
    p = Split(rectS, "|")
    If UBound(p) <> 3 Then Exit Sub

    Dim gw As Single, gh As Single, sx As Double, sy As Double
    gw = Val(oldChart.Tags(TAG_GW))
    gh = Val(oldChart.Tags(TAG_GH))
    sx = 1: sy = 1
    If gw > 0 Then sx = oldChart.Width / gw     ' user resized -> scale
    If gh > 0 Then sy = oldChart.Height / gh
    l = Val(p(0)) + (oldChart.Left - Val(oldChart.Tags(TAG_LEFT)))
    t = Val(p(1)) + (oldChart.Top - Val(oldChart.Tags(TAG_TOP)))
    w = Val(p(2)) * sx
    h = Val(p(3)) * sy
End Sub

' Where to draw: old chart's rect if replacing, else right of the table.
Public Sub ChartRect(tbl As Shape, oldChart As Shape, _
                     ByRef l As Single, ByRef t As Single, _
                     ByRef w As Single, ByRef h As Single)
    If Not oldChart Is Nothing Then
        RectFromOldChart oldChart, l, t, w, h
        oldChart.Delete
    Else
        LoadStyle
        w = StyleNum("PlotWidthCm", 12) * CM_TO_PT
        h = StyleNum("PlotHeightCm", 8) * CM_TO_PT
        If Not tbl Is Nothing Then
            l = ShpRight(tbl) + 20: t = tbl.Top
            If l + w > SlideW() Then l = SlideW() - w - 10
        Else
            l = (SlideW() - w) / 2: t = (SlideH() - h) / 2
        End If
    End If
End Sub

' ---------- scaling ----------

' Map a value range to the plot. Returns points-per-unit and the y
' coordinate of value 0 (baseline), honoring negative values.
Public Sub CalcScale(ByVal vMin As Double, ByVal vMax As Double, _
                     ByVal plotTop As Single, ByVal plotH As Single, _
                     ByRef ppu As Double, ByRef y0 As Single)
    If vMax < 0 Then vMax = 0
    If vMin > 0 Then vMin = 0
    If vMax - vMin < 0.000001 Then vMax = vMin + 1
    ppu = plotH / (vMax - vMin)
    y0 = plotTop + vMax * ppu
End Sub

' ---------- drawing helpers ----------

Public Function AccentColorRGB(ByVal idx As Long) As Long
    On Error Resume Next
    AccentColorRGB = CurrentSlide().ThemeColorScheme( _
        msoThemeColorAccent1 + ((idx - 1) Mod 6)).RGB
    If Err.Number <> 0 Then AccentColorRGB = RGB(79, 129, 189)
    On Error GoTo 0
End Function

' ---------- user-editable chart palette ----------
' If <container>/SlideAid/chartcolors.txt exists, its colors are used
' for chart series instead of the theme accents. One color per line,
' as hex (1F497D) or R,G,B (31,73,125). Lines starting with # are
' comments. Loaded once per chart build.

Public Function ChartPalettePath() As String
    ChartPalettePath = StoreDir() & "/chartcolors.txt"
End Function

' Chart family for palette grouping.
Public Function PaletteGroupOf(ByVal kind As String) As String
    Select Case kind
        Case "LINE", "AREA", "SCAT", "BUB": PaletteGroupOf = "LINES"
        Case "PIE", "DON":                  PaletteGroupOf = "PIES"
        Case Else:                          PaletteGroupOf = "BARS"
    End Select
End Function

Public Function GroupPalettePath(ByVal grp As String) As String
    GroupPalettePath = StoreDir() & "/chartcolors_" & LCase$(grp) & ".txt"
End Function

Private Sub LoadPaletteFile(ByVal path As String, arr() As Long, ByRef n As Long)
    n = 0
    ReDim arr(1 To 50)
    If Dir(path) = "" Then Exit Sub
    Dim f As Integer, ln As String, v As Long, ok As Boolean
    f = FreeFile
    On Error GoTo Done                ' Open failed: nothing to close
    Open path For Input As #f
    On Error GoTo CloseIt             ' from here on the file IS open
    Do While Not EOF(f)
        Line Input #f, ln
        v = ParseColorText(ln, ok)
        If ok And n < 50 Then
            n = n + 1
            arr(n) = v
        End If
    Loop
CloseIt:
    Close #f
Done:
    On Error GoTo 0
End Sub

' Load palettes for a chart kind ("" = global only).
Public Sub LoadChartPalette(Optional ByVal kind As String = "")
    LoadPaletteFile ChartPalettePath(), palRGB, palN
    grpN = 0
    curGroup = ""
    If Len(kind) > 0 Then
        curGroup = PaletteGroupOf(kind)
        LoadPaletteFile GroupPalettePath(curGroup), grpRGB, grpN
    End If
End Sub

Public Function ChartColorRGB(ByVal idx As Long) As Long
    If grpN > 0 Then
        ChartColorRGB = grpRGB(((idx - 1) Mod grpN) + 1)
    ElseIf palN > 0 Then
        ChartColorRGB = palRGB(((idx - 1) Mod palN) + 1)
    Else
        ChartColorRGB = AccentColorRGB(idx)
    End If
End Function

' "1F497D", "#1F497D" or "31,73,125" -> RGB value
Public Function ParseColorText(ByVal s As String, ByRef ok As Boolean) As Long
    ok = False
    s = Trim$(s)
    If Left$(s, 1) = "#" Then s = Mid$(s, 2)
    If Len(s) = 0 Then Exit Function
    On Error GoTo Fail
    If InStr(s, ",") > 0 Then
        Dim p() As String
        p = Split(s, ",")
        If UBound(p) <> 2 Then Exit Function
        ParseColorText = RGB(CLng(Trim$(p(0))), CLng(Trim$(p(1))), CLng(Trim$(p(2))))
        ok = True
    ElseIf Len(s) = 6 Then
        ParseColorText = RGB(CLng("&H" & Mid$(s, 1, 2)), _
                             CLng("&H" & Mid$(s, 3, 2)), _
                             CLng("&H" & Mid$(s, 5, 2)))
        ok = True
    End If
Fail:
End Function

Public Function MakeRect(ByVal l As Single, ByVal t As Single, _
                         ByVal w As Single, ByVal h As Single, _
                         ByVal rgbFill As Long) As Shape
    Dim s As Shape
    Set s = CurrentSlide().Shapes.AddShape(msoShapeRectangle, l, t, w, h)
    s.Fill.ForeColor.RGB = rgbFill
    s.Line.Visible = msoFalse
    Set MakeRect = s
End Function

Public Function MakeLabel(ByVal txt As String, ByVal l As Single, _
                          ByVal t As Single, ByVal w As Single, _
                          Optional ByVal sizePt As Single = 10, _
                          Optional ByVal align As Long = ppAlignCenter, _
                          Optional ByVal rgbText As Long = -1) As Shape
    Dim s As Shape
    Set s = CurrentSlide().Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, w, 14)
    With s.TextFrame
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeShapeToFitText
        With .TextRange
            .Text = txt
            .Font.Size = sizePt
            .ParagraphFormat.Alignment = align
            If rgbText >= 0 Then .Font.Color.RGB = rgbText Else .Font.Color.RGB = RGB(64, 64, 64)
        End With
    End With
    Set MakeLabel = s
End Function

Public Function MakeLine(ByVal x1 As Single, ByVal y1 As Single, _
                         ByVal x2 As Single, ByVal y2 As Single, _
                         Optional ByVal rgbLine As Long = -1, _
                         Optional ByVal weightPt As Single = 0.75, _
                         Optional ByVal dashed As Boolean = False) As Shape
    Dim s As Shape
    Set s = CurrentSlide().Shapes.AddLine(x1, y1, x2, y2)
    If rgbLine >= 0 Then s.Line.ForeColor.RGB = rgbLine Else s.Line.ForeColor.RGB = RGB(128, 128, 128)
    s.Line.Weight = weightPt
    If dashed Then s.Line.DashStyle = msoLineDash
    Set MakeLine = s
End Function

Public Function FmtNum(ByVal v As Double) As String
    Select Case LCase$(StyleStr("Decimals", "auto"))
        Case "0": FmtNum = Format$(v, "#,##0")
        Case "1": FmtNum = Format$(v, "#,##0.0")
        Case "2": FmtNum = Format$(v, "#,##0.00")
        Case Else
            If Abs(v - Int(v)) < 0.000001 Then
                FmtNum = Format$(v, "#,##0")
            Else
                FmtNum = Format$(v, "#,##0.0")
            End If
    End Select
End Function

' Group a set of shape names, tag the group, return it.
' rl/rt/rw/rh = the build rect the chart was asked to fill (kept in
' TAG_RECT so rebuilds don't drift). Numeric tags use Str$ (always a
' period decimal) because Val can't read locale CStr output ("1,5").
Public Function GroupAndTag(names() As String, ByVal n As Long, _
                            ByVal chartType As String, ByVal dataS As String, _
                            ByVal y0 As Single, ByVal ppu As Double, _
                            ByVal rl As Single, ByVal rt As Single, _
                            ByVal rw As Single, ByVal rh As Single) As Shape
    Dim arr() As String, i As Long
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = names(i)
    Next i
    Dim g As Shape
    Set g = CurrentSlide().Shapes.Range(arr).Group
    g.Tags.Add TAG_TYPE, chartType
    g.Tags.Add TAG_DATA, dataS
    g.Tags.Add TAG_Y0, Str$(y0)
    g.Tags.Add TAG_PPU, Str$(ppu)
    g.Tags.Add TAG_TOP, Str$(g.Top)
    g.Tags.Add TAG_LEFT, Str$(g.Left)
    g.Tags.Add TAG_GW, Str$(g.Width)
    g.Tags.Add TAG_GH, Str$(g.Height)
    g.Tags.Add TAG_RECT, Trim$(Str$(rl)) & "|" & Trim$(Str$(rt)) & "|" & _
                         Trim$(Str$(rw)) & "|" & Trim$(Str$(rh))
    g.Select
    Set GroupAndTag = g
End Function

' ---------- manual recolor persistence ----------

Public Function RGBToHex6(ByVal v As Long) As String
    RGBToHex6 = Right$("0" & Hex(v And &HFF), 2) & _
                Right$("0" & Hex((v \ &H100) And &HFF), 2) & _
                Right$("0" & Hex((v \ &H10000) And &HFF), 2)
End Function

' Remember "series ser is manually colored rgbVal" on the chart group.
Public Sub UpsertColorOverride(g As Object, ByVal ser As String, ByVal rgbVal As Long)
    On Error Resume Next
    Dim ovr As String, entries() As String, i As Long
    Dim out As String, hit As Boolean, hexC As String
    hexC = RGBToHex6(rgbVal)
    ovr = g.Tags(TAG_COLOVR)
    If Len(ovr) > 0 Then
        entries = Split(ovr, "|")
        For i = 0 To UBound(entries)
            If Split(entries(i), "=")(0) = ser Then
                entries(i) = ser & "=" & hexC
                hit = True
            End If
            If Len(out) > 0 Then out = out & "|"
            out = out & entries(i)
        Next i
    End If
    If Not hit Then
        If Len(out) > 0 Then out = out & "|"
        out = out & ser & "=" & hexC
    End If
    g.Tags.Add TAG_COLOVR, out
    On Error GoTo 0
End Sub

' Re-apply stored manual recolors to a (rebuilt) chart group.
Public Sub ApplyColorOverrides(g As Shape, ByVal ovr As String)
    If Len(ovr) = 0 Then Exit Sub
    Dim entries() As String, p() As String, i As Long
    Dim m As Shape, c As Long, okc As Boolean
    entries = Split(ovr, "|")
    For i = 0 To UBound(entries)
        p = Split(entries(i), "=")
        If UBound(p) = 1 Then
            c = ParseColorText(p(1), okc)
            If okc Then
                For Each m In g.GroupItems
                    If m.Tags(TAG_SER) = p(0) Then
                        On Error Resume Next
                        m.Fill.ForeColor.RGB = c
                        m.Line.ForeColor.RGB = c
                        On Error GoTo 0
                    End If
                Next m
            End If
        End If
    Next i
    g.Tags.Add TAG_COLOVR, ovr
End Sub

' Convenience for the rebuild flows: the builders leave the new
' chart group selected, so apply the carried-over overrides to it.
Public Sub ApplyOverridesToNewChart(ByVal ovr As String)
    If Len(ovr) = 0 Then Exit Sub
    On Error Resume Next
    Dim g As Shape
    Set g = ActiveWindow.Selection.ShapeRange(1)
    If Not g Is Nothing Then
        If Len(g.Tags(TAG_TYPE)) > 0 Then ApplyColorOverrides g, ovr
    End If
    On Error GoTo 0
End Sub

' Shape-name collector. Grows the array as needed, so large data
' tables never die with "Subscript out of range" mid-build.
Public Sub AddName(names() As String, ByRef n As Long, s As Shape)
    If n >= UBound(names) Then ReDim Preserve names(1 To UBound(names) * 2)
    n = n + 1
    names(n) = s.Name
End Sub
