Attribute VB_Name = "modChartLines"
' =====================================================================
' Chart Aid - line, area, pie, doughnut, scatter and bubble charts.
' =====================================================================
Option Explicit

Private Const PI As Double = 3.14159265358979

' ---------------------------------------------------------------
' LINE (area:=False) and stacked AREA (area:=True).
' Grid data: row 1 categories, column 1 series names.
' ---------------------------------------------------------------
Public Sub BuildLine(ByRef d As ChartData, ByVal l As Single, ByVal t As Single, _
                     ByVal w As Single, ByVal h As Single, ByVal area As Boolean)
    Dim nSer As Long, nCat As Long
    nSer = d.rows: nCat = d.cols
    If nSer < 1 Or nCat < 2 Then
        MsgBox "Line/area charts need at least 1 series row and 2 categories.", _
               vbExclamation, "Chart Aid"
        Exit Sub
    End If

    Dim names() As String, nn As Long
    ReDim names(1 To 256)
    Dim labH As Single: labH = 16
    Dim plotT As Single, plotH As Single, plotW As Single
    plotT = t + 14: plotH = h - labH - 14
    plotW = w - 55                                ' room for series labels right

    ' range
    Dim r As Long, c As Long, v As Double, vMin As Double, vMax As Double, cum As Double
    For c = 1 To nCat
        cum = 0
        For r = 1 To nSer
            v = CellNum(d, r, c)
            If area Then
                cum = cum + v
                If cum > vMax Then vMax = cum
                If cum < vMin Then vMin = cum
            Else
                If v > vMax Then vMax = v
                If v < vMin Then vMin = v
            End If
        Next r
    Next c

    Dim ppu As Double, y0 As Single
    CalcScale vMin, vMax, plotT, plotH, ppu, y0

    Dim slotW As Single
    slotW = plotW / nCat

    Dim px As Single, py As Single, qx As Single, qy As Single
    Dim s As Shape, lab As Shape

    If area Then
        ' stacked areas: freeform per series (bottom = cumulative below)
        Dim base() As Double, topv() As Double
        ReDim base(1 To nCat): ReDim topv(1 To nCat)
        For r = 1 To nSer
            For c = 1 To nCat
                topv(c) = base(c) + CellNum(d, r, c)
            Next c
            Dim fb As FreeformBuilder
            Set fb = CurrentSlide().Shapes.BuildFreeform(msoEditingAuto, _
                     l + slotW / 2, y0 - topv(1) * ppu)
            For c = 2 To nCat
                fb.AddNodes msoSegmentLine, msoEditingAuto, _
                    l + (c - 1) * slotW + slotW / 2, y0 - topv(c) * ppu
            Next c
            For c = nCat To 1 Step -1
                fb.AddNodes msoSegmentLine, msoEditingAuto, _
                    l + (c - 1) * slotW + slotW / 2, y0 - base(c) * ppu
            Next c
            fb.AddNodes msoSegmentLine, msoEditingAuto, l + slotW / 2, y0 - topv(1) * ppu
            Set s = fb.ConvertToShape
            s.Fill.ForeColor.RGB = ChartColorRGB(r)
            s.Fill.Transparency = 0.15
            s.Line.Visible = msoFalse
            s.Tags.Add TAG_SER, CStr(r)
            AddName names, nn, s
            ' series label at last point
            AddName names, nn, MakeLabel(d.cells(r, 0), l + plotW + 4, _
                y0 - (base(nCat) + topv(nCat)) / 2 * ppu - 7, 50, 9, ppAlignLeft, ChartColorRGB(r))
            For c = 1 To nCat
                base(c) = topv(c)
            Next c
        Next r
    Else
        For r = 1 To nSer
            For c = 1 To nCat - 1
                px = l + (c - 1) * slotW + slotW / 2
                py = y0 - CellNum(d, r, c) * ppu
                qx = l + c * slotW + slotW / 2
                qy = y0 - CellNum(d, r, c + 1) * ppu
                Set s = MakeLine(px, py, qx, qy, ChartColorRGB(r), 2)
                s.Tags.Add TAG_SER, CStr(r)
                AddName names, nn, s
            Next c
            For c = 1 To nCat
                px = l + (c - 1) * slotW + slotW / 2
                py = y0 - CellNum(d, r, c) * ppu
                Dim mk As Single: mk = StyleNum("MarkerSizePt", 5)
                Set s = CurrentSlide().Shapes.AddShape(msoShapeOval, px - mk / 2, py - mk / 2, mk, mk)
                s.Fill.ForeColor.RGB = ChartColorRGB(r)
                s.Line.Visible = msoFalse
                s.Tags.Add TAG_VAL, Str$(CellNum(d, r, c))
                s.Tags.Add TAG_SER, CStr(r)
                AddName names, nn, s
                If StyleNum("ValueLabels", 1) <> 0 Then
                    Set lab = MakeLabel(FmtNum(CellNum(d, r, c)), px - 20, py - 16, 40, LSz() - 1)
                    lab.Left = px - lab.Width / 2
                    AddName names, nn, lab
                End If
            Next c
            AddName names, nn, MakeLabel(d.cells(r, 0), l + plotW + 4, _
                y0 - CellNum(d, r, nCat) * ppu - 7, 50, 9, ppAlignLeft, ChartColorRGB(r))
        Next r
    End If

    ' category labels + baseline
    For c = 1 To nCat
        Set lab = MakeLabel(d.cells(0, c), l + (c - 1) * slotW, plotT + plotH + 2, slotW, LSz())
        lab.Left = l + (c - 1) * slotW + (slotW - lab.Width) / 2
        AddName names, nn, lab
    Next c
    AddName names, nn, MakeLine(l, y0, l + plotW, y0, RGB(89, 89, 89), 1)

    GroupAndTag names, nn, IIf(area, "AREA", "LINE"), SerializeData(d), y0, ppu, l, t, w, h
End Sub

' ---------------------------------------------------------------
' PIE / DOUGHNUT: rows of  label | value.
' ---------------------------------------------------------------
Public Sub BuildPie(ByRef d As ChartData, ByVal l As Single, ByVal t As Single, _
                    ByVal w As Single, ByVal h As Single, ByVal doughnut As Boolean)
    Dim n As Long
    n = UBound(d.cells, 1) + 1
    If n < 1 Then Exit Sub

    Dim total As Double, i As Long
    For i = 0 To n - 1
        total = total + Abs(CellNum(d, i, 1))
    Next i
    If total = 0 Then Exit Sub

    Dim names() As String, nn As Long
    ReDim names(1 To 256)
    Dim dia As Single
    dia = h * 0.8
    If dia > w * 0.6 Then dia = w * 0.6
    Dim cx As Single, cy As Single
    cx = l + dia / 2 + 10: cy = t + h / 2

    Dim a0 As Double, sweep As Double, s As Shape, lab As Shape
    a0 = -90
    For i = 0 To n - 1
        Dim v As Double
        v = Abs(CellNum(d, i, 1))
        sweep = 360 * v / total
        If sweep > 0.5 Then
            If doughnut Then
                Set s = CurrentSlide().Shapes.AddShape(msoShapeBlockArc, _
                        cx - dia / 2, cy - dia / 2, dia, dia)
                On Error Resume Next
                s.Adjustments(1) = a0
                s.Adjustments(2) = a0 + sweep
                s.Adjustments(3) = 0.28            ' hole size
                On Error GoTo 0
            Else
                Set s = CurrentSlide().Shapes.AddShape(msoShapePie, _
                        cx - dia / 2, cy - dia / 2, dia, dia)
                On Error Resume Next
                s.Adjustments(1) = a0
                s.Adjustments(2) = a0 + sweep
                On Error GoTo 0
            End If
            s.Fill.ForeColor.RGB = ChartColorRGB(i + 1)
            s.Line.ForeColor.RGB = RGB(255, 255, 255)
            s.Line.Weight = 1
            s.Tags.Add TAG_VAL, Str$(v)
            s.Tags.Add TAG_SER, CStr(i + 1)
            AddName names, nn, s

            ' label at mid-angle
            Dim midRad As Double, lr As Double
            midRad = (a0 + sweep / 2) * PI / 180
            If doughnut Or sweep < 30 Then lr = dia * 0.62 Else lr = dia * 0.34
            Set lab = MakeLabel(d.cells(i, 0) & " " & FmtNum(v), 0, 0, 80, LSz(), ppAlignLeft)
            lab.Left = cx + Cos(midRad) * lr - IIf(Cos(midRad) < -0.2, lab.Width, IIf(Abs(Cos(midRad)) <= 0.2, lab.Width / 2, 0))
            lab.Top = cy + Sin(midRad) * lr - 7
            If Not doughnut And sweep >= 30 Then
                lab.TextFrame.TextRange.Font.Color.RGB = LabelColorOn(ChartColorRGB(i + 1))
            End If
            AddName names, nn, lab
        End If
        a0 = a0 + sweep
    Next i

    GroupAndTag names, nn, IIf(doughnut, "DON", "PIE"), SerializeData(d), cy, 1, l, t, w, h
End Sub

' ---------------------------------------------------------------
' SCATTER / BUBBLE: rows of  label | x | y [| size].
' ---------------------------------------------------------------
Public Sub BuildScatter(ByRef d As ChartData, ByVal l As Single, ByVal t As Single, _
                        ByVal w As Single, ByVal h As Single)
    Dim n As Long
    n = UBound(d.cells, 1) + 1
    If n < 1 Or UBound(d.cells, 2) < 2 Then
        MsgBox "Scatter needs rows of: label | x | y (optional | size).", vbExclamation, "Chart Aid"
        Exit Sub
    End If
    Dim hasSize As Boolean
    hasSize = (UBound(d.cells, 2) >= 3)

    Dim i As Long
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double, szMax As Double
    xMin = 1E+300: yMin = 1E+300: xMax = -1E+300: yMax = -1E+300
    For i = 0 To n - 1
        Dim xv As Double, yv As Double
        xv = CellNum(d, i, 1): yv = CellNum(d, i, 2)
        If xv < xMin Then xMin = xv
        If xv > xMax Then xMax = xv
        If yv < yMin Then yMin = yv
        If yv > yMax Then yMax = yv
        If hasSize Then If CellNum(d, i, 3) > szMax Then szMax = CellNum(d, i, 3)
    Next i
    ' 8% padding
    Dim padX As Double, padY As Double
    padX = (xMax - xMin) * 0.08: If padX = 0 Then padX = 1
    padY = (yMax - yMin) * 0.08: If padY = 0 Then padY = 1
    xMin = xMin - padX: xMax = xMax + padX
    yMin = yMin - padY: yMax = yMax + padY

    Dim names() As String, nn As Long
    ReDim names(1 To 256)
    Dim plotL As Single, plotT As Single, plotW As Single, plotH As Single
    plotL = l + 30: plotT = t + 6
    plotW = w - 40: plotH = h - 30

    Dim ppx As Double, ppy As Double
    ppx = plotW / (xMax - xMin)
    ppy = plotH / (yMax - yMin)

    ' axes
    AddName names, nn, MakeLine(plotL, plotT, plotL, plotT + plotH, RGB(89, 89, 89), 1)
    AddName names, nn, MakeLine(plotL, plotT + plotH, plotL + plotW, plotT + plotH, RGB(89, 89, 89), 1)
    AddName names, nn, MakeLabel(FmtNum(yMax), l - 2, plotT - 5, 30, 8, ppAlignRight)
    AddName names, nn, MakeLabel(FmtNum(yMin), l - 2, plotT + plotH - 10, 30, 8, ppAlignRight)
    AddName names, nn, MakeLabel(FmtNum(xMin), plotL - 10, plotT + plotH + 3, 40, 8, ppAlignLeft)
    AddName names, nn, MakeLabel(FmtNum(xMax), plotL + plotW - 30, plotT + plotH + 3, 40, 8, ppAlignRight)

    Dim s As Shape, lab As Shape
    For i = 0 To n - 1
        Dim px As Single, py As Single, dia As Single
        px = plotL + (CellNum(d, i, 1) - xMin) * ppx
        py = plotT + plotH - (CellNum(d, i, 2) - yMin) * ppy
        If hasSize And szMax > 0 Then
            dia = 6 + 22 * Sqr(CellNum(d, i, 3) / szMax)
        Else
            dia = 8
        End If
        Set s = CurrentSlide().Shapes.AddShape(msoShapeOval, px - dia / 2, py - dia / 2, dia, dia)
        s.Fill.ForeColor.RGB = ChartColorRGB(i + 1)   ' wraps by palette length
        s.Fill.Transparency = IIf(hasSize, 0.25, 0)
        s.Line.Visible = msoFalse
        s.Tags.Add TAG_VAL, Str$(CellNum(d, i, 2))
        AddName names, nn, s
        Set lab = MakeLabel(d.cells(i, 0), px + dia / 2 + 2, py - 7, 70, 8, ppAlignLeft)
        AddName names, nn, lab
    Next i

    GroupAndTag names, nn, IIf(hasSize, "BUB", "SCAT"), SerializeData(d), plotT + plotH, ppy, l, t, w, h
End Sub
