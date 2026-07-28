Attribute VB_Name = "modChartBars"
' =====================================================================
' Chart Aid - bar-family charts + router.
' Column, Bar, Stacked, Stacked Bar, 100% Column, Mekko, Waterfall.
' Chart semantics include waterfall 'e' subtotals,
' Mekko column widths proportional to column totals).
' =====================================================================
Option Explicit

' Entry point for all chart buttons (called by the ribbon dispatcher).
Public Sub BuildChart(ByVal kind As String)
    Dim tbl As Shape, oldC As Shape
    Set oldC = FindOldChart()
    Set tbl = FindTableShape()

    Dim d As ChartData
    If Not tbl Is Nothing Then
        If Not ReadTable(tbl, d) Then Set tbl = Nothing
    End If
    If tbl Is Nothing Then
        If Not oldC Is Nothing Then
            DeserializeData oldC.Tags(TAG_DATA), d   ' rebuild from tags
        Else
            MsgBox "Select a data table first (see Chart Aid > Help for the layout), " & _
                   "or select an existing chart to rebuild it.", vbExclamation, "Chart Aid"
            Exit Sub
        End If
    End If

    ' carry manual recolors across the rebuild (ChartRect deletes oldC)
    Dim ovr As String
    If Not oldC Is Nothing Then ovr = oldC.Tags(TAG_COLOVR)

    Dim l As Single, t As Single, w As Single, h As Single
    ChartRect tbl, oldC, l, t, w, h

    If BuildChartFromData(kind, d, l, t, w, h) Then
        ApplyOverridesToNewChart ovr
        ' auto-remove the datasheet if WE created it (Edit Data)
        If Not tbl Is Nothing Then
            If tbl.Tags("SACH_DATASHEET") = "1" Then tbl.Delete
        End If
    End If
End Sub

' REBUILD: rebuild the selected chart using its own stored type -
' no need to remember which chart button created it. Works with
' chart alone (stored data) or chart + edited data table.
Public Sub RebuildChart()
    Dim g As Shape
    Set g = FindOldChart()
    If g Is Nothing Then
        MsgBox "Select a Chart Aid chart (optionally together with its " & _
               "edited data table).", vbExclamation, "Chart Aid"
        Exit Sub
    End If
    BuildChart g.Tags(TAG_TYPE)
End Sub

' Every chart type a chart group's SACH_TYPE tag can carry. Checked
' before rebuild flows delete the old chart, so an unknown kind can
' never mean silent chart loss.
Public Function KnownChartKind(ByVal kind As String) As Boolean
    KnownChartKind = InStr(",COL,STK,PCT,BAR,SBR,MEK,WF,LINE,AREA,PIE,DON,SCAT,BUB,GANTT,", _
                           "," & kind & ",") > 0
End Function

' Build a chart directly from data (used by the ribbon flow above,
' the Sample Slides generator and the restyle flows). Returns True
' if the kind was recognized.
Public Function BuildChartFromData(ByVal kind As String, ByRef d As ChartData, _
                                   ByVal l As Single, ByVal t As Single, _
                                   ByVal w As Single, ByVal h As Single) As Boolean
    LoadChartPalette kind                        ' group palette wins
    LoadStyle
    SetStyleKind kind                            ' "<KIND>.Param" overrides
    BuildChartFromData = True
    Select Case kind
        Case "COL":   BuildBars d, l, t, w, h, False, False, False, kind
        Case "STK":   BuildBars d, l, t, w, h, True, False, False, kind
        Case "PCT":   BuildBars d, l, t, w, h, True, True, False, kind
        Case "BAR":   BuildBars d, l, t, w, h, False, False, True, kind
        Case "SBR":   BuildBars d, l, t, w, h, True, False, True, kind
        Case "MEK":   BuildMekko d, l, t, w, h
        Case "WF":    BuildWaterfall d, l, t, w, h
        Case "LINE":  BuildLine d, l, t, w, h, False
        Case "AREA":  BuildLine d, l, t, w, h, True
        Case "PIE":   BuildPie d, l, t, w, h, False
        Case "DON":   BuildPie d, l, t, w, h, True
        Case "SCAT", "BUB": BuildScatter d, l, t, w, h
        Case "GANTT": BuildGantt d, l, t, w, h
        Case Else:    BuildChartFromData = False
    End Select
    SetStyleKind ""                              ' don't leak per-type
                                                 ' overrides into later
                                                 ' StyleStr/FmtNum calls
End Function

' Dark or white label depending on fill luminance.
Public Function LabelColorOn(ByVal rgbFill As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = rgbFill And &HFF
    g = (rgbFill \ &H100) And &HFF
    b = (rgbFill \ &H10000) And &HFF
    If 0.299 * r + 0.587 * g + 0.114 * b > 150 Then
        LabelColorOn = RGB(50, 50, 50)
    Else
        LabelColorOn = RGB(255, 255, 255)
    End If
End Function

' ---------------------------------------------------------------
' Clustered / stacked / 100% columns and bars.
' horiz:=True rotates to bar charts (value axis horizontal).
' ---------------------------------------------------------------
Private Sub BuildBars(ByRef d As ChartData, ByVal l As Single, ByVal t As Single, _
                      ByVal w As Single, ByVal h As Single, _
                      ByVal stacked As Boolean, ByVal normalize As Boolean, _
                      ByVal horiz As Boolean, ByVal kind As String)
    Dim nSer As Long, nCat As Long
    nSer = d.rows: nCat = d.cols
    If nSer < 1 Or nCat < 1 Then Exit Sub

    Dim names() As String, nn As Long
    ReDim names(1 To 256)                        ' grown by AddName as needed
    Dim showVals As Boolean, showTots As Boolean, showLeg As Boolean
    showVals = (StyleNum("ValueLabels", 1) <> 0)
    showTots = (StyleNum("TotalLabels", 1) <> 0)
    showLeg = (StyleNum("Legend", 1) <> 0) And nSer > 1
    Dim labH As Single: labH = 16                ' room for category labels
    Dim legH As Single: If showLeg Then legH = 16 Else legH = 0
    Dim plotT As Single, plotH As Single, plotL As Single, plotW As Single
    plotT = t + legH: plotH = h - labH - legH
    plotL = l: plotW = w
    If horiz Then plotL = l + 60: plotW = w - 60  ' room for category labels left

    ' value range
    Dim r As Long, c As Long, v As Double
    Dim vMin As Double, vMax As Double, posSum As Double, negSum As Double
    For c = 1 To nCat
        posSum = 0: negSum = 0
        For r = 1 To nSer
            v = CellNum(d, r, c)
            If normalize Then v = Abs(v)
            If stacked Then
                If v >= 0 Then posSum = posSum + v Else negSum = negSum + v
            Else
                If v > vMax Then vMax = v
                If v < vMin Then vMin = v
            End If
        Next r
        If stacked Then
            If posSum > vMax Then vMax = posSum
            If negSum < vMin Then vMin = negSum
        End If
    Next c
    If normalize Then vMin = 0: vMax = 100

    Dim ppu As Double, y0 As Single
    If horiz Then
        CalcScale vMin, vMax, 0, plotW, ppu, y0   ' y0 = x of 0 rel. to plotL
        y0 = plotL + (0 - vMin) * ppu
    Else
        CalcScale vMin, vMax, plotT, plotH, ppu, y0
    End If

    ' legend
    If showLeg Then
        Dim lx As Single: lx = plotL
        For r = 1 To nSer
            AddName names, nn, MakeRect(lx, t + 3, 8, 8, ChartColorRGB(r))
            AddName names, nn, MakeLabel(d.cells(r, 0), lx + 11, t, 60, LSz(), ppAlignLeft)
            lx = lx + 12 + Len(d.cells(r, 0)) * 5.5 + 12
        Next r
    End If

    Dim slotSz As Single, barSz As Single
    If horiz Then slotSz = plotH / nCat Else slotSz = plotW / nCat

    Dim s As Shape, colTotal As Double
    For c = 1 To nCat
        Dim cumPos As Double, cumNeg As Double
        cumPos = 0: cumNeg = 0
        colTotal = 0
        If normalize Then
            For r = 1 To nSer: colTotal = colTotal + Abs(CellNum(d, r, c)): Next r
            If colTotal = 0 Then colTotal = 1
        End If

        For r = 1 To nSer
            v = CellNum(d, r, c)
            If normalize Then v = Abs(v) / colTotal * 100

            Dim bl As Single, bt As Single, bw As Single, bh As Single
            If stacked Then
                barSz = slotSz * StyleNum("StackFill", 0.65)
                If horiz Then
                    bt = plotT + (c - 1) * slotSz + (slotSz - barSz) / 2
                    bh = barSz
                    If v >= 0 Then
                        bl = y0 + cumPos * ppu: bw = v * ppu: cumPos = cumPos + v
                    Else
                        bl = y0 + (cumNeg + v) * ppu: bw = -v * ppu: cumNeg = cumNeg + v
                    End If
                Else
                    bl = plotL + (c - 1) * slotSz + (slotSz - barSz) / 2
                    bw = barSz
                    If v >= 0 Then
                        bt = y0 - (cumPos + v) * ppu: bh = v * ppu: cumPos = cumPos + v
                    Else
                        bt = y0 - cumNeg * ppu: bh = -v * ppu: cumNeg = cumNeg + v
                    End If
                End If
            Else
                barSz = slotSz * StyleNum("ClusterFill", 0.72) / nSer
                If horiz Then
                    bt = plotT + (c - 1) * slotSz + slotSz * 0.14 + (r - 1) * barSz
                    bh = barSz * 0.92
                    If v >= 0 Then
                        bl = y0: bw = v * ppu
                    Else
                        bl = y0 + v * ppu: bw = -v * ppu
                    End If
                Else
                    bl = plotL + (c - 1) * slotSz + slotSz * 0.14 + (r - 1) * barSz
                    bw = barSz * 0.92
                    If v >= 0 Then
                        bt = y0 - v * ppu: bh = v * ppu
                    Else
                        bt = y0: bh = -v * ppu
                    End If
                End If
            End If
            If bw < 0.5 Then bw = 0.5
            If bh < 0.5 Then bh = 0.5

            Set s = MakeRect(bl, bt, bw, bh, ChartColorRGB(r))
            s.Tags.Add TAG_VAL, Str$(CellNum(d, r, c))
            s.Tags.Add TAG_SER, CStr(r)
            AddName names, nn, s

            ' labels
            If Not showVals Then GoTo SkipLabels
            Dim lbl As String
            If normalize Then lbl = Format$(v, "0") & "%" Else lbl = FmtNum(v)
            If stacked Then
                If (Not horiz And bh > 11) Or (horiz And bw > 20) Then
                    Dim lab As Shape
                    Set lab = MakeLabel(lbl, bl, bt + bh / 2 - 7, bw, LSz(), ppAlignCenter, _
                                        LabelColorOn(ChartColorRGB(r)))
                    lab.Left = bl + (bw - lab.Width) / 2
                    AddName names, nn, lab
                End If
            Else
                Dim lab2 As Shape
                If horiz Then
                    Set lab2 = MakeLabel(lbl, bl + bw + 3, bt + bh / 2 - 7, 40, LSz(), ppAlignLeft)
                Else
                    Set lab2 = MakeLabel(lbl, bl - 10, IIf(v >= 0, bt - 13, bt + bh + 1), bw + 20, LSz())
                    lab2.Left = bl + (bw - lab2.Width) / 2
                End If
                AddName names, nn, lab2
            End If
SkipLabels:
        Next r

        ' stacked total label
        If stacked And Not normalize And showTots Then
            Dim tot As Double: tot = cumPos
            Dim tl As Shape
            If horiz Then
                Set tl = MakeLabel(FmtNum(tot), y0 + cumPos * ppu + 3, _
                       plotT + (c - 1) * slotSz + slotSz / 2 - 7, 40, 9, ppAlignLeft)
            Else
                Set tl = MakeLabel(FmtNum(tot), plotL + (c - 1) * slotSz, _
                       y0 - cumPos * ppu - 13, slotSz, 9)
                tl.Left = plotL + (c - 1) * slotSz + (slotSz - tl.Width) / 2
            End If
            AddName names, nn, tl
        End If

        ' category label
        Dim cl As Shape
        If horiz Then
            Set cl = MakeLabel(d.cells(0, c), l, plotT + (c - 1) * slotSz + slotSz / 2 - 7, 56, LSz(), ppAlignRight)
        Else
            Set cl = MakeLabel(d.cells(0, c), plotL + (c - 1) * slotSz, plotT + plotH + 2, slotSz, LSz())
            cl.Left = plotL + (c - 1) * slotSz + (slotSz - cl.Width) / 2
        End If
        AddName names, nn, cl
    Next c

    ' baseline
    If horiz Then
        AddName names, nn, MakeLine(y0, plotT, y0, plotT + plotH, RGB(89, 89, 89), 1)
    Else
        AddName names, nn, MakeLine(plotL, y0, plotL + plotW, y0, RGB(89, 89, 89), 1)
    End If

    GroupAndTag names, nn, kind, SerializeData(d), y0, ppu, l, t, w, h
End Sub

' ---------------------------------------------------------------
' MEKKO (percent axis): column widths proportional to column
' totals, segments normalized to 100% per column.
' ---------------------------------------------------------------
Private Sub BuildMekko(ByRef d As ChartData, ByVal l As Single, ByVal t As Single, _
                       ByVal w As Single, ByVal h As Single)
    Dim nSer As Long, nCat As Long
    nSer = d.rows: nCat = d.cols
    If nSer < 1 Or nCat < 1 Then Exit Sub

    Dim names() As String, nn As Long
    ReDim names(1 To 256)
    Dim labH As Single: labH = 30                ' category + total labels
    Dim plotT As Single, plotH As Single
    plotT = t + 16: plotH = h - labH - 16

    Dim colTot() As Double, grand As Double
    ReDim colTot(1 To nCat)
    Dim r As Long, c As Long
    For c = 1 To nCat
        For r = 1 To nSer
            colTot(c) = colTot(c) + Abs(CellNum(d, r, c))
        Next r
        grand = grand + colTot(c)
    Next c
    If grand = 0 Then Exit Sub

    Dim gap As Single: gap = StyleNum("MekkoGapPt", 2)
    Dim xw As Single, x As Single
    x = l
    Dim s As Shape
    For c = 1 To nCat
        xw = (w - gap * (nCat - 1)) * colTot(c) / grand
        Dim cum As Double: cum = 0
        For r = 1 To nSer
            Dim v As Double, segH As Single, segT As Single
            v = Abs(CellNum(d, r, c))
            If colTot(c) > 0 Then
                segH = plotH * v / colTot(c)
                segT = plotT + plotH - (cum + v) / colTot(c) * plotH
            End If
            If segH > 0.5 Then
                Set s = MakeRect(x, segT, xw, segH, ChartColorRGB(r))
                s.Tags.Add TAG_VAL, Str$(v)
                s.Tags.Add TAG_SER, CStr(r)
                AddName names, nn, s
                If segH > 11 And xw > 18 Then
                    Dim lab As Shape
                    Set lab = MakeLabel(FmtNum(v), x, segT + segH / 2 - 7, xw, 9, _
                                        ppAlignCenter, LabelColorOn(ChartColorRGB(r)))
                    lab.Left = x + (xw - lab.Width) / 2
                    AddName names, nn, lab
                End If
            End If
            cum = cum + v
        Next r
        ' column label + width (total) label
        Dim cl As Shape
        Set cl = MakeLabel(d.cells(0, c), x, plotT + plotH + 2, xw, LSz())
        cl.Left = x + (xw - cl.Width) / 2
        AddName names, nn, cl
        Set cl = MakeLabel(FmtNum(colTot(c)), x, plotT - 14, xw, LSz())
        cl.Left = x + (xw - cl.Width) / 2
        AddName names, nn, cl
        x = x + xw + gap
    Next c

    ' series legend (left of nothing better): swatch column on the right
    Dim ly As Single: ly = plotT
    For r = 1 To nSer
        AddName names, nn, MakeRect(l + w + 6, ly, 8, 8, ChartColorRGB(r))
        AddName names, nn, MakeLabel(d.cells(r, 0), l + w + 17, ly - 3, 60, LSz(), ppAlignLeft)
        ly = ly + 14
    Next r

    GroupAndTag names, nn, "MEK", SerializeData(d), plotT + plotH, plotH / 100, l, t, w, h
End Sub

' ---------------------------------------------------------------
' WATERFALL: rows of  label | value.  Value "e" or "=" creates a
' computed subtotal bar from the baseline.
' Positive = green, negative = red, subtotals = grey; connectors
' between consecutive bars.
' ---------------------------------------------------------------
Private Sub BuildWaterfall(ByRef d As ChartData, ByVal l As Single, ByVal t As Single, _
                           ByVal w As Single, ByVal h As Single)
    Dim n As Long
    n = UBound(d.cells, 1) + 1                   ' all rows incl. row 0
    If n < 2 Then Exit Sub

    Dim cUp As Long, cDown As Long, cTot As Long
    cUp = StyleColor("WaterfallUp", RGB(155, 187, 89))
    cDown = StyleColor("WaterfallDown", RGB(192, 80, 77))
    cTot = StyleColor("WaterfallTotal", RGB(191, 191, 191))

    ' trajectory for scaling
    Dim i As Long, cum As Double, vMin As Double, vMax As Double
    cum = 0
    For i = 0 To n - 1
        If Not CellIsSubtotal(d, i, 1) Then cum = cum + CellNum(d, i, 1)
        If cum > vMax Then vMax = cum
        If cum < vMin Then vMin = cum
    Next i

    Dim names() As String, nn As Long
    ReDim names(1 To 256)
    Dim labH As Single: labH = 16
    Dim plotT As Single, plotH As Single
    plotT = t + 14: plotH = h - labH - 14

    Dim ppu As Double, y0 As Single
    CalcScale vMin, vMax, plotT, plotH, ppu, y0

    Dim slotW As Single, barW As Single
    slotW = w / n: barW = slotW * StyleNum("WaterfallFill", 0.62)

    cum = 0
    Dim prevY As Single, prevRight As Single, havePrev As Boolean
    Dim s As Shape, lab As Shape
    For i = 0 To n - 1
        Dim bl As Single, bt As Single, bh As Single, v As Double
        Dim fillC As Long, labelV As Double, labelTop As Single
        bl = l + i * slotW + (slotW - barW) / 2

        If CellIsSubtotal(d, i, 1) Then
            v = cum
            fillC = cTot
            If v >= 0 Then
                bt = y0 - v * ppu: bh = v * ppu
            Else
                bt = y0: bh = -v * ppu
            End If
            labelV = v
        Else
            v = CellNum(d, i, 1)
            If v >= 0 Then fillC = cUp Else fillC = cDown
            If v >= 0 Then
                bt = y0 - (cum + v) * ppu: bh = v * ppu
            Else
                bt = y0 - cum * ppu: bh = -v * ppu
            End If
            cum = cum + v
            labelV = v
        End If
        If bh < 0.75 Then bh = 0.75

        Set s = MakeRect(bl, bt, barW, bh, fillC)
        s.Tags.Add TAG_VAL, Str$(labelV)
        AddName names, nn, s

        ' connector from previous bar at the previous cumulative level
        If havePrev Then
            AddName names, nn, MakeLine(prevRight, prevY, bl, prevY, RGB(150, 150, 150), 0.5)
        End If
        prevY = y0 - cum * ppu
        prevRight = bl + barW
        havePrev = True

        ' value label
        labelTop = bt - 13
        If labelV < 0 And Not CellIsSubtotal(d, i, 1) Then labelTop = bt + bh + 1
        Set lab = MakeLabel(FmtNum(labelV), bl - 10, labelTop, barW + 20, LSz())
        lab.Left = bl + (barW - lab.Width) / 2
        AddName names, nn, lab

        ' category label
        Set lab = MakeLabel(d.cells(i, 0), l + i * slotW, plotT + plotH + 2, slotW, LSz())
        lab.Left = l + i * slotW + (slotW - lab.Width) / 2
        AddName names, nn, lab
    Next i

    AddName names, nn, MakeLine(l, y0, l + w, y0, RGB(89, 89, 89), 1)

    GroupAndTag names, nn, "WF", SerializeData(d), y0, ppu, l, t, w, h
End Sub
