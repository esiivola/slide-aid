Attribute VB_Name = "modChartGantt"
' =====================================================================
' Chart Aid - Gantt chart (timeline) from a table of
'   activity | start | end     (numbers, or dates like 1.3.2026)
' start = end -> milestone (diamond). Layout:
' activities left, bars on a scaled timeline, axis labels on top.
' =====================================================================
Option Explicit

' "d.m.yyyy" handled explicitly so parsing does not depend on the
' system locale (IsDate/CDate would read "1.9.2026" as Jan 9 on a
' US-locale Mac, and reject "31.12.2026" outright).
Private Function DottedDate(ByVal s As String, ByRef ok As Boolean) As Double
    ok = False
    Dim p() As String
    p = Split(s, ".")
    If UBound(p) <> 2 Then Exit Function
    If Not (IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2))) Then Exit Function
    On Error GoTo Fail
    DottedDate = CDbl(DateSerial(CLng(p(2)), CLng(p(1)), CLng(p(0))))
    ok = True
Fail:
End Function

Private Function TimeVal(ByVal s As String) As Double
    s = Trim$(s)
    Dim ok As Boolean, v As Double
    v = DottedDate(s, ok)
    If ok Then
        TimeVal = v
    ElseIf IsDate(s) Then
        TimeVal = CDbl(CDate(s))
    Else
        TimeVal = Val(Replace(s, ",", "."))
    End If
End Function

Private Function IsDateData(ByRef d As ChartData) As Boolean
    Dim s As String, ok As Boolean
    s = Trim$(d.cells(0, 1))
    DottedDate s, ok
    IsDateData = ok Or IsDate(s)
End Function

Public Sub BuildGantt(ByRef d As ChartData, ByVal l As Single, ByVal t As Single, _
                      ByVal w As Single, ByVal h As Single)
    Dim n As Long
    n = UBound(d.cells, 1) + 1
    If n < 1 Or UBound(d.cells, 2) < 2 Then
        MsgBox "Gantt needs rows of: activity | start | end.", vbExclamation, "Chart Aid"
        Exit Sub
    End If

    Dim i As Long, tMin As Double, tMax As Double
    tMin = 1E+300: tMax = -1E+300
    For i = 0 To n - 1
        Dim a As Double, b As Double
        a = TimeVal(d.cells(i, 1)): b = TimeVal(d.cells(i, 2))
        If b < a Then b = a
        If a < tMin Then tMin = a
        If b > tMax Then tMax = b
    Next i
    If tMax - tMin < 0.000001 Then tMax = tMin + 1

    Dim names() As String, nn As Long
    ReDim names(1 To 256)
    Dim labW As Single: labW = 78                 ' activity labels
    Dim axH As Single: axH = 16
    Dim plotL As Single, plotT As Single, plotW As Single, plotH As Single
    plotL = l + labW: plotT = t + axH
    plotW = w - labW: plotH = h - axH

    Dim pptScale As Double
    pptScale = plotW / (tMax - tMin)

    Dim rowH As Single
    rowH = plotH / n
    If rowH > 26 Then rowH = 26

    Dim dateMode As Boolean
    dateMode = IsDateData(d)

    ' axis: top line + ~5 gridlines with labels
    AddName names, nn, MakeLine(plotL, plotT, plotL + plotW, plotT, RGB(89, 89, 89), 1)
    Dim k As Long, gx As Single, gv As Double
    For k = 0 To 4
        gv = tMin + (tMax - tMin) * k / 4
        gx = plotL + (gv - tMin) * pptScale
        AddName names, nn, MakeLine(gx, plotT, gx, plotT + rowH * n, RGB(217, 217, 217), 0.5)
        Dim lbl As String
        If dateMode Then lbl = Format$(CDate(gv), "d.m.yy") Else lbl = FmtNum(gv)
        Dim lab As Shape
        Set lab = MakeLabel(lbl, gx - 25, t, 50, 8)
        lab.Left = gx - lab.Width / 2
        AddName names, nn, lab
    Next k

    ' rows
    Dim s As Shape
    For i = 0 To n - 1
        Dim x1 As Single, x2 As Single, yy As Single
        x1 = plotL + (TimeVal(d.cells(i, 1)) - tMin) * pptScale
        x2 = plotL + (TimeVal(d.cells(i, 2)) - tMin) * pptScale
        yy = plotT + i * rowH

        ' activity label
        Set lab = MakeLabel(d.cells(i, 0), l, yy + rowH / 2 - 7, labW - 6, LSz(), ppAlignLeft)
        AddName names, nn, lab

        If x2 - x1 < 1 Then
            ' milestone
            Set s = CurrentSlide().Shapes.AddShape(msoShapeDiamond, _
                    x1 - rowH * 0.28, yy + rowH * 0.22, rowH * 0.56, rowH * 0.56)
            s.Fill.ForeColor.RGB = RGB(89, 89, 89)
            s.Line.Visible = msoFalse
        Else
            Set s = CurrentSlide().Shapes.AddShape(msoShapeRoundedRectangle, _
                    x1, yy + rowH * 0.2, x2 - x1, rowH * 0.6)
            On Error Resume Next
            s.Adjustments(1) = 0.5
            On Error GoTo 0
            s.Fill.ForeColor.RGB = IIf(LCase$(StyleStr("GanttBarColor", "theme")) = "theme", ChartColorRGB(1), StyleColor("GanttBarColor", ChartColorRGB(1)))
            s.Line.Visible = msoFalse
        End If
        s.Tags.Add TAG_VAL, Str$(TimeVal(d.cells(i, 2)) - TimeVal(d.cells(i, 1)))
        AddName names, nn, s
    Next i

    GroupAndTag names, nn, "GANTT", SerializeData(d), plotT, pptScale, l, t, w, h
End Sub
