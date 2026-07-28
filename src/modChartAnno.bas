Attribute VB_Name = "modChartAnno"
' =====================================================================
' Chart Aid - annotations & elements:
'   * difference arrow (absolute / percent) between two bars/points
'   * CAGR arrow between two bars
'   * value line at a given value, average line over selected bars
'   * Harvey balls, checkboxes (with state cycling)
' Bars/points created by Chart Aid carry their value in a tag, so
' differences and CAGR are calculated from actual data, not pixels.
' =====================================================================
Option Explicit

Private Const GREY As Long = 5855577          ' RGB(89,89,89)

' Read the data value a Chart Aid bar/point carries.
Private Function ShapeVal(s As Shape, ByRef ok As Boolean) As Double
    Dim tg As String
    tg = s.Tags(TAG_VAL)
    If Len(tg) > 0 Then
        ShapeVal = Val(tg)
        ok = True
    Else
        ok = False
    End If
End Function

Private Function TwoBars(ByRef s1 As Shape, ByRef s2 As Shape) As Boolean
    TwoBars = False
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Function
    Set s1 = sr(1)
    Set s2 = sr(sr.Count)
    If s1.Left > s2.Left Then          ' left-to-right
        Dim tmp As Shape
        Set tmp = s1: Set s1 = s2: Set s2 = tmp
    End If
    TwoBars = True
End Function

' The y coordinate of a bar's VALUE end: top for positive values,
' bottom for negative ones (whose top edge is the baseline).
Private Function ValueEdgeY(s As Shape) As Single
    Dim ok As Boolean, v As Double
    v = ShapeVal(s, ok)
    If ok And v < 0 Then
        ValueEdgeY = ShpBottom(s)
    Else
        ValueEdgeY = s.Top
    End If
End Function

' ---------------------------------------------------------------
' DIFFERENCE ARROW between two selected bars/segments/points.
' mode "ABS" = absolute difference, "PCT" = percent difference.
' Click into the chart group to select individual bars.
' ---------------------------------------------------------------
Public Sub DifferenceArrow(ByVal mode As String)
    Dim s1 As Shape, s2 As Shape
    If Not TwoBars(s1, s2) Then Exit Sub

    Dim ok1 As Boolean, ok2 As Boolean, v1 As Double, v2 As Double
    v1 = ShapeVal(s1, ok1)
    v2 = ShapeVal(s2, ok2)

    Dim y1 As Single, y2 As Single, x As Single
    y1 = ValueEdgeY(s1): y2 = ValueEdgeY(s2)
    x = (ShpRight(s1) + s2.Left) / 2

    ' helper lines from each bar top to the arrow
    Dim h1 As Shape, h2 As Shape, ar As Shape, lab As Shape
    Set h1 = MakeLine(ShpRight(s1), y1, x, y1, GREY, 0.5, True)
    Set h2 = MakeLine(x, y2, s2.Left, y2, GREY, 0.5, True)

    Set ar = MakeLine(x, y1, x, y2, GREY, 1.5)
    ar.Line.BeginArrowheadStyle = msoArrowheadTriangle
    ar.Line.EndArrowheadStyle = msoArrowheadTriangle

    Dim txt As String
    If ok1 And ok2 Then
        If mode = "PCT" Then
            If v1 <> 0 Then
                ' (v2-v1)/|v1| keeps the sign meaningful for negative bases
                txt = Format$((v2 - v1) / Abs(v1) * 100, "+0.0;-0.0") & "%"
            Else
                txt = "n/a"
            End If
        Else
            txt = Format$(v2 - v1, "+#,##0.#;-#,##0.#")
        End If
    Else
        txt = "?"
    End If
    Set lab = MakeLabel(txt, x + 4, (y1 + y2) / 2 - 7, 60, 10, ppAlignLeft)
    lab.TextFrame.TextRange.Font.Bold = msoTrue

    CurrentSlide().Shapes.Range(Array(h1.Name, h2.Name, ar.Name, lab.Name)).Group.Select
End Sub

' ---------------------------------------------------------------
' CAGR ARROW between two selected bars.
' ---------------------------------------------------------------
Public Sub CagrArrow()
    Dim s1 As Shape, s2 As Shape
    If Not TwoBars(s1, s2) Then Exit Sub

    Dim ok1 As Boolean, ok2 As Boolean, v1 As Double, v2 As Double
    v1 = ShapeVal(s1, ok1)
    v2 = ShapeVal(s2, ok2)
    If Not (ok1 And ok2) Or v1 <= 0 Or v2 <= 0 Then
        MsgBox "Select two bars created by Chart Aid (positive values).", _
               vbExclamation, "Chart Aid"
        Exit Sub
    End If

    Dim nPer As Long
    nPer = AskInt("Number of periods between the two bars:", "1")
    If nPer < 1 Then Exit Sub

    Dim cagr As Double
    cagr = (v2 / v1) ^ (1# / nPer) - 1

    Dim ar As Shape, lab As Shape
    Set ar = MakeLine(ShpCenterX(s1), s1.Top - 14, ShpCenterX(s2), s2.Top - 14, GREY, 1.5)
    ar.Line.EndArrowheadStyle = msoArrowheadTriangle

    Set lab = MakeLabel("CAGR " & Format$(cagr * 100, "+0.0;-0.0") & "%" & _
                        IIf(nPer > 1, " p.a.", ""), 0, 0, 90, 10)
    lab.TextFrame.TextRange.Font.Bold = msoTrue
    lab.Left = (ShpCenterX(s1) + ShpCenterX(s2)) / 2 - lab.Width / 2
    lab.Top = (s1.Top + s2.Top) / 2 - 32

    CurrentSlide().Shapes.Range(Array(ar.Name, lab.Name)).Group.Select
End Sub

' ---------------------------------------------------------------
' VALUE LINE: select a Chart Aid chart, enter a value; a dashed
' line is drawn across the chart at that value, label on the right.
' ---------------------------------------------------------------
Public Sub ValueLine()
    Dim g As Shape
    Set g = FindOldChart()
    If g Is Nothing Then
        MsgBox "Select a chart created by Chart Aid first.", vbExclamation, "Chart Aid"
        Exit Sub
    End If
    If Len(g.Tags(TAG_PPU)) = 0 Then
        MsgBox "Select a chart created by Chart Aid first.", vbExclamation, "Chart Aid"
        Exit Sub
    End If

    Dim s As String
    s = InputBox("Value for the line:", "Chart Aid", "0")
    If Len(Trim$(s)) = 0 Then Exit Sub
    Dim v As Double
    v = Val(Replace(s, ",", "."))

    ' SCAT/BUB store the plot bottom (padded y-min), not a value-0
    ' baseline, so a value line cannot be placed on them.
    Dim typ As String
    typ = g.Tags(TAG_TYPE)
    If InStr(",PIE,DON,GANTT,SCAT,BUB,", "," & typ & ",") > 0 Then
        MsgBox "Value lines are not applicable to this chart type.", _
               vbExclamation, "Chart Aid"
        Exit Sub
    End If

    Dim y0 As Double, ppu As Double, ln As Shape, lab As Shape
    y0 = Val(g.Tags(TAG_Y0))
    ppu = Val(g.Tags(TAG_PPU))

    If typ = "BAR" Or typ = "SBR" Then
        ' horizontal charts: the value axis runs along x
        Dim x As Single
        x = y0 + v * ppu + (g.Left - Val(g.Tags(TAG_LEFT)))
        Set ln = MakeLine(x, g.Top, x, ShpBottom(g), RGB(192, 80, 77), 1, True)
        Set lab = MakeLabel(FmtNum(v), x - 20, g.Top - 14, 40, 9, ppAlignCenter, RGB(192, 80, 77))
        lab.Left = x - lab.Width / 2
    Else
        Dim y As Single
        y = y0 - v * ppu + (g.Top - Val(g.Tags(TAG_TOP)))
        Set ln = MakeLine(g.Left, y, ShpRight(g), y, RGB(192, 80, 77), 1, True)
        Set lab = MakeLabel(FmtNum(v), ShpRight(g) + 3, y - 7, 50, 9, ppAlignLeft, RGB(192, 80, 77))
    End If
    CurrentSlide().Shapes.Range(Array(ln.Name, lab.Name)).Group.Select
End Sub

' ---------------------------------------------------------------
' AVERAGE LINE over the selected bars (uses their data values).
' ---------------------------------------------------------------
Public Sub AverageLine()
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim i As Long, sum As Double, n As Long, ok As Boolean, v As Double
    Dim xMin As Single, xMax As Single
    Dim g As Object                     ' ParentGroup: late-bound for Mac
    xMin = 1E+30: xMax = -1E+30
    For i = 1 To sr.Count
        v = ShapeVal(sr(i), ok)
        If ok Then
            sum = sum + v: n = n + 1
            If sr(i).Left < xMin Then xMin = sr(i).Left
            If ShpRight(sr(i)) > xMax Then xMax = ShpRight(sr(i))
        End If
    Next i
    If n = 0 Then
        MsgBox "Select bars created by Chart Aid (click into the chart group).", _
               vbExclamation, "Chart Aid"
        Exit Sub
    End If

    ' scale from the parent chart group (fully late-bound for Mac)
    Dim oShp As Object
    On Error Resume Next
    Set oShp = sr(1)
    Set g = oShp.ParentGroup
    On Error GoTo 0
    If g Is Nothing Then
        MsgBox "The bars must belong to a Chart Aid chart.", vbExclamation, "Chart Aid"
        Exit Sub
    End If
    If Len(g.Tags(TAG_PPU)) = 0 Then
        MsgBox "The bars must belong to a Chart Aid chart.", vbExclamation, "Chart Aid"
        Exit Sub
    End If

    Dim avg As Double, y As Single
    avg = sum / n
    y = Val(g.Tags(TAG_Y0)) - avg * Val(g.Tags(TAG_PPU)) + (g.Top - Val(g.Tags(TAG_TOP)))

    Dim ln As Shape, lab As Shape
    Set ln = MakeLine(xMin - 4, y, xMax + 4, y, GREY, 1, True)
    Set lab = MakeLabel("Ø " & FmtNum(avg), xMax + 7, y - 7, 60, 9, ppAlignLeft, GREY)
    CurrentSlide().Shapes.Range(Array(ln.Name, lab.Name)).Group.Select
End Sub

' ---------------------------------------------------------------
' RECOLOR SERIES: click one bar/segment/point inside a Chart Aid
' chart, enter a color; the entire series is recolored across the
' chart. Works on plain shapes too (recolors just the selection).
' ---------------------------------------------------------------
Public Sub RecolorSeries()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    ' Whole chart group selected -> explain instead of recoloring
    ' everything one color (the classic first-click mistake).
    If sr.Count = 1 Then
        If Len(sr(1).Tags(TAG_TYPE)) > 0 Then
            MsgBox "You selected the whole chart." & vbCr & vbCr & _
                   "Click the chart once, then click ONE bar, segment or " & _
                   "point of the series you want to recolor - then run " & _
                   "Recolor Series again.", vbExclamation, "Chart Aid"
            Exit Sub
        End If
    End If

    ' Native macOS color panel (falls back to text entry if the
    ' AppleScript helper isn't installed). Default = clicked shape's
    ' fill. If the fill can't be read (e.g. the GROUP is selected, not
    ' a bar inside it), use a visible blue - a black default makes the
    ' color wheel render fully black (brightness slider at 0), which
    ' looks broken.
    Dim ok As Boolean, newC As Long, dflt As Long
    dflt = RGB(79, 129, 189)
    On Error Resume Next
    If sr(1).Fill.Visible = msoTrue Then dflt = sr(1).Fill.ForeColor.RGB
    On Error GoTo 0
    newC = NativePickColor(dflt, ok)
    If Not ok Then Exit Sub

    Dim i As Long, shp As Shape, m As Shape, ser As String, done As Long
    Dim g As Object                     ' ParentGroup: late-bound for Mac
    For i = 1 To sr.Count
        Set shp = sr(i)
        ser = shp.Tags(TAG_SER)
        Set g = Nothing
        Dim oShp2 As Object
        On Error Resume Next
        Set oShp2 = shp
        Set g = oShp2.ParentGroup
        On Error GoTo 0
        If Len(ser) > 0 And Not g Is Nothing Then
            For Each m In g.GroupItems
                If m.Tags(TAG_SER) = ser Then
                    On Error Resume Next
                    m.Fill.ForeColor.RGB = newC
                    m.Line.ForeColor.RGB = newC   ' line-chart segments
                    On Error GoTo 0
                    done = done + 1
                End If
            Next m
            UpsertColorOverride g, ser, newC      ' survive rebuilds
        Else
            On Error Resume Next
            shp.Fill.ForeColor.RGB = newC
            On Error GoTo 0
            done = done + 1
        End If
    Next i
End Sub

' ---------------------------------------------------------------
' HARVEY BALL: 0/25/50/75/100 % filled circle. Tagged with its
' percentage so Cycle State can step it 0 -> 25 -> ... -> 100 -> 0.
' ---------------------------------------------------------------
Public Sub InsertHarveyBall()
    Dim s As String
    s = InputBox("Fill percentage (0-100):", "Chart Aid", GetPref("HarveyPct", "50"))
    If Len(Trim$(s)) = 0 Then Exit Sub
    Dim p As Double
    p = Val(Replace(s, ",", "."))
    If p < 0 Then p = 0
    If p > 100 Then p = 100
    SetPref "HarveyPct", CStr(p)

    Dim d As Single: d = 18
    BuildHarveyBall(p, SlideW() / 2 - d / 2, SlideH() / 2 - d / 2, d).Select
End Sub

' Draw a Harvey ball at the given rect; returns the (tagged) shape.
Private Function BuildHarveyBall(ByVal p As Double, ByVal l As Single, _
                                 ByVal t As Single, ByVal d As Single) As Shape
    Dim bg As Shape, fg As Shape, out As Shape
    Set bg = CurrentSlide().Shapes.AddShape(msoShapeOval, l, t, d, d)
    bg.Fill.ForeColor.RGB = RGB(255, 255, 255)
    bg.Line.ForeColor.RGB = GREY
    bg.Line.Weight = 1.25

    If p >= 99.5 Then
        Set fg = CurrentSlide().Shapes.AddShape(msoShapeOval, l, t, d, d)
        fg.Fill.ForeColor.RGB = GREY
        fg.Line.Visible = msoFalse
        Set out = CurrentSlide().Shapes.Range(Array(bg.Name, fg.Name)).Group
    ElseIf p > 0.5 Then
        Set fg = CurrentSlide().Shapes.AddShape(msoShapePie, l, t, d, d)
        On Error Resume Next
        fg.Adjustments(1) = -90
        fg.Adjustments(2) = -90 + 3.6 * p
        On Error GoTo 0
        fg.Fill.ForeColor.RGB = GREY
        fg.Line.Visible = msoFalse
        Set out = CurrentSlide().Shapes.Range(Array(bg.Name, fg.Name)).Group
    Else
        Set out = bg
    End If
    out.Tags.Add "SACH_HARVEY", CStr(p)
    Set BuildHarveyBall = out
End Function

' Next stop on the 0/25/50/75/100 cycle.
Private Function NextHarveyPct(ByVal p As Double) As Double
    If p >= 100 Then
        NextHarveyPct = 0
    ElseIf p >= 75 Then
        NextHarveyPct = 100
    ElseIf p >= 50 Then
        NextHarveyPct = 75
    ElseIf p >= 25 Then
        NextHarveyPct = 50
    Else
        NextHarveyPct = 25
    End If
End Function

' ---------------------------------------------------------------
' CHECKBOX: insert, or cycle the state of selected ones
' (checked -> crossed -> empty -> checked).
' ---------------------------------------------------------------
Public Sub InsertCheckbox()
    Dim s As Shape
    Set s = CurrentSlide().Shapes.AddShape(msoShapeRoundedRectangle, _
            SlideW() / 2 - 9, SlideH() / 2 - 9, 18, 18)
    On Error Resume Next
    s.Adjustments(1) = 0.15
    On Error GoTo 0
    s.Fill.ForeColor.RGB = RGB(255, 255, 255)
    s.Line.ForeColor.RGB = GREY
    s.Line.Weight = 1.25
    With s.TextFrame
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .TextRange.Text = ChrW(10003)             ' check mark
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Color.RGB = RGB(79, 129, 189)
    End With
    s.Tags.Add "SACH_CHECK", "1"
    s.Select
End Sub

' Cycles checkboxes (checked -> crossed -> empty) AND Harvey balls
' (0 -> 25 -> 50 -> 75 -> 100 -> 0) in the selection.
Public Sub CycleCheckbox()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    ' Collect first: cycling a Harvey ball rebuilds (deletes) shapes,
    ' which must not happen while indexing the live ShapeRange.
    Dim shapes As New Collection
    Dim i As Long
    For i = 1 To sr.Count
        shapes.Add sr(i)
    Next i

    Dim s As Shape, cur As String
    Dim names() As String, nn As Long
    ReDim names(1 To shapes.Count)

    For Each s In shapes
        If Len(s.Tags("SACH_HARVEY")) > 0 Then
            Dim p As Double, l As Single, t As Single, d As Single
            p = NextHarveyPct(Val(s.Tags("SACH_HARVEY")))
            l = s.Left: t = s.Top: d = s.Width
            s.Delete
            nn = nn + 1
            names(nn) = BuildHarveyBall(p, l, t, d).Name
        ElseIf s.HasTextFrame Then
            cur = s.TextFrame.TextRange.Text
            With s.TextFrame.TextRange
                If cur = ChrW(10003) Then
                    .Text = ChrW(10007)           ' cross
                    .Font.Color.RGB = RGB(192, 80, 77)
                ElseIf cur = ChrW(10007) Then
                    .Text = ""
                Else
                    .Text = ChrW(10003)
                    .Font.Color.RGB = RGB(79, 129, 189)
                End If
                .Font.Size = 12
                .Font.Bold = msoTrue
            End With
            nn = nn + 1
            names(nn) = s.Name
        End If
    Next s

    ' Reselect so repeated clicks keep cycling.
    If nn > 0 Then
        ReDim Preserve names(1 To nn)
        On Error Resume Next
        CurrentSlide().Shapes.Range(names).Select
        On Error GoTo 0
    End If
End Sub
