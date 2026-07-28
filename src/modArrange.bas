Attribute VB_Name = "modArrange"
' =====================================================================
' Slide Aid - Arrange: stack, matrix, spacing, swap, slice/multiply
' Stacking and matrix placement follow SELECTION ORDER.
' =====================================================================
Option Explicit

' ---------------------------------------------------------------
' STACK objects so they touch, in selection order.
' axis "H": side by side, tops aligned to first object.
' axis "V": on top of each other, lefts aligned to first object.
' Optional gap asked from user (0 = touching).
' ---------------------------------------------------------------
Public Sub StackObjects(ByVal axis As String, Optional ByVal askGap As Boolean = False)
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim gap As Single, ok As Boolean
    gap = 0
    If askGap Then
        gap = AskCm("Gap between objects (cm, negative = overlap):", "0", ok, "StackGap")
        If Not ok Then Exit Sub
    End If

    Dim i As Long, x As Single, y As Single
    x = sr(1).Left: y = sr(1).Top
    For i = 1 To sr.Count
        With sr(i)
            .Left = x
            .Top = y
        End With
        If axis = "H" Then
            x = x + sr(i).Width + gap
        Else
            y = y + sr(i).Height + gap
        End If
    Next i
End Sub

' ---------------------------------------------------------------
' MATRIX: arrange objects in a grid, in selection order, row by
' row. One-click version uses a near-square grid with objects
' touching; the "..." version asks columns and gaps.
' Cell pitch = widest / tallest object + gap.
' ---------------------------------------------------------------
Public Sub ArrangeMatrixQuick()
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub
    MatrixCore sr, Int(Sqr(sr.Count) + 0.999), 0, 0
End Sub

Public Sub ArrangeMatrix()
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim cols As Long
    cols = AskInt("Number of columns:", CStr(Int(Sqr(sr.Count) + 0.999))) ' default ~ square
    If cols < 1 Then Exit Sub

    Dim hGap As Single, vGap As Single, ok As Boolean
    hGap = AskCm("Horizontal gap (cm):", "0", ok, "MatrixHGap"): If Not ok Then Exit Sub
    vGap = AskCm("Vertical gap (cm):", "0", ok, "MatrixVGap"): If Not ok Then Exit Sub

    MatrixCore sr, cols, hGap, vGap
End Sub

Private Sub MatrixCore(sr As ShapeRange, ByVal cols As Long, _
                       ByVal hGap As Single, ByVal vGap As Single)
    ' Cell size from largest object
    Dim i As Long, cw As Single, ch As Single
    For i = 1 To sr.Count
        If sr(i).Width > cw Then cw = sr(i).Width
        If sr(i).Height > ch Then ch = sr(i).Height
    Next i

    Dim x0 As Single, y0 As Single, r As Long, c As Long
    x0 = sr(1).Left: y0 = sr(1).Top
    For i = 1 To sr.Count
        r = (i - 1) \ cols
        c = (i - 1) Mod cols
        sr(i).Left = x0 + c * (cw + hGap)
        sr(i).Top = y0 + r * (ch + vGap)
    Next i
End Sub

' ---------------------------------------------------------------
' SPACING: set precise gap between objects.
' axis "H": equal horizontal gap, left-to-right by position.
' axis "V": equal vertical gap, top-to-bottom by position.
' ---------------------------------------------------------------
Public Sub SetSpacing(ByVal axis As String)
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim gap As Single, ok As Boolean
    gap = AskCm("Exact spacing between objects (cm, negative = overlap):", "0.2", ok, _
                "Spacing" & axis)
    If Not ok Then Exit Sub

    Dim idx() As Long
    SortIndicesByPosition sr, axis, idx

    Dim i As Long, pos As Single
    If axis = "H" Then
        pos = sr(idx(1)).Left + sr(idx(1)).Width
        For i = 2 To sr.Count
            sr(idx(i)).Left = pos + gap
            pos = sr(idx(i)).Left + sr(idx(i)).Width
        Next i
    Else
        pos = sr(idx(1)).Top + sr(idx(1)).Height
        For i = 2 To sr.Count
            sr(idx(i)).Top = pos + gap
            pos = sr(idx(i)).Top + sr(idx(i)).Height
        Next i
    End If
End Sub

' ---------------------------------------------------------------
' SWAP: rotate the selected objects' places, in selection order
' (1 -> 2 -> ... -> n -> 1). For two objects this is a plain swap.
' You choose the reference point relative to
' which the objects are swapped, and the layer (z-order) position
' is swapped by default.
' mode:  "P"   = positions only
'        +"S"  = also swap sizes
'        +"Z"  = also swap layer position
' refPt: "C"  = centers (irrelevant for identical shapes)
'        "TL" / "TR" / "BL" / "BR" = anchor corner
' ---------------------------------------------------------------
Public Sub SwapObjects(ByVal mode As String, Optional ByVal refPt As String = "C")
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim n As Long: n = sr.Count
    Dim l() As Single, tp() As Single, r() As Single, b() As Single
    Dim w() As Single, h() As Single, z() As Long
    ReDim l(1 To n): ReDim tp(1 To n): ReDim r(1 To n): ReDim b(1 To n)
    ReDim w(1 To n): ReDim h(1 To n): ReDim z(1 To n)

    Dim i As Long
    For i = 1 To n
        l(i) = sr(i).Left: tp(i) = sr(i).Top
        r(i) = ShpRight(sr(i)): b(i) = ShpBottom(sr(i))
        w(i) = sr(i).Width: h(i) = sr(i).Height
        z(i) = sr(i).ZOrderPosition
    Next i

    ' Each object takes the slot of the NEXT one in selection order.
    Dim t As Long
    For i = 1 To n
        t = (i Mod n) + 1
        If InStr(mode, "S") > 0 Then
            sr(i).Width = w(t)
            sr(i).Height = h(t)
        End If
        Select Case refPt
            Case "TL"
                sr(i).Left = l(t):                sr(i).Top = tp(t)
            Case "TR"
                sr(i).Left = r(t) - sr(i).Width:  sr(i).Top = tp(t)
            Case "BL"
                sr(i).Left = l(t):                sr(i).Top = b(t) - sr(i).Height
            Case "BR"
                sr(i).Left = r(t) - sr(i).Width:  sr(i).Top = b(t) - sr(i).Height
            Case Else ' "C"
                sr(i).Left = (l(t) + r(t)) / 2 - sr(i).Width / 2
                sr(i).Top = (tp(t) + b(t)) / 2 - sr(i).Height / 2
        End Select
        If InStr(mode, "Z") > 0 Then SetZOrder sr(i), z(t)
    Next i
End Sub

' ---------------------------------------------------------------
' DISTRIBUTE: outermost objects keep their positions, the spaces
' between all objects in between are distributed evenly.
' axis "H" or "V".
' ---------------------------------------------------------------
Public Sub DistributeObjects(ByVal axis As String)
    Dim sr As ShapeRange
    Set sr = GetSelection(3)
    If sr Is Nothing Then Exit Sub

    Dim idx() As Long
    SortIndicesByPosition sr, axis, idx

    ' Envelope = min left/top .. max right/bottom over ALL shapes (the
    ' shape with the largest Left is not necessarily the one that
    ' reaches furthest right, e.g. when widths differ).
    Dim i As Long, sumSize As Single, span As Single, gap As Single, pos As Single
    Dim minPos As Single, maxEnd As Single
    minPos = 1E+30: maxEnd = -1E+30
    For i = 1 To sr.Count
        If axis = "H" Then
            sumSize = sumSize + sr(i).Width
            If sr(i).Left < minPos Then minPos = sr(i).Left
            If ShpRight(sr(i)) > maxEnd Then maxEnd = ShpRight(sr(i))
        Else
            sumSize = sumSize + sr(i).Height
            If sr(i).Top < minPos Then minPos = sr(i).Top
            If ShpBottom(sr(i)) > maxEnd Then maxEnd = ShpBottom(sr(i))
        End If
    Next i
    span = maxEnd - minPos
    gap = (span - sumSize) / (sr.Count - 1)

    If axis = "H" Then
        pos = minPos
        For i = 1 To sr.Count
            sr(idx(i)).Left = pos
            pos = pos + sr(idx(i)).Width + gap
        Next i
    Else
        pos = minPos
        For i = 1 To sr.Count
            sr(idx(i)).Top = pos
            pos = pos + sr(idx(i)).Height + gap
        Next i
    End If
End Sub

' ---------------------------------------------------------------
' GOLDEN CANON: vertically align objects inside the Master so the
' bottom margin is twice the top margin (most pleasing to the eye).
' The Master (last selected) should be higher than the objects.
' ---------------------------------------------------------------
Public Sub GoldenCanon()
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim m As Shape
    Set m = GetMaster(sr)

    Dim i As Long
    For i = 1 To sr.Count - 1
        sr(i).Top = m.Top + (m.Height - sr(i).Height) / 3
    Next i
End Sub

Private Sub SetZOrder(s As Shape, ByVal target As Long)
    Dim guard As Long
    Do While s.ZOrderPosition < target And guard < 500
        s.ZOrder msoBringForward
        guard = guard + 1
    Loop
    Do While s.ZOrderPosition > target And guard < 1000
        s.ZOrder msoSendBackward
        guard = guard + 1
    Loop
End Sub

' ---------------------------------------------------------------
' SLICE / MULTIPLY the selected shape into rows x columns.
' The original shape is resized to the first cell; copies fill
' the rest of the original footprint (slice) or extend beyond it
' (multiply keeps original size and duplicates it).
' ---------------------------------------------------------------
Public Sub SliceShape()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub
    If sr.Count > 1 Then
        MsgBox "Select a single shape to slice.", vbExclamation, "Slide Aid"
        Exit Sub
    End If

    Dim rows As Long, cols As Long, gap As Single, ok As Boolean
    rows = AskInt("Number of rows:", "2", "SliceRows"): If rows < 1 Then Exit Sub
    cols = AskInt("Number of columns:", "2", "SliceCols"): If cols < 1 Then Exit Sub
    gap = AskCm("Gap between pieces (cm):", "0.1", ok, "SliceGap")
    If Not ok Or gap < 0 Then Exit Sub

    Dim s As Shape
    Set s = sr(1)

    Dim x0 As Single, y0 As Single, cw As Single, ch As Single
    x0 = s.Left: y0 = s.Top
    cw = (s.Width - (cols - 1) * gap) / cols
    ch = (s.Height - (rows - 1) * gap) / rows
    If cw <= 0 Or ch <= 0 Then
        MsgBox "Gap too large for the shape size.", vbExclamation, "Slide Aid"
        Exit Sub
    End If

    s.Width = cw
    s.Height = ch

    Dim r As Long, c As Long, d As Shape
    For r = 0 To rows - 1
        For c = 0 To cols - 1
            If Not (r = 0 And c = 0) Then
                Set d = s.Duplicate(1)
                d.Left = x0 + c * (cw + gap)
                d.Top = y0 + r * (ch + gap)
            End If
        Next c
    Next r
    s.Left = x0
    s.Top = y0
End Sub

' MULTIPLY: duplicate the shape into a rows x cols grid at its
' current size (original stays as cell 1,1).
Public Sub MultiplyShape()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub
    If sr.Count > 1 Then
        MsgBox "Select a single shape to multiply.", vbExclamation, "Slide Aid"
        Exit Sub
    End If

    Dim rows As Long, cols As Long, gap As Single, ok As Boolean
    rows = AskInt("Number of rows:", "1", "MultRows"): If rows < 1 Then Exit Sub
    cols = AskInt("Number of columns:", "3", "MultCols"): If cols < 1 Then Exit Sub
    gap = AskCm("Gap between copies (cm, negative = overlap):", "0.2", ok, "MultGap")
    If Not ok Then Exit Sub

    Dim s As Shape
    Set s = sr(1)
    Dim x0 As Single, y0 As Single
    x0 = s.Left: y0 = s.Top

    Dim r As Long, c As Long, d As Shape
    For r = 0 To rows - 1
        For c = 0 To cols - 1
            If Not (r = 0 And c = 0) Then
                Set d = s.Duplicate(1)
                d.Left = x0 + c * (s.Width + gap)
                d.Top = y0 + r * (s.Height + gap)
            End If
        Next c
    Next r
End Sub
