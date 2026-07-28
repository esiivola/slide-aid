Attribute VB_Name = "modAlignDock"
' =====================================================================
' Slide Aid - Align / Dock / Stretch / Fill Gap
' Master = last selected object. With a single object selected the
' slide is used as reference instead.
' =====================================================================
Option Explicit

' ---------------------------------------------------------------
' ALIGN: set the given edge of every object to the Master's edge.
' edge: "L","R","T","B","CH" (center horiz.), "CV" (center vert.)
' forceSlide:=True aligns to the slide even with multiple objects
' ---------------------------------------------------------------
Public Sub AlignToMaster(ByVal edge As String, Optional ByVal forceSlide As Boolean = False)
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim useSlide As Boolean
    useSlide = forceSlide Or (sr.Count = 1)

    Dim mL As Single, mT As Single, mR As Single, mB As Single
    Dim lastIdx As Long
    If useSlide Then
        mL = 0: mT = 0: mR = SlideW(): mB = SlideH()
        lastIdx = sr.Count + 1  ' align every shape
    Else
        Dim m As Shape
        Set m = GetMaster(sr)
        mL = m.Left: mT = m.Top: mR = ShpRight(m): mB = ShpBottom(m)
        lastIdx = sr.Count      ' skip the master itself
    End If

    Dim i As Long, s As Shape
    For i = 1 To sr.Count
        If i <> lastIdx Then
            Set s = sr(i)
            Select Case edge
                Case "L":  s.Left = mL
                Case "R":  s.Left = mR - s.Width
                Case "T":  s.Top = mT
                Case "B":  s.Top = mB - s.Height
                Case "CH": s.Left = (mL + mR) / 2 - s.Width / 2
                Case "CV": s.Top = (mT + mB) / 2 - s.Height / 2
            End Select
        End If
    Next i
End Sub

' ---------------------------------------------------------------
' DOCK: move objects in a direction until they touch the Master.
' Moving left => object ends up touching the Master's right edge, etc.
' Single object => moved to the corresponding slide edge.
' ---------------------------------------------------------------
Public Sub DockToMaster(ByVal direction As String)
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim i As Long, s As Shape

    If sr.Count = 1 Then
        Set s = sr(1)
        Select Case direction
            Case "L": s.Left = 0
            Case "R": s.Left = SlideW() - s.Width
            Case "T": s.Top = 0
            Case "B": s.Top = SlideH() - s.Height
        End Select
        Exit Sub
    End If

    Dim m As Shape
    Set m = GetMaster(sr)

    For i = 1 To sr.Count - 1
        Set s = sr(i)
        Select Case direction
            Case "L": s.Left = ShpRight(m)          ' moved left, stops at master's right edge
            Case "R": s.Left = m.Left - s.Width     ' moved right, stops at master's left edge
            Case "T": s.Top = ShpBottom(m)          ' moved up, stops at master's bottom edge
            Case "B": s.Top = m.Top - s.Height      ' moved down, stops at master's top edge
        End Select
    Next i
End Sub

' ---------------------------------------------------------------
' STRETCH: extend objects to the FAR edge of the Master
' (opposite edge stays fixed). Single object => slide edge.
' ---------------------------------------------------------------
Public Sub StretchToMaster(ByVal direction As String)
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim mL As Single, mT As Single, mR As Single, mB As Single
    Dim lastIdx As Long
    If sr.Count = 1 Then
        mL = 0: mT = 0: mR = SlideW(): mB = SlideH()
        lastIdx = 2
    Else
        Dim m As Shape
        Set m = GetMaster(sr)
        mL = m.Left: mT = m.Top: mR = ShpRight(m): mB = ShpBottom(m)
        lastIdx = sr.Count
    End If

    Dim i As Long, s As Shape, fixedEdge As Single
    For i = 1 To sr.Count
        If i <> lastIdx Then
            Set s = sr(i)
            Select Case direction
                Case "L"                       ' left edge -> master's left edge
                    fixedEdge = ShpRight(s)
                    If mL < fixedEdge Then
                        s.Left = mL
                        s.Width = fixedEdge - mL
                    End If
                Case "R"                       ' right edge -> master's right edge
                    If mR > s.Left Then s.Width = mR - s.Left
                Case "T"
                    fixedEdge = ShpBottom(s)
                    If mT < fixedEdge Then
                        s.Top = mT
                        s.Height = fixedEdge - mT
                    End If
                Case "B"
                    If mB > s.Top Then s.Height = mB - s.Top
            End Select
        End If
    Next i
End Sub

' ---------------------------------------------------------------
' PLACE ON SLIDE: position presets (halves, thirds,
' quadrants, full slide). A single object is resized to fill the
' region (small margin); several objects are moved as a block
' (top-left of their bounding box to the region's top-left),
' keeping their sizes and spacing.
' ---------------------------------------------------------------
Public Sub PlaceOnSlide(ByVal preset As String)
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim W As Single, H As Single
    W = SlideW(): H = SlideH()

    Dim l As Single, t As Single, rw As Single, rh As Single
    Select Case preset
        Case "LH": l = 0:         t = 0:     rw = W / 2: rh = H
        Case "RH": l = W / 2:     t = 0:     rw = W / 2: rh = H
        Case "TH": l = 0:         t = 0:     rw = W:     rh = H / 2
        Case "BH": l = 0:         t = H / 2: rw = W:     rh = H / 2
        Case "L3": l = 0:         t = 0:     rw = W / 3: rh = H
        Case "C3": l = W / 3:     t = 0:     rw = W / 3: rh = H
        Case "R3": l = W * 2 / 3: t = 0:     rw = W / 3: rh = H
        Case "Q1": l = 0:         t = 0:     rw = W / 2: rh = H / 2
        Case "Q2": l = W / 2:     t = 0:     rw = W / 2: rh = H / 2
        Case "Q3": l = 0:         t = H / 2: rw = W / 2: rh = H / 2
        Case "Q4": l = W / 2:     t = H / 2: rw = W / 2: rh = H / 2
        Case "FULL": l = 0:       t = 0:     rw = W:     rh = H
        Case Else: Exit Sub
    End Select

    Const M As Single = 12                 ' ~0.4 cm breathing room
    l = l + M: t = t + M: rw = rw - 2 * M: rh = rh - 2 * M

    If sr.Count = 1 Then
        With sr(1)
            .LockAspectRatio = msoFalse
            .Left = l: .Top = t
            .Width = rw: .Height = rh
        End With
    Else
        Dim i As Long, minL As Single, minT As Single
        minL = 1E+30: minT = 1E+30
        For i = 1 To sr.Count
            If sr(i).Left < minL Then minL = sr(i).Left
            If sr(i).Top < minT Then minT = sr(i).Top
        Next i
        For i = 1 To sr.Count
            sr(i).Left = sr(i).Left + (l - minL)
            sr(i).Top = sr(i).Top + (t - minT)
        Next i
    End If
End Sub

' ---------------------------------------------------------------
' FILL GAP: extend the edge of each object facing the Master until
' it touches the Master's near edge. Direction of extension given.
' Single object => extend to slide edge.
' ---------------------------------------------------------------
Public Sub FillGapToMaster(ByVal direction As String)
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim i As Long, s As Shape

    If sr.Count = 1 Then
        Set s = sr(1)
        Select Case direction
            Case "L": s.Width = ShpRight(s): s.Left = 0: ' keep right edge, extend to 0
            Case "R": s.Width = SlideW() - s.Left
            Case "T": s.Height = ShpBottom(s): s.Top = 0
            Case "B": s.Height = SlideH() - s.Top
        End Select
        Exit Sub
    End If

    Dim m As Shape
    Set m = GetMaster(sr)

    For i = 1 To sr.Count - 1
        Set s = sr(i)
        Select Case direction
            Case "L"   ' object right of master: extend left edge to master's right edge
                If s.Left > ShpRight(m) Then
                    s.Width = ShpRight(s) - ShpRight(m)
                    s.Left = ShpRight(m)
                End If
            Case "R"   ' object left of master: extend right edge to master's left edge
                If ShpRight(s) < m.Left Then s.Width = m.Left - s.Left
            Case "T"   ' object below master: extend top edge up to master's bottom edge
                If s.Top > ShpBottom(m) Then
                    s.Height = ShpBottom(s) - ShpBottom(m)
                    s.Top = ShpBottom(m)
                End If
            Case "B"   ' object above master: extend bottom edge down to master's top edge
                If ShpBottom(s) < m.Top Then s.Height = m.Top - s.Top
        End Select
    Next i
End Sub
