Attribute VB_Name = "modSizeAngle"
' =====================================================================
' Slide Aid - Match size to Master, align block-arrow angles
' Master = last selected object.
' =====================================================================
Option Explicit

' ---------------------------------------------------------------
' MATCH SIZE: dim = "W", "H" or "WH" set width and/or height to the
' Master's, independently. "WR"/"HR" match one dimension and scale the
' other by the same factor, so each object keeps its own aspect ratio -
' the variant to use on images, which "WH" would distort. Objects keep
' their center.
' ---------------------------------------------------------------
Public Sub MatchSizeToMaster(ByVal dim_ As String)
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim m As Shape
    Set m = GetMaster(sr)

    Dim i As Long, s As Shape, cx As Single, cy As Single, f As Single
    For i = 1 To sr.Count - 1
        Set s = sr(i)
        cx = ShpCenterX(s): cy = ShpCenterY(s)
        s.LockAspectRatio = msoFalse
        If dim_ = "WR" Then
            If s.Width <> 0 Then f = m.Width / s.Width Else f = 1
            s.Width = m.Width
            s.Height = s.Height * f
        ElseIf dim_ = "HR" Then
            If s.Height <> 0 Then f = m.Height / s.Height Else f = 1
            s.Height = m.Height
            s.Width = s.Width * f
        Else
            If InStr(dim_, "W") > 0 Then s.Width = m.Width
            If InStr(dim_, "H") > 0 Then s.Height = m.Height
        End If
        s.Left = cx - s.Width / 2
        s.Top = cy - s.Height / 2
    Next i
End Sub

' ---------------------------------------------------------------
' MAGIC RESIZER: scale the complete selected arrangement around its
' bounding-box center, including object positions and font sizes.
' ---------------------------------------------------------------
Public Sub MagicResizer()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim pct As Long
    pct = AskInt("Resize to (% of current size, e.g. 120):", "120", "MagicPct")
    If pct <= 0 Then Exit Sub
    Dim f As Single
    f = pct / 100

    Dim blockCx As Single, blockCy As Single
    blockCx = sr.Left + sr.Width / 2
    blockCy = sr.Top + sr.Height / 2

    Dim i As Long, s As Shape, cx As Single, cy As Single
    Dim newCx As Single, newCy As Single
    For i = 1 To sr.Count
        Set s = sr(i)
        cx = ShpCenterX(s): cy = ShpCenterY(s)
        newCx = blockCx + (cx - blockCx) * f
        newCy = blockCy + (cy - blockCy) * f
        s.LockAspectRatio = msoFalse
        s.Width = s.Width * f
        s.Height = s.Height * f
        s.Left = newCx - s.Width / 2
        s.Top = newCy - s.Height / 2
        On Error Resume Next
        If s.HasTextFrame Then
            Dim r As TextRange
            For Each r In s.TextFrame.TextRange.Runs
                r.Font.Size = r.Font.Size * f
            Next r
        End If
        On Error GoTo 0
    Next i
End Sub

' ---------------------------------------------------------------
' ALIGN PROCESS CHAIN: form a process chain from the selected
' block arrows. The Master (last selected) defines angle, vertical
' position and height; gaps are filled from left to right.
' ---------------------------------------------------------------
Public Sub AlignProcessChain()
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim m As Shape
    Set m = GetMaster(sr)

    ' Adopt the Master's height, top, rotation and adjustments
    Dim i As Long, j As Long, s As Shape
    For i = 1 To sr.Count - 1
        Set s = sr(i)
        s.LockAspectRatio = msoFalse
        s.Top = m.Top
        s.Height = m.Height
        s.Rotation = m.Rotation
        On Error Resume Next
        For j = 1 To m.Adjustments.Count
            s.Adjustments(j) = m.Adjustments(j)
        Next j
        On Error GoTo 0
    Next i

    ' Close the gaps from left to right
    Dim idx() As Long, x As Single
    SortIndicesByPosition sr, "H", idx
    x = sr(idx(1)).Left
    For i = 1 To sr.Count
        sr(idx(i)).Left = x
        x = x + sr(idx(i)).Width
    Next i
End Sub

' ---------------------------------------------------------------
' ALIGN BLOCK ARROWS: apply the Master's metrics (adjustment
' handles: arrowhead size, shaft thickness, ...) to all selected
' block arrows, regardless of exact arrow type.
' ---------------------------------------------------------------
Public Sub AlignBlockArrows()
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim m As Shape
    Set m = GetMaster(sr)

    Dim i As Long, j As Long, nAdj As Long
    For i = 1 To sr.Count - 1
        On Error Resume Next
        nAdj = sr(i).Adjustments.Count
        If nAdj > m.Adjustments.Count Then nAdj = m.Adjustments.Count
        For j = 1 To nAdj
            sr(i).Adjustments(j) = m.Adjustments(j)
        Next j
        On Error GoTo 0
    Next i
End Sub

' ---------------------------------------------------------------
' ALIGN ROUNDED RECTANGLES: give all selected rounded rectangles
' the same ABSOLUTE corner radius. Default = the Master's radius;
' enter a value to override. (PowerPoint stores the radius relative
' to the smaller side, so equal-looking corners need per-shape
' adjustment values - this fixes that.)
' ---------------------------------------------------------------
Public Sub AlignRoundedRectangles()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim m As Shape
    Set m = GetMaster(sr)

    Dim masterRadPt As Single, radPt As Single
    On Error Resume Next
    masterRadPt = m.Adjustments(1) * MinSide(m)
    On Error GoTo 0

    Dim ok As Boolean
    radPt = AskCm("Corner radius (cm):", _
                  Format(masterRadPt / CM_TO_PT, "0.00"), ok)
    If Not ok Or radPt < 0 Then Exit Sub

    Dim i As Long
    For i = 1 To sr.Count
        On Error Resume Next
        If sr(i).AutoShapeType = msoShapeRoundedRectangle Then
            sr(i).Adjustments(1) = radPt / MinSide(sr(i))
        End If
        On Error GoTo 0
    Next i
End Sub

Private Function MinSide(s As Shape) As Single
    If s.Width < s.Height Then MinSide = s.Width Else MinSide = s.Height
End Function

' ---------------------------------------------------------------
' ALIGN ANGLES: copy rotation of the Master to all selected
' shapes. If a shape is the same AutoShape type as the Master
' (e.g. same block arrow), its adjustment handles (arrowhead
' width, shaft thickness, ...) are copied too.
' ---------------------------------------------------------------
Public Sub AlignAnglesToMaster()
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim m As Shape
    Set m = GetMaster(sr)

    Dim i As Long, j As Long, s As Shape
    For i = 1 To sr.Count - 1
        Set s = sr(i)
        s.Rotation = m.Rotation
        On Error Resume Next
        If s.AutoShapeType = m.AutoShapeType And m.Adjustments.Count > 0 Then
            For j = 1 To m.Adjustments.Count
                s.Adjustments(j) = m.Adjustments(j)
            Next j
        End If
        On Error GoTo 0
    Next i
End Sub
