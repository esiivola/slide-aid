Attribute VB_Name = "modTableSnap"
' =====================================================================
' Slide Aid - Snap objects to the table cell they are (roughly) over.
' Works on icons, flags, harvey balls, traffic lights, ... anything
' positioned over a table on the same slide.
' mode: "C"  = center in cell
'       "L"  = left-align in cell (vertically centered)
'       "R"  = right-align in cell (vertically centered)
' Margin (cm) is asked for L/R modes.
' =====================================================================
Option Explicit

Public Sub SnapToTableCells(ByVal mode As String)
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    ' Find the table: use a selected table if any, else the first
    ' table on the slide.
    Dim tblShape As Shape
    Set tblShape = FindTable(sr)
    If tblShape Is Nothing Then
        MsgBox "No table found on this slide.", vbExclamation, "Slide Aid"
        Exit Sub
    End If

    Dim margin As Single, ok As Boolean
    margin = 0
    If mode = "L" Or mode = "R" Then
        margin = AskCm("Margin from cell edge (cm):", "0.15", ok, "TblSnapMargin")
        If Not ok Then Exit Sub
    End If

    Dim tbl As Table
    Set tbl = tblShape.Table

    ' Precompute cell boundaries
    Dim rowTop() As Single, colLeft() As Single
    Dim r As Long, c As Long
    ReDim rowTop(1 To tbl.Rows.Count + 1)
    ReDim colLeft(1 To tbl.Columns.Count + 1)
    rowTop(1) = tblShape.Top
    For r = 1 To tbl.Rows.Count
        rowTop(r + 1) = rowTop(r) + tbl.Rows(r).Height
    Next r
    colLeft(1) = tblShape.Left
    For c = 1 To tbl.Columns.Count
        colLeft(c + 1) = colLeft(c) + tbl.Columns(c).Width
    Next c

    Dim i As Long, s As Shape, cx As Single, cy As Single
    For i = 1 To sr.Count
        Set s = sr(i)
        If Not s.HasTable Then
            cx = ShpCenterX(s): cy = ShpCenterY(s)
            ' Find the cell whose bounds contain the object's center
            Dim cr As Long, cc As Long
            cr = 0: cc = 0
            For r = 1 To tbl.Rows.Count
                If cy >= rowTop(r) And cy < rowTop(r + 1) Then cr = r: Exit For
            Next r
            For c = 1 To tbl.Columns.Count
                If cx >= colLeft(c) And cx < colLeft(c + 1) Then cc = c: Exit For
            Next c
            If cr > 0 And cc > 0 Then
                ' Vertical: always center in the cell
                s.Top = (rowTop(cr) + rowTop(cr + 1)) / 2 - s.Height / 2
                Select Case mode
                    Case "C": s.Left = (colLeft(cc) + colLeft(cc + 1)) / 2 - s.Width / 2
                    Case "L": s.Left = colLeft(cc) + margin
                    Case "R": s.Left = colLeft(cc + 1) - margin - s.Width
                End Select
            End If
        End If
    Next i
End Sub

Private Function FindTable(sr As ShapeRange) As Shape
    Dim i As Long
    For i = 1 To sr.Count
        If sr(i).HasTable Then
            Set FindTable = sr(i)
            Exit Function
        End If
    Next i
    Dim s As Shape
    For Each s In CurrentSlide().Shapes
        If s.HasTable Then
            Set FindTable = s
            Exit Function
        End If
    Next s
    Set FindTable = Nothing
End Function
