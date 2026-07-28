Attribute VB_Name = "modPainter"
' =====================================================================
' Slide Aid - Advanced Format Painter
' Copy the FULL format of the Master (last selected) to all other
' selected shapes: fill, line, text font, margins, wrap, autosize.
' Position and size are NOT copied (use the Size tools for that).
' =====================================================================
Option Explicit

Public Sub AdvancedFormatPainter()
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim m As Shape
    Set m = GetMaster(sr)

    Dim i As Long
    For i = 1 To sr.Count - 1
        PaintShape m, sr(i)
    Next i
End Sub

Private Sub PaintShape(m As Shape, s As Shape)
    On Error Resume Next

    ' --- Fill ---
    If m.Fill.Visible = msoTrue Then
        s.Fill.Visible = msoTrue
        CopyColorFormat m.Fill.ForeColor, s.Fill.ForeColor
        s.Fill.Transparency = m.Fill.Transparency
    Else
        s.Fill.Visible = msoFalse
    End If

    ' --- Line ---
    If m.Line.Visible = msoTrue Then
        s.Line.Visible = msoTrue
        CopyColorFormat m.Line.ForeColor, s.Line.ForeColor
        s.Line.Weight = m.Line.Weight
        s.Line.DashStyle = m.Line.DashStyle
        s.Line.Transparency = m.Line.Transparency
    Else
        s.Line.Visible = msoFalse
    End If

    ' --- Shadow ---
    s.Shadow.Visible = m.Shadow.Visible
    If m.Shadow.Visible = msoTrue Then
        s.Shadow.Type = m.Shadow.Type
        s.Shadow.Blur = m.Shadow.Blur
        s.Shadow.OffsetX = m.Shadow.OffsetX
        s.Shadow.OffsetY = m.Shadow.OffsetY
        s.Shadow.Transparency = m.Shadow.Transparency
    End If

    ' --- Text ---
    If m.HasTextFrame And s.HasTextFrame Then
        Dim mf As Font, r As TextRange
        Set mf = m.TextFrame.TextRange.Font   ' first-run format of the Master
        For Each r In s.TextFrame.TextRange.Runs
            r.Font.Name = mf.Name
            r.Font.Size = mf.Size
            r.Font.Bold = mf.Bold
            r.Font.Italic = mf.Italic
            CopyColorFormat mf.Color, r.Font.Color
        Next r
        With s.TextFrame
            .MarginLeft = m.TextFrame.MarginLeft
            .MarginRight = m.TextFrame.MarginRight
            .MarginTop = m.TextFrame.MarginTop
            .MarginBottom = m.TextFrame.MarginBottom
            .WordWrap = m.TextFrame.WordWrap
            .AutoSize = m.TextFrame.AutoSize
            .VerticalAnchor = m.TextFrame.VerticalAnchor
        End With
        s.TextFrame.TextRange.ParagraphFormat.Alignment = _
            m.TextFrame.TextRange.ParagraphFormat.Alignment
    End If

    On Error GoTo 0
End Sub
