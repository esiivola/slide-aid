Attribute VB_Name = "modSelect"
' =====================================================================
' Slide Aid - Select Similar Shapes
' Select all shapes on the current slide similar to the Master
' (last selected shape).
' criteria: "T" = same AutoShape type
'           "F" = same fill color
'           "TF" = both
' =====================================================================
Option Explicit

Public Sub SelectSimilarShapes(ByVal criteria As String)
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim m As Shape
    Set m = GetMaster(sr)

    Dim mType As Long, mFill As Long
    On Error Resume Next
    mType = m.AutoShapeType
    mFill = m.Fill.ForeColor.RGB
    On Error GoTo 0

    Dim names() As String, n As Long
    ReDim names(1 To CurrentSlide().Shapes.Count)

    Dim s As Shape, ok As Boolean
    For Each s In CurrentSlide().Shapes
        ok = True
        On Error Resume Next
        If InStr(criteria, "T") > 0 Then
            If s.AutoShapeType <> mType Or s.Type <> m.Type Then ok = False
        End If
        If ok And InStr(criteria, "F") > 0 Then
            If s.Fill.Visible <> m.Fill.Visible Then
                ok = False
            ElseIf m.Fill.Visible = msoTrue And s.Fill.ForeColor.RGB <> mFill Then
                ok = False
            End If
        End If
        If Err.Number <> 0 Then ok = False: Err.Clear
        On Error GoTo 0
        If ok Then
            n = n + 1
            names(n) = s.Name
        End If
    Next s

    If n = 0 Then
        MsgBox "No similar shapes found.", vbInformation, "Slide Aid"
        Exit Sub
    End If

    ReDim Preserve names(1 To n)
    CurrentSlide().Shapes.Range(names).Select
End Sub
