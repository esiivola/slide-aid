Attribute VB_Name = "modView"
' =====================================================================
' Slide Aid - View helpers
' Temporarily hide objects that are in the way while editing
' crowded slides; unhide restores position and layer (visibility
' does not change either).
' =====================================================================
Option Explicit

Public Sub HideSelectedObjects()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To sr.Count
        sr(i).Visible = msoFalse
    Next i
End Sub

Public Sub UnhideHiddenObjects()
    Dim s As Shape
    For Each s In CurrentSlide().Shapes
        If s.Visible = msoFalse Then s.Visible = msoTrue
    Next s
End Sub
