Attribute VB_Name = "modAgenda"
' =====================================================================
' Slide Aid - Agenda from Sections
' Organize your slides into PowerPoint SECTIONS (right-click in the
' thumbnail pane -> Add Section). Then run GenerateAgenda:
'  * an agenda overview slide is inserted at the start
'  * a separator slide is inserted before each section, with the
'    current item highlighted
' Re-running regenerates everything (generated slides are tagged).
' =====================================================================
Option Explicit

Private Const TAG_KEY As String = "SLIDEAID"
Private Const TAG_VAL As String = "AGENDA"

Public Sub GenerateAgenda()
    Dim p As Presentation
    Set p = ActivePresentation

    ' 1. Remove previously generated agenda slides
    Dim i As Long
    For i = p.Slides.Count To 1 Step -1
        If p.Slides(i).Tags(TAG_KEY) = TAG_VAL Then p.Slides(i).Delete
    Next i

    ' 2. Read sections
    Dim sp As Object
    On Error Resume Next
    Set sp = p.SectionProperties
    On Error GoTo 0
    If sp Is Nothing Then
        MsgBox "Sections are not available in this PowerPoint version.", vbExclamation, "Slide Aid"
        Exit Sub
    End If
    If sp.Count = 0 Then
        MsgBox "Organize your slides into sections first: right-click between " & _
               "slides in the thumbnail pane and choose 'Add Section'.", _
               vbInformation, "Slide Aid"
        Exit Sub
    End If

    Dim n As Long: n = sp.Count
    Dim names() As String, firsts() As Long
    ReDim names(1 To n): ReDim firsts(1 To n)
    For i = 1 To n
        names(i) = sp.Name(i)
        firsts(i) = sp.FirstSlide(i)
    Next i

    ' Final slide number of each section's first CONTENT slide, i.e.
    ' after the overview (+1) and the separators of sections 1..i are
    ' inserted. 0 = empty section (no number shown).
    Dim nums() As Long, k As Long
    ReDim nums(1 To n)
    For i = 1 To n
        If firsts(i) > 0 Then
            k = k + 1
            nums(i) = firsts(i) + k + 1
        End If
    Next i

    ' 3. Insert separators, LAST section first so indexes stay valid
    For i = n To 1 Step -1
        If firsts(i) > 0 Then   ' skip empty sections
            BuildAgendaSlide p, firsts(i), names, nums, i
        End If
    Next i

    ' 4. Overview slide at the very start (no highlight)
    ' (No success popup - the new slides are their own feedback.)
    BuildAgendaSlide p, 1, names, nums, 0
End Sub

Private Sub BuildAgendaSlide(p As Presentation, ByVal pos As Long, _
                             names() As String, nums() As Long, ByVal hi As Long)
    Dim sl As Slide
    Set sl = p.Slides.Add(pos, ppLayoutBlank)
    sl.Tags.Add TAG_KEY, TAG_VAL

    Dim W As Single, H As Single
    W = p.PageSetup.SlideWidth: H = p.PageSetup.SlideHeight

    Dim accent As Long, grey As Long
    accent = RGB(31, 73, 125): grey = RGB(128, 128, 128)
    On Error Resume Next
    accent = sl.ThemeColorScheme(msoThemeColorAccent1).RGB
    On Error GoTo 0

    ' Title
    Dim t As Shape
    Set t = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, W * 0.08, H * 0.08, W * 0.84, H * 0.12)
    With t.TextFrame.TextRange
        .Text = "Agenda"
        .Font.Size = 32
        .Font.Bold = msoTrue
        .Font.Color.RGB = accent
    End With

    ' Items
    Dim n As Long: n = UBound(names)
    Dim itemH As Single
    itemH = (H * 0.68) / n
    If itemH > H * 0.12 Then itemH = H * 0.12

    Dim i As Long, b As Shape
    For i = 1 To n
        Set b = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                W * 0.1, H * 0.24 + (i - 1) * itemH, W * 0.8, itemH * 0.85)
        With b.TextFrame.TextRange
            .Text = i & ".   " & names(i)
            .Font.Size = 20
            If i = hi Then
                .Font.Bold = msoTrue
                .Font.Color.RGB = accent
            Else
                .Font.Bold = msoFalse
                If hi = 0 Then
                    .Font.Color.RGB = RGB(64, 64, 64)  ' overview: all dark
                Else
                    .Font.Color.RGB = grey
                End If
            End If
        End With
        ' page number, right-aligned
        If nums(i) > 0 Then
            Dim pn As Shape
            Set pn = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                     W * 0.86, H * 0.24 + (i - 1) * itemH, W * 0.06, itemH * 0.85)
            With pn.TextFrame.TextRange
                .Text = CStr(nums(i))
                .Font.Size = 14
                .ParagraphFormat.Alignment = ppAlignRight
                If i = hi Then
                    .Font.Color.RGB = accent
                ElseIf hi = 0 Then
                    .Font.Color.RGB = RGB(64, 64, 64)
                Else
                    .Font.Color.RGB = grey
                End If
            End With
        End If
        ' highlight bar for the current item
        If i = hi Then
            Dim bar As Shape
            Set bar = sl.Shapes.AddShape(msoShapeRectangle, W * 0.07, _
                      H * 0.24 + (i - 1) * itemH, W * 0.012, itemH * 0.7)
            bar.Fill.ForeColor.RGB = accent
            bar.Line.Visible = msoFalse
        End If
    Next i
End Sub

' Remove all generated agenda slides.
Public Sub RemoveAgenda()
    Dim p As Presentation
    Set p = ActivePresentation
    Dim i As Long, removed As Long
    For i = p.Slides.Count To 1 Step -1
        If p.Slides(i).Tags(TAG_KEY) = TAG_VAL Then
            p.Slides(i).Delete
            removed = removed + 1
        End If
    Next i
    MsgBox removed & " agenda slide(s) removed.", vbInformation, "Slide Aid"
End Sub
