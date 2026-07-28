Attribute VB_Name = "modCleanup"
' =====================================================================
' Slide Aid - Clean-up & expert tools
' =====================================================================
Option Explicit

' Remove all speaker notes (avoid leaking confidential remarks).
Public Sub RemoveAllNotes()
    If MsgBox("Remove speaker notes from ALL slides?", vbYesNo + vbQuestion, _
              "Slide Aid") <> vbYes Then Exit Sub
    Dim sl As Slide, s As Shape
    For Each sl In ActivePresentation.Slides
        On Error Resume Next
        For Each s In sl.NotesPage.Shapes
            If s.PlaceholderFormat.Type = ppPlaceholderBody Then
                If s.HasTextFrame Then s.TextFrame.TextRange.Text = ""
            End If
        Next s
        On Error GoTo 0
    Next sl
    MsgBox "Speaker notes removed.", vbInformation, "Slide Aid"
End Sub

' Remove all animations from all slides.
Public Sub RemoveAllAnimations()
    If MsgBox("Remove animations from ALL slides?", vbYesNo + vbQuestion, _
              "Slide Aid") <> vbYes Then Exit Sub
    Dim sl As Slide
    For Each sl In ActivePresentation.Slides
        On Error Resume Next
        Do While sl.TimeLine.MainSequence.Count > 0
            sl.TimeLine.MainSequence(1).Delete
        Loop
        On Error GoTo 0
    Next sl
    MsgBox "Animations removed.", vbInformation, "Slide Aid"
End Sub

' Delete slide designs (masters) not used by any slide - shrinks files.
Public Sub DeleteUnusedDesigns()
    Dim p As Presentation
    Set p = ActivePresentation
    Dim i As Long, sl As Slide, used As Boolean, removed As Long
    For i = p.Designs.Count To 1 Step -1
        If p.Designs.Count = 1 Then Exit For   ' always keep one
        used = False
        For Each sl In p.Slides
            If sl.Design Is p.Designs(i) Then used = True: Exit For
        Next sl
        If Not used Then
            On Error Resume Next
            p.Designs(i).Delete
            If Err.Number = 0 Then removed = removed + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next i
    MsgBox removed & " unused design(s) removed.", vbInformation, "Slide Aid"
End Sub

' Copy the titles of the selected (or all) slides to the clipboard.
Public Sub CopySummaryToClipboard()
    Dim txt As String, sl As Slide
    Dim rng As SlideRange
    On Error Resume Next
    Set rng = ActiveWindow.Selection.SlideRange
    On Error GoTo 0

    If rng Is Nothing Then
        For Each sl In ActivePresentation.Slides
            txt = txt & SlideTitle(sl) & vbCr
        Next sl
    Else
        For Each sl In rng
            txt = txt & SlideTitle(sl) & vbCr
        Next sl
    End If

    ' Mac VBA has no DataObject: use a temp text box + TextRange.Copy
    Dim tmp As Shape
    Set tmp = CurrentSlide().Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 400, 100)
    tmp.TextFrame.TextRange.Text = txt
    tmp.TextFrame.TextRange.Copy
    tmp.Delete
    MsgBox "Slide summary copied to clipboard.", vbInformation, "Slide Aid"
End Sub

Private Function SlideTitle(sl As Slide) As String
    On Error Resume Next
    If sl.Shapes.HasTitle Then
        SlideTitle = sl.SlideIndex & vbTab & sl.Shapes.Title.TextFrame.TextRange.Text
    End If
    If Len(SlideTitle) = 0 Then SlideTitle = sl.SlideIndex & vbTab & "(no title)"
    On Error GoTo 0
End Function

' Paste the clipboard on every selected slide.
Public Sub PasteOnSelectedSlides()
    Dim rng As SlideRange
    On Error GoTo NoSel
    Set rng = ActiveWindow.Selection.SlideRange
    Dim sl As Slide
    For Each sl In rng
        sl.Shapes.Paste
    Next sl
    Exit Sub
NoSel:
    MsgBox "Select slides in the thumbnail pane first (the clipboard must contain a copied object).", _
           vbExclamation, "Slide Aid"
End Sub

' Copy the selected slides into a new presentation (send/extract
' slides'). The user saves it via File > Save As (VBA SaveAs is
' unreliable on Mac).
Public Sub ExtractSelectedSlides()
    Dim rng As SlideRange
    On Error GoTo NoSel
    Set rng = ActiveWindow.Selection.SlideRange
    On Error GoTo 0

    Dim src As Presentation
    Set src = ActivePresentation
    rng.Copy

    Dim p As Presentation
    Set p = Presentations.Add(WithWindow:=msoTrue)
    p.PageSetup.SlideWidth = src.PageSetup.SlideWidth
    p.PageSetup.SlideHeight = src.PageSetup.SlideHeight
    p.Slides.Paste

    MsgBox rng.Count & " slide(s) copied to a new presentation. " & _
           "Save it via File > Save As.", vbInformation, "Slide Aid"
    Exit Sub
NoSel:
    MsgBox "Select slides in the thumbnail pane first.", vbExclamation, "Slide Aid"
End Sub

' Toggle visibility of slide-master background objects on this slide.
Public Sub ToggleMasterShapes()
    With CurrentSlide()
        .DisplayMasterShapes = Not .DisplayMasterShapes
    End With
End Sub

' Set proofing language. code: FI, ENUS, ENUK, SV, DE
' scope: "SEL" = selected shapes, "ALL" = whole presentation
Public Sub SetSpellLanguage(ByVal code As String, ByVal scope As String)
    Dim langID As Long
    Select Case code
        Case "FI":   langID = msoLanguageIDFinnish
        Case "ENUS": langID = msoLanguageIDEnglishUS
        Case "ENUK": langID = msoLanguageIDEnglishUK
        Case "SV":   langID = msoLanguageIDSwedish
        Case "DE":   langID = msoLanguageIDGerman
        Case Else:   Exit Sub
    End Select

    Dim s As Shape
    If scope = "SEL" Then
        Dim sr As ShapeRange
        Set sr = GetSelection(1)
        If sr Is Nothing Then Exit Sub
        For Each s In sr
            ApplyLang s, langID
        Next s
    Else
        Dim sl As Slide
        For Each sl In ActivePresentation.Slides
            For Each s In sl.Shapes
                ApplyLang s, langID
            Next s
        Next sl
        ' DefaultLanguageID is not in the Mac object library - late-bind
        On Error Resume Next
        Dim oPres As Object
        Set oPres = ActivePresentation
        oPres.DefaultLanguageID = langID
        On Error GoTo 0
    End If
End Sub

' LanguageID / Cell.Shape are missing from the Mac object library -
' everything text-related here is late-bound on purpose.
Private Sub ApplyLang(s As Shape, ByVal langID As Long)
    On Error Resume Next
    Dim oTR As Object
    If s.Type = msoGroup Then
        Dim g As Shape
        For Each g In s.GroupItems
            ApplyLang g, langID
        Next g
    ElseIf s.HasTable Then
        Dim r As Long, c As Long, oCell As Object
        For r = 1 To s.Table.Rows.Count
            For c = 1 To s.Table.Columns.Count
                Set oCell = s.Table.Cell(r, c)
                Set oTR = oCell.Shape.TextFrame.TextRange
                oTR.LanguageID = langID
            Next c
        Next r
    ElseIf s.HasTextFrame Then
        Set oTR = s.TextFrame.TextRange
        oTR.LanguageID = langID
    End If
    On Error GoTo 0
End Sub
