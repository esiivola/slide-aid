Attribute VB_Name = "modReview"
' =====================================================================
' Slide Aid - on-slide review markup: sticky comment notes, TODO/EDIT
' markers, and callouts that point at an object.
'
' Unlike PowerPoint's native comments (which live in a side pane and do
' not print), these are real shapes stamped with the reviewer's initials
' and the date, drawn in deliberately loud fixed colors so they cannot be
' missed. Every mark is tagged "SA_REVIEW" so the whole deck can be swept
' clean before the final send.
'
' Initials come from Application.UserName (PowerPoint's registered author
' name) with no setup, are cached in prefs, and can be changed any time.
' =====================================================================
Option Explicit

Private Const REVIEW_TAG As String = "SA_REVIEW"
Private Const NOTE_W As Single = 156
Private Const NOTE_H As Single = 52
Private Const CORNER_MARGIN As Single = 10
Private Const CASCADE As Single = 16

' ---------------------------------------------------------------
' Reviewer initials: derived from the account name on first use,
' cached in prefs, editable via SetReviewInitials.
' ---------------------------------------------------------------
Public Function ReviewInitials() As String
    Dim v As String
    v = GetPref("ReviewInitials", "")
    If Len(v) = 0 Then
        v = DeriveInitials(Application.UserName)
        If Len(v) = 0 Then v = "?"
        SetPref "ReviewInitials", v
    End If
    ReviewInitials = v
End Function

Private Function DeriveInitials(ByVal nm As String) As String
    Dim words() As String, i As Long, r As String
    nm = Trim$(nm)
    If Len(nm) = 0 Then Exit Function
    words = Split(nm, " ")
    For i = LBound(words) To UBound(words)
        If Len(words(i)) > 0 Then r = r & UCase$(Left$(words(i), 1))
    Next i
    If Len(r) > 3 Then r = Left$(r, 3)
    DeriveInitials = r
End Function

Public Sub SetReviewInitials()
    Dim cur As String, v As String
    cur = ReviewInitials()
    v = InputBox("Your initials for review marks:", "Slide Aid", cur)
    If StrPtr(v) = 0 Then Exit Sub          ' Cancel (as opposed to an empty box)
    v = Trim$(v)
    If Len(v) > 0 Then SetPref "ReviewInitials", Left$(v, 6)
End Sub

' ---------------------------------------------------------------
' Add a sticky comment note (kind "NOTE"), or a TODO/EDIT marker,
' in the top-right corner, cascading below any marks already there.
' ---------------------------------------------------------------
Public Sub AddReviewNote(ByVal kind As String)
    Dim sl As Slide
    On Error Resume Next
    Set sl = ActiveWindow.View.Slide
    On Error GoTo 0
    If sl Is Nothing Then Exit Sub

    Dim body As String
    body = InputBox(PromptFor(kind), "Slide Aid - " & KindLabel(kind))
    If StrPtr(body) = 0 Then Exit Sub       ' Cancel

    Dim leftP As Single, topP As Single
    leftP = SlideW() - NOTE_W - CORNER_MARGIN
    topP = CORNER_MARGIN + ReviewCountOnSlide(sl) * CASCADE

    Dim s As Shape
    Set s = sl.Shapes.AddShape(msoShapeRoundedRectangle, leftP, topP, NOTE_W, NOTE_H)
    On Error Resume Next
    s.Adjustments(1) = 0.08
    On Error GoTo 0
    StyleNote s, kind
    s.TextFrame.TextRange.Text = HeaderFor(kind) & vbCrLf & body
    On Error Resume Next
    s.TextFrame.TextRange.Paragraphs(1, 1).Font.Bold = msoTrue
    On Error GoTo 0
    TagReview s, kind
End Sub

' ---------------------------------------------------------------
' Add a callout: a note plus a leader line to the selected object.
' ---------------------------------------------------------------
Public Sub AddReviewCallout()
    Dim sl As Slide, target As Shape
    On Error Resume Next
    Set sl = ActiveWindow.View.Slide
    If ActiveWindow.Selection.Type = ppSelectionShapes Then Set target = ActiveWindow.Selection.ShapeRange(1)
    On Error GoTo 0
    If sl Is Nothing Then Exit Sub
    If target Is Nothing Then
        MsgBox "Select the object you want to point at first, then click Callout.", vbInformation, "Slide Aid"
        Exit Sub
    End If

    Dim body As String
    body = InputBox("Callout note:", "Slide Aid - Callout")
    If StrPtr(body) = 0 Then Exit Sub

    ' Place the note above-right of the target, clamped to the slide.
    Dim leftP As Single, topP As Single
    leftP = ShpRight(target) + 16
    If leftP > SlideW() - NOTE_W - CORNER_MARGIN Then leftP = SlideW() - NOTE_W - CORNER_MARGIN
    If leftP < CORNER_MARGIN Then leftP = CORNER_MARGIN
    topP = target.Top - NOTE_H - 16
    If topP < CORNER_MARGIN Then topP = CORNER_MARGIN

    Dim note As Shape
    Set note = sl.Shapes.AddShape(msoShapeRoundedRectangle, leftP, topP, NOTE_W, NOTE_H)
    On Error Resume Next
    note.Adjustments(1) = 0.08
    On Error GoTo 0
    StyleNote note, "NOTE"
    note.TextFrame.TextRange.Text = HeaderFor("NOTE") & vbCrLf & body
    On Error Resume Next
    note.TextFrame.TextRange.Paragraphs(1, 1).Font.Bold = msoTrue
    On Error GoTo 0

    ' Leader from the note's bottom-center to the target's center.
    Dim ln As Shape
    Set ln = sl.Shapes.AddLine(leftP + NOTE_W / 2, topP + NOTE_H, ShpCenterX(target), ShpCenterY(target))
    ln.Line.ForeColor.RGB = RGB(255, 59, 48)
    ln.Line.Weight = 2.25
    ln.Line.EndArrowheadStyle = msoArrowheadTriangle

    Dim g As Shape
    Set g = sl.Shapes.Range(Array(note.Name, ln.Name)).Group
    TagReview g, "CALLOUT"
End Sub

' ---------------------------------------------------------------
' Remove every Slide Aid review mark across the whole deck.
' ---------------------------------------------------------------
Public Sub RemoveReviewMarkup()
    Dim doomed As New Collection, sl As Slide, shp As Shape, i As Long
    For Each sl In ActivePresentation.Slides
        For Each shp In sl.Shapes
            If HasReviewTag(shp) Then doomed.Add shp
        Next shp
    Next sl
    If doomed.Count = 0 Then
        MsgBox "No Slide Aid review marks were found in this deck.", vbInformation, "Slide Aid"
        Exit Sub
    End If
    If MsgBox("Remove " & doomed.Count & " review mark(s) from the whole deck?", _
              vbQuestion + vbYesNo, "Slide Aid") <> vbYes Then Exit Sub
    For i = 1 To doomed.Count
        On Error Resume Next
        doomed(i).Delete
        On Error GoTo 0
    Next i
End Sub

' ---------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------
Private Sub StyleNote(ByVal s As Shape, ByVal kind As String)
    ' Loud, fixed colors - review marks must be impossible to miss and are
    ' removed before export, so matching the deck theme is a non-goal.
    Dim fillC As Long, textC As Long
    Select Case UCase$(kind)
        Case "TODO": fillC = RGB(255, 59, 48):  textC = RGB(255, 255, 255)   ' red
        Case "EDIT": fillC = RGB(255, 159, 10): textC = RGB(17, 17, 17)      ' orange
        Case Else:   fillC = RGB(255, 214, 10): textC = RGB(17, 17, 17)      ' yellow
    End Select
    s.Fill.Solid
    s.Fill.ForeColor.RGB = fillC
    s.Line.Visible = msoFalse
    With s.TextFrame
        .MarginLeft = 4: .MarginRight = 4: .MarginTop = 3: .MarginBottom = 3
        .WordWrap = msoTrue
        On Error Resume Next
        .AutoSize = ppAutoSizeShapeToFitText
        On Error GoTo 0
        With .TextRange
            .Font.Size = 9
            .Font.Name = "Arial"
            .Font.Color.RGB = textC
            .ParagraphFormat.Alignment = ppAlignLeft
        End With
    End With
End Sub

Private Sub TagReview(ByVal s As Shape, ByVal kind As String)
    On Error Resume Next
    s.Tags.Add REVIEW_TAG, kind
    On Error GoTo 0
    s.Name = "SAReview_" & kind & "_" & CStr(Int(Timer * 1000))
End Sub

Private Function HasReviewTag(ByVal s As Shape) As Boolean
    On Error Resume Next
    HasReviewTag = (Len(s.Tags(REVIEW_TAG)) > 0)
    On Error GoTo 0
End Function

Private Function ReviewCountOnSlide(ByVal sl As Slide) As Long
    Dim shp As Shape, n As Long
    For Each shp In sl.Shapes
        If HasReviewTag(shp) Then n = n + 1
    Next shp
    ReviewCountOnSlide = n
End Function

Private Function HeaderFor(ByVal kind As String) As String
    Dim stamp As String
    stamp = ReviewInitials() & " " & Chr$(183) & " " & Format$(Date, "d mmm")
    If UCase$(kind) = "NOTE" Then
        HeaderFor = stamp
    Else
        HeaderFor = UCase$(kind) & " " & Chr$(183) & " " & stamp
    End If
End Function

Private Function KindLabel(ByVal kind As String) As String
    Select Case UCase$(kind)
        Case "TODO": KindLabel = "To-Do"
        Case "EDIT": KindLabel = "Edit"
        Case Else:   KindLabel = "Comment"
    End Select
End Function

Private Function PromptFor(ByVal kind As String) As String
    Select Case UCase$(kind)
        Case "TODO": PromptFor = "To-do:"
        Case "EDIT": PromptFor = "Edit needed:"
        Case Else:   PromptFor = "Comment:"
    End Select
End Function
