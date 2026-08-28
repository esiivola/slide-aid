Attribute VB_Name = "modReview"
' =====================================================================
' Slide Aid - on-slide review markup, modelled on the reference review
' toolbar (BCG "Cool Macros" / PPT Productivity):
'
'   * Sticky Notes  - a coloured note (6 colours, yellow default) docked
'                     top-right, stamped with the reviewer's initials and
'                     the date, with the comment below.
'   * Status Stamps - a bold banner (Draft, WIP, Confidential, ...) toggled
'                     on the slide.
'   * Callout       - a note with a leader line to the selected object.
'
' Unlike PowerPoint's native comments (side pane, do not print) these are
' real shapes that show in exported PDFs. Every mark is tagged so the deck
' can be swept clean before the final send.
'
' Notes/stamps are plain (sharp-cornered) rectangles with no shadow, to
' match the reference. Initials are seeded with no setup from the account
' name (app user name where available, else the Mac login), cached in
' prefs, and editable.
' =====================================================================
Option Explicit

Private Const NOTE_TAG As String = "SA_REVIEW"     ' sticky notes + callouts
Private Const STAMP_TAG As String = "SA_STAMP"     ' status stamps
Private Const NOTE_W As Single = 156
Private Const NOTE_H As Single = 50
Private Const CORNER_MARGIN As Single = 10
Private Const CASCADE As Single = 14
Private Const NOTE_TEXT As Long = 2171169          ' RGB(34,34,34) dark grey

' ---------------------------------------------------------------
' Reviewer initials: derived from the account name on first use,
' cached in prefs, editable via SetReviewInitials.
' ---------------------------------------------------------------
Public Function ReviewInitials() As String
    Dim v As String
    v = GetPref("ReviewInitials", "")
    If Len(v) = 0 Then
        v = DeriveInitials(AccountName())
        If Len(v) = 0 Then v = "?"
        SetPref "ReviewInitials", v
    End If
    ReviewInitials = v
End Function

' Account name for seeding initials, with no setup. PowerPoint's Application
' object (unlike Word/Excel) has no UserName property on every build, so read
' it LATE-BOUND - a missing property fails at runtime instead of breaking
' compilation - and fall back to the Mac login name from the environment.
Private Function AccountName() As String
    Dim nm As String, app As Object
    Set app = Application
    On Error Resume Next
    nm = app.UserName
    On Error GoTo 0
    If Len(Trim$(nm)) = 0 Then nm = Environ("USER")
    AccountName = Trim$(nm)
End Function

' Compress a name or login into up to three uppercase initials. Handles
' "First Last", first.last, first_last and first-last. Falls back to the whole
' name for a single-token login rather than a lonely letter.
Private Function DeriveInitials(ByVal nm As String) As String
    Dim t As String, words() As String, i As Long, r As String
    t = Trim$(nm)
    If Len(t) = 0 Then Exit Function
    t = Replace(t, ".", " ")
    t = Replace(t, "_", " ")
    t = Replace(t, "-", " ")
    words = Split(t, " ")
    For i = LBound(words) To UBound(words)
        If Len(words(i)) > 0 Then r = r & UCase$(Left$(words(i), 1))
    Next i
    If Len(r) > 3 Then r = Left$(r, 3)
    If Len(r) < 2 Then r = Trim$(nm)          ' single token: show the name itself
    DeriveInitials = r
End Function

Public Sub SetReviewInitials()
    Dim cur As String, v As String
    cur = ReviewInitials()
    v = InputBox("Your initials for review marks:", "Slide Aid", cur)
    v = Trim$(v)
    If Len(v) = 0 Then Exit Sub               ' Cancel or empty box: leave unchanged
    SetPref "ReviewInitials", Left$(v, 6)
End Sub

' ---------------------------------------------------------------
' Sticky note in one of six colours (colourKey), docked top-right and
' cascading below any notes already on the slide.
' ---------------------------------------------------------------
Public Sub AddReviewNote(ByVal colourKey As String)
    Dim sl As Slide
    On Error Resume Next
    Set sl = ActiveWindow.View.Slide
    On Error GoTo 0
    If sl Is Nothing Then Exit Sub

    Dim body As String
    body = InputBox("Comment:", "Slide Aid - Sticky Note")
    If Len(body) = 0 Then Exit Sub            ' Cancel or empty box

    Dim leftP As Single, topP As Single
    leftP = SlideW() - NOTE_W - CORNER_MARGIN
    topP = CORNER_MARGIN + NoteCountOnSlide(sl) * (NOTE_H + CASCADE)

    Dim s As Shape
    Set s = sl.Shapes.AddShape(msoShapeRectangle, leftP, topP, NOTE_W, NOTE_H)
    StyleNote s, colourKey
    s.TextFrame.TextRange.Text = StampLine() & vbCrLf & body
    BoldFirstLine s
    TagShape s, NOTE_TAG, "NOTE:" & UCase$(colourKey)
End Sub

' ---------------------------------------------------------------
' Callout: a yellow note plus a leader line to the selected object.
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
    If Len(body) = 0 Then Exit Sub

    Dim leftP As Single, topP As Single
    leftP = ShpRight(target) + 16
    If leftP > SlideW() - NOTE_W - CORNER_MARGIN Then leftP = SlideW() - NOTE_W - CORNER_MARGIN
    If leftP < CORNER_MARGIN Then leftP = CORNER_MARGIN
    topP = target.Top - NOTE_H - 16
    If topP < CORNER_MARGIN Then topP = CORNER_MARGIN

    Dim note As Shape
    Set note = sl.Shapes.AddShape(msoShapeRectangle, leftP, topP, NOTE_W, NOTE_H)
    StyleNote note, "YELLOW"
    note.TextFrame.TextRange.Text = StampLine() & vbCrLf & body
    BoldFirstLine note

    Dim ln As Shape
    Set ln = sl.Shapes.AddLine(leftP + NOTE_W / 2, topP + NOTE_H, ShpCenterX(target), ShpCenterY(target))
    ln.Line.ForeColor.RGB = RGB(230, 60, 50)
    ln.Line.Weight = 2#

    Dim g As Shape
    Set g = sl.Shapes.Range(Array(note.Name, ln.Name)).Group
    TagShape g, NOTE_TAG, "CALLOUT"
End Sub

' ---------------------------------------------------------------
' Status stamp: a bold banner (Draft, WIP, ...) toggled on this slide.
' Clicking the same stamp again removes it.
' ---------------------------------------------------------------
Public Sub AddStatusStamp(ByVal kind As String)
    Dim sl As Slide
    On Error Resume Next
    Set sl = ActiveWindow.View.Slide
    On Error GoTo 0
    If sl Is Nothing Then Exit Sub

    ' Toggle off: if this stamp is already on the slide, remove it and stop.
    Dim shp As Shape
    For Each shp In sl.Shapes
        If ShapeTag(shp, STAMP_TAG) = UCase$(kind) Then
            shp.Delete
            Exit Sub
        End If
    Next shp

    Dim s As Shape
    Set s = sl.Shapes.AddShape(msoShapeRectangle, 0, 18, 240, 40)
    s.Fill.Solid
    s.Fill.ForeColor.RGB = StampColor(kind)
    s.Line.Visible = msoFalse
    With s.TextFrame
        .MarginLeft = 10: .MarginRight = 10: .MarginTop = 3: .MarginBottom = 3
        .WordWrap = msoFalse
        On Error Resume Next
        .AutoSize = ppAutoSizeShapeToFitText
        On Error GoTo 0
        With .TextRange
            .Text = StampLabel(kind)
            .Font.Size = 20
            .Font.Bold = msoTrue
            .Font.Name = "Arial"
            .Font.Color.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = ppAlignCenter
        End With
    End With
    s.Left = (SlideW() - s.Width) / 2          ' centre horizontally, after autosize
    s.Top = 18
    TagShape s, STAMP_TAG, UCase$(kind)
End Sub

' ---------------------------------------------------------------
' Remove every Slide Aid review mark (notes, callouts, stamps) across
' the whole deck.
' ---------------------------------------------------------------
Public Sub RemoveReviewMarkup()
    Dim doomed As New Collection, sl As Slide, shp As Shape, i As Long
    For Each sl In ActivePresentation.Slides
        For Each shp In sl.Shapes
            If Len(ShapeTag(shp, NOTE_TAG)) > 0 Or Len(ShapeTag(shp, STAMP_TAG)) > 0 Then doomed.Add shp
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
Private Sub StyleNote(ByVal s As Shape, ByVal colourKey As String)
    s.Fill.Solid
    s.Fill.ForeColor.RGB = NoteColor(colourKey)
    s.Line.Visible = msoFalse                  ' sharp rectangle, no border, no shadow
    With s.TextFrame
        .MarginLeft = 5: .MarginRight = 5: .MarginTop = 3: .MarginBottom = 3
        .WordWrap = msoTrue
        On Error Resume Next
        .AutoSize = ppAutoSizeShapeToFitText
        On Error GoTo 0
        With .TextRange
            .Font.Size = 9
            .Font.Name = "Arial"
            .Font.Color.RGB = NOTE_TEXT
            .ParagraphFormat.Alignment = ppAlignLeft
        End With
    End With
End Sub

Private Sub BoldFirstLine(ByVal s As Shape)
    On Error Resume Next
    s.TextFrame.TextRange.Paragraphs(1, 1).Font.Bold = msoTrue
    On Error GoTo 0
End Sub

' Header line: "ES <middot> 28 Aug". ChrW gives a real Unicode middot;
' Chr(183) is "sum" in Mac's code page.
Private Function StampLine() As String
    StampLine = ReviewInitials() & " " & ChrW$(183) & " " & Format$(Date, "d mmm")
End Function

Private Sub TagShape(ByVal s As Shape, ByVal tagName As String, ByVal value As String)
    On Error Resume Next
    s.Tags.Add tagName, value
    On Error GoTo 0
    s.Name = "SAReview_" & value & "_" & CStr(Int(Timer * 1000))
End Sub

Private Function ShapeTag(ByVal s As Shape, ByVal tagName As String) As String
    On Error Resume Next
    ShapeTag = s.Tags(tagName)
    On Error GoTo 0
End Function

Private Function NoteCountOnSlide(ByVal sl As Slide) As Long
    Dim shp As Shape, n As Long
    For Each shp In sl.Shapes
        If Len(ShapeTag(shp, NOTE_TAG)) > 0 Then n = n + 1
    Next shp
    NoteCountOnSlide = n
End Function

' The six reference note colours; yellow is the default.
Private Function NoteColor(ByVal key As String) As Long
    Select Case UCase$(key)
        Case "GREEN":    NoteColor = RGB(178, 223, 138)
        Case "BLUE":     NoteColor = RGB(160, 210, 255)
        Case "PINK":     NoteColor = RGB(255, 179, 198)
        Case "LAVENDER": NoteColor = RGB(214, 196, 240)
        Case "LAVBLUE":  NoteColor = RGB(190, 200, 245)
        Case Else:       NoteColor = RGB(255, 224, 79)   ' yellow (default)
    End Select
End Function

Private Function StampColor(ByVal kind As String) As Long
    Select Case UCase$(kind)
        Case "NEW":         StampColor = RGB(40, 160, 70)
        Case "UPDATED":     StampColor = RGB(40, 110, 200)
        Case "WIP":         StampColor = RGB(230, 140, 0)
        Case "ONHOLD":      StampColor = RGB(130, 70, 160)
        Case "DRAFT":       StampColor = RGB(120, 120, 120)
        Case Else:          StampColor = RGB(200, 30, 30)   ' Confidential / Out of Date / Remove
    End Select
End Function

Private Function StampLabel(ByVal kind As String) As String
    Select Case UCase$(kind)
        Case "WIP":       StampLabel = "WORK IN PROGRESS"
        Case "ONHOLD":    StampLabel = "ON HOLD"
        Case "OUTOFDATE": StampLabel = "OUT OF DATE"
        Case Else:        StampLabel = UCase$(kind)
    End Select
End Function
