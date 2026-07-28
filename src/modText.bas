Attribute VB_Name = "modText"
' =====================================================================
' Slide Aid - Split / Merge text boxes
' =====================================================================
Option Explicit

' ---------------------------------------------------------------
' SPLIT the text box at the cursor position into two text boxes.
' Place the cursor where the split should happen, then run.
' Formatting is preserved (the box is duplicated, then trimmed).
' ---------------------------------------------------------------
Public Sub SplitTextBox()
    Dim sel As Selection
    On Error GoTo Fail
    Set sel = ActiveWindow.Selection
    If sel.Type <> ppSelectionText Then GoTo Fail

    Dim src As Shape
    Set src = sel.TextRange.Parent.Parent  ' TextRange -> TextFrame -> Shape
    Dim full As TextRange
    Set full = src.TextFrame.TextRange

    Dim pos As Long
    pos = sel.TextRange.Start   ' 1-based index of the insertion point
    If pos <= 1 Or pos > full.Length Then
        MsgBox "Place the cursor inside the text (not at the very start or end).", _
               vbExclamation, "Slide Aid"
        Exit Sub
    End If

    ' Duplicate keeps all formatting; trim each copy.
    Dim dup As Shape
    Set dup = src.Duplicate(1)
    dup.Left = src.Left
    dup.Top = ShpBottom(src) + 5 ' just below the original

    ' Original keeps chars 1 .. pos-1
    full.Characters(pos, full.Length - pos + 1).Delete
    ' Duplicate keeps chars pos .. end
    dup.TextFrame.TextRange.Characters(1, pos - 1).Delete

    ' Shrink boxes to fit if autosize is off
    Exit Sub
Fail:
    MsgBox "Click into a text box and place the cursor where the split should happen.", _
           vbExclamation, "Slide Aid"
End Sub

' ---------------------------------------------------------------
' MERGE the selected text boxes into one, in SELECTION ORDER.
' Each merged box becomes a new paragraph in the first box.
' Tries a clipboard paste first (keeps formatting); falls back to
' plain-text append.
' ---------------------------------------------------------------
Public Sub MergeTextBoxes()
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To sr.Count
        If Not sr(i).HasTextFrame Then
            MsgBox "All selected objects must contain text.", vbExclamation, "Slide Aid"
            Exit Sub
        End If
    Next i

    Dim dst As Shape
    Set dst = sr(1)

    ' Collect sources first (deleting while iterating a ShapeRange is unsafe)
    Dim sources As New Collection
    For i = 2 To sr.Count
        sources.Add sr(i)
    Next i

    Dim s As Shape
    For Each s In sources
        AppendText dst, s
    Next s
    For Each s In sources
        s.Delete
    Next s
    dst.Select
End Sub

' ---------------------------------------------------------------
' SET MARGINS: set all four internal margins of the selected text
' boxes/shapes at once.
' ---------------------------------------------------------------
Public Sub SetTextMargins()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim m As Single, ok As Boolean
    m = AskCm("Internal margin, all sides (cm):", "0.1", ok, "TextMargin")
    If Not ok Or m < 0 Then Exit Sub

    Dim i As Long
    For i = 1 To sr.Count
        On Error Resume Next
        If sr(i).HasTextFrame Then
            With sr(i).TextFrame
                .MarginLeft = m
                .MarginRight = m
                .MarginTop = m
                .MarginBottom = m
            End With
        End If
        On Error GoTo 0
    Next i
End Sub

' FIT FORM TO TEXT: shrink/grow the shape to its text.
Public Sub FitFormToText()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To sr.Count
        On Error Resume Next
        If sr(i).HasTextFrame Then
            sr(i).TextFrame.AutoSize = ppAutoSizeShapeToFitText
        End If
        On Error GoTo 0
    Next i
End Sub

' WRAP TEXT: toggle word wrap in the selected shapes.
Public Sub ToggleWrapText()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To sr.Count
        On Error Resume Next
        If sr(i).HasTextFrame Then
            sr(i).TextFrame.WordWrap = Not sr(i).TextFrame.WordWrap
        End If
        On Error GoTo 0
    Next i
End Sub

' ---------------------------------------------------------------
' CHANGE CASE of the selected shapes' text.
' mode: "U"=UPPER, "L"=lower, "T"=Title Case, "S"=Sentence case.
' Uses TextRange.ChangeCase, so per-run formatting is preserved.
' ---------------------------------------------------------------
Public Sub ChangeTextCase(ByVal mode As String)
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim cs As Long
    Select Case mode
        Case "U": cs = ppCaseUpper
        Case "L": cs = ppCaseLower
        Case "T": cs = ppCaseTitle
        Case "S": cs = ppCaseSentence
        Case Else: Exit Sub
    End Select

    Dim i As Long
    For i = 1 To sr.Count
        On Error Resume Next
        If sr(i).HasTextFrame Then sr(i).TextFrame.TextRange.ChangeCase cs
        On Error GoTo 0
    Next i
End Sub

' ---------------------------------------------------------------
' TIDY TEXT: collapse runs of spaces to one, in the selection.
' Uses TextRange.Replace, so formatting is preserved.
' ---------------------------------------------------------------
Public Sub TidyText()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim i As Long, guard As Long
    For i = 1 To sr.Count
        On Error Resume Next
        If sr(i).HasTextFrame Then
            guard = 0
            Do While InStr(sr(i).TextFrame.TextRange.Text, "  ") > 0 And guard < 200
                sr(i).TextFrame.TextRange.Replace "  ", " "
                guard = guard + 1
            Loop
        End If
        On Error GoTo 0
    Next i
End Sub

' ---------------------------------------------------------------
' SWAP TEXT between exactly two shapes. Each text adopts the
' receiving shape's default formatting.
' ---------------------------------------------------------------
Public Sub SwapText()
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub
    If sr.Count <> 2 Or Not sr(1).HasTextFrame Or Not sr(2).HasTextFrame Then
        MsgBox "Select exactly two objects that contain text.", vbExclamation, "Slide Aid"
        Exit Sub
    End If

    Dim t1 As String
    t1 = sr(1).TextFrame.TextRange.Text
    sr(1).TextFrame.TextRange.Text = sr(2).TextFrame.TextRange.Text
    sr(2).TextFrame.TextRange.Text = t1
End Sub

Private Sub AppendText(dst As Shape, src As Shape)
    Dim dr As TextRange
    Set dr = dst.TextFrame.TextRange
    dr.InsertAfter vbCr

    ' Formatting-preserving path: paste the copied source text after
    ' the destination text; fall back to plain text if paste fails.
    On Error Resume Next
    src.TextFrame.TextRange.Copy
    Dim lenBefore As Long
    lenBefore = dst.TextFrame.TextRange.Length
    dst.TextFrame.TextRange.Characters(lenBefore + 1).Paste
    If Err.Number <> 0 Or dst.TextFrame.TextRange.Length = lenBefore Then
        Err.Clear
        dr.InsertAfter src.TextFrame.TextRange.Text
    End If
    On Error GoTo 0
End Sub
