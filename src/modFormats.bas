Attribute VB_Name = "modFormats"
' =====================================================================
' Slide Aid - "My Formats"
' Saved shape formats stored as pipe-delimited lines in
' ~/SlideAid/formats.txt. The ribbon menu is built dynamically.
' Line: name|fillVis|fillRGB|lineVis|lineRGB|lineW|dash|font|size|bold|italic|fontRGB|mL|mR|mT|mB|wrap
' =====================================================================
Option Explicit

Public Function FormatsPath() As String
    FormatsPath = StoreDir() & "/formats.txt"
End Function

' ---------------------------------------------------------------
' Save the Master's (last selected) format under a chosen name.
' ---------------------------------------------------------------
Public Sub SaveFormatFromMaster()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim m As Shape
    Set m = GetMaster(sr)

    Dim nm As String
    nm = InputBox("Name for this format:", "Slide Aid", "Format")
    If Len(Trim$(nm)) = 0 Then Exit Sub
    nm = Replace(nm, "|", "/")

    Dim parts(0 To 16) As String
    parts(0) = nm
    On Error Resume Next
    parts(1) = CLng(m.Fill.Visible):  parts(2) = m.Fill.ForeColor.RGB
    parts(3) = CLng(m.Line.Visible):  parts(4) = m.Line.ForeColor.RGB
    parts(5) = m.Line.Weight:         parts(6) = m.Line.DashStyle
    If m.HasTextFrame Then
        With m.TextFrame
            parts(7) = .TextRange.Font.Name
            parts(8) = .TextRange.Font.Size
            parts(9) = CLng(.TextRange.Font.Bold)
            parts(10) = CLng(.TextRange.Font.Italic)
            parts(11) = .TextRange.Font.Color.RGB
            parts(12) = .MarginLeft: parts(13) = .MarginRight
            parts(14) = .MarginTop:  parts(15) = .MarginBottom
            parts(16) = CLng(.WordWrap)
        End With
    End If
    On Error GoTo 0

    EnsureStore
    Dim f As Integer
    f = FreeFile
    Open FormatsPath() For Append As #f
    Print #f, Join(parts, "|")
    Close #f
    RefreshDynamicMenus
    MsgBox "'" & nm & "' saved to My Formats.", vbInformation, "Slide Aid"
End Sub

' ---------------------------------------------------------------
' Apply saved format #idx to all selected shapes.
' ---------------------------------------------------------------
Public Sub ApplySavedFormat(ByVal idx As Long)
    Dim lines() As String
    If Not ReadFormatLines(lines) Then Exit Sub
    If idx < 0 Or idx > UBound(lines) Then Exit Sub

    Dim p() As String
    p = Split(lines(idx), "|")
    If UBound(p) < 16 Then Exit Sub

    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim s As Shape
    For Each s In sr
        On Error Resume Next
        s.Fill.Visible = CLng(p(1))
        If CLng(p(1)) = msoTrue Then s.Fill.ForeColor.RGB = CLng(p(2))
        s.Line.Visible = CLng(p(3))
        If CLng(p(3)) = msoTrue Then
            s.Line.ForeColor.RGB = CLng(p(4))
            s.Line.Weight = CSng(p(5))
            s.Line.DashStyle = CLng(p(6))
        End If
        If s.HasTextFrame And Len(p(7)) > 0 Then
            Dim r As TextRange
            For Each r In s.TextFrame.TextRange.Runs
                r.Font.Name = p(7)
                r.Font.Size = CSng(p(8))
                r.Font.Bold = CLng(p(9))
                r.Font.Italic = CLng(p(10))
                r.Font.Color.RGB = CLng(p(11))
            Next r
            With s.TextFrame
                .MarginLeft = CSng(p(12)): .MarginRight = CSng(p(13))
                .MarginTop = CSng(p(14)):  .MarginBottom = CSng(p(15))
                .WordWrap = CLng(p(16))
            End With
        End If
        On Error GoTo 0
    Next s
End Sub

Public Sub ShowFormatsFile()
    MsgBox "My Formats are stored in:" & vbCr & FormatsPath() & vbCr & vbCr & _
           "Edit or delete lines in any text editor to manage them.", _
           vbInformation, "Slide Aid"
End Sub

Private Function ReadFormatLines(ByRef lines() As String) As Boolean
    ReadFormatLines = False
    If Dir(FormatsPath()) = "" Then Exit Function
    Dim f As Integer, txt As String
    f = FreeFile
    Open FormatsPath() For Input As #f
    txt = Input$(LOF(f), f)
    Close #f
    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    Do While Right$(txt, 1) = vbLf
        txt = Left$(txt, Len(txt) - 1)
    Loop
    If Len(txt) = 0 Then Exit Function
    lines = Split(txt, vbLf)
    ReadFormatLines = True
End Function

' Menu XML for the ribbon dynamicMenu (called from modRibbon).
Public Function FormatsMenuXML() As String
    Dim xml As String
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"

    Dim lines() As String
    If ReadFormatLines(lines) Then
        Dim i As Long, nm As String
        For i = 0 To UBound(lines)
            nm = Split(lines(i), "|")(0)
            If Len(nm) > 0 Then
                xml = xml & "<button id=""saFmtI" & i & """ label=""" & XmlEsc(nm) & _
                      """ onAction=""RB_Dispatch"" tag=""FmtApply:" & i & """/>"
            End If
        Next i
        xml = xml & "<menuSeparator id=""saFmtSep""/>"
    End If

    xml = xml & "<button id=""saFmtSave"" label=""Save Master's Format..."" " & _
          "onAction=""RB_Dispatch"" tag=""FmtSave:0""/>"
    xml = xml & "<button id=""saFmtMan"" label=""Manage My Formats..."" " & _
          "onAction=""RB_Dispatch"" tag=""FmtOpen:0""/>"
    xml = xml & "</menu>"
    FormatsMenuXML = xml
End Function
