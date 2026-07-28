Attribute VB_Name = "modLibrary"
' =====================================================================
' Slide Aid - "My Elements" library
' Reusable slide elements stored in ~/SlideAid/library.pptx,
' one element per slide. The ribbon menu is built dynamically from
' the library file (see RB_GetLibraryMenu in modRibbon).
' =====================================================================
Option Explicit

Public Function StoreDir() As String
    StoreDir = Environ("HOME") & "/SlideAid"
End Function

Public Function LibraryPath() As String
    LibraryPath = StoreDir() & "/library.pptx"
End Function

Public Sub EnsureStore()
    On Error Resume Next
    If Dir(StoreDir(), vbDirectory) = "" Then MkDir StoreDir()
    On Error GoTo 0
End Sub

Private Function LibraryExists() As Boolean
    LibraryExists = (Dir(LibraryPath()) <> "")
End Function

' ---------------------------------------------------------------
' Add the current selection to the library as a new element.
' Uses a visible window (clipboard paste and SaveAs are unreliable
' on windowless presentations on Mac) - expect a brief flash.
' ---------------------------------------------------------------
Public Sub AddSelectionToLibrary()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim elName As String
    elName = InputBox("Name for this element:", "Slide Aid", "Element")
    If Len(Trim$(elName)) = 0 Then Exit Sub

    sr.Copy

    EnsureStore
    Dim lib As Presentation, isNew As Boolean
    On Error GoTo Fail
    If LibraryExists() Then
        Set lib = Presentations.Open(LibraryPath(), WithWindow:=msoTrue)
    Else
        isNew = True
        Set lib = Presentations.Add(WithWindow:=msoTrue)
        lib.PageSetup.SlideWidth = ActivePresentation.PageSetup.SlideWidth
        lib.PageSetup.SlideHeight = ActivePresentation.PageSetup.SlideHeight
    End If

    Dim sl As Slide
    Set sl = lib.Slides.Add(lib.Slides.Count + 1, ppLayoutBlank)
    sl.Shapes.Paste
    sl.Tags.Add "SANAME", elName

    If Not SaveLibrary(lib, isNew) Then
        lib.Close
        MsgBox "Saving the library failed (known Mac SaveAs issue) - " & _
               "please try again.", vbExclamation, "Slide Aid"
        Exit Sub
    End If
    lib.Close
    RefreshDynamicMenus
    MsgBox "'" & elName & "' added to My Elements.", vbInformation, "Slide Aid"
    Exit Sub
Fail:
    MsgBox "Could not write the library (" & Err.Description & "). " & _
           "If macOS asked for file access, grant it and retry.", vbExclamation, "Slide Aid"
End Sub

' SaveAs is flaky on Mac right after clipboard operations - flush
' events and retry a few times.
Private Function SaveLibrary(lib As Presentation, ByVal isNew As Boolean) As Boolean
    Dim i As Long
    For i = 1 To 3
        On Error Resume Next
        Err.Clear
        DoEvents
        If isNew Then
            lib.SaveAs LibraryPath()
        Else
            lib.Save
        End If
        If Err.Number = 0 Then
            On Error GoTo 0
            SaveLibrary = True
            Exit Function
        End If
        On Error GoTo 0
        Pause 0.7
    Next i
    SaveLibrary = False
End Function

' ---------------------------------------------------------------
' Insert element #idx on the current slide.
' ---------------------------------------------------------------
Public Sub InsertLibraryElement(ByVal idx As Long)
    If Not LibraryExists() Then
        MsgBox "No elements yet. Select shapes and use 'Add Selection to My Elements' first.", _
               vbInformation, "Slide Aid"
        Exit Sub
    End If

    ' Visible window: clipboard copy is unreliable from windowless
    ' presentations on Mac (brief flash is expected).
    Dim lib As Presentation, dst As Slide, copied As Boolean
    Set dst = CurrentSlide()
    On Error GoTo Fail
    Set lib = Presentations.Open(LibraryPath(), ReadOnly:=msoTrue, WithWindow:=msoTrue)
    If idx >= 1 And idx <= lib.Slides.Count Then
        lib.Slides(idx).Shapes.Range.Copy
        DoEvents
        copied = True
    End If
    lib.Close
    If copied Then dst.Shapes.Paste   ' never paste stale clipboard content
    Exit Sub
Fail:
    MsgBox "Could not open the library: " & Err.Description, vbExclamation, "Slide Aid"
End Sub

' Open the library for manual editing (rename via slide tags, delete
' slides, reorder). Save and close when done.
Public Sub OpenLibraryForEditing()
    If Not LibraryExists() Then
        MsgBox "No library yet - add an element first.", vbInformation, "Slide Aid"
        Exit Sub
    End If
    Presentations.Open LibraryPath(), WithWindow:=msoTrue
End Sub

' Menu XML for the ribbon dynamicMenu (called from modRibbon).
Public Function LibraryMenuXML() As String
    Dim xml As String
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"

    If LibraryExists() Then
        Dim lib As Presentation
        On Error Resume Next
        Set lib = Presentations.Open(LibraryPath(), ReadOnly:=msoTrue, WithWindow:=msoFalse)
        If Not lib Is Nothing Then
            Dim i As Long, nm As String
            For i = 1 To lib.Slides.Count
                nm = lib.Slides(i).Tags("SANAME")
                If Len(nm) = 0 Then nm = "Element " & i
                xml = xml & "<button id=""saLibI" & i & """ label=""" & XmlEsc(nm) & _
                      """ onAction=""RB_Dispatch"" tag=""LibIns:" & i & """/>"
            Next i
            lib.Close
            If i > 1 Then xml = xml & "<menuSeparator id=""saLibSep""/>"
        End If
        On Error GoTo 0
    End If

    xml = xml & "<button id=""saLibAdd"" label=""Add Selection to My Elements..."" " & _
          "onAction=""RB_Dispatch"" tag=""LibAdd:0""/>"
    xml = xml & "<button id=""saLibMan"" label=""Manage Library..."" " & _
          "onAction=""RB_Dispatch"" tag=""LibOpen:0""/>"
    xml = xml & "</menu>"
    LibraryMenuXML = xml
End Function

Public Function XmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    XmlEsc = s
End Function
