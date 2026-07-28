Attribute VB_Name = "modHelpers"
' =====================================================================
' Slide Aid - Shared helpers
' Conventions:
'  - "Master" = LAST shape in the selection ShapeRange (PowerPoint
'    preserves selection order in ShapeRange).
'  - All geometry in points internally; user input in cm.
' =====================================================================
Option Explicit

Public Const CM_TO_PT As Single = 28.3464567

' ---------- Selection ----------

' Returns the selected ShapeRange, or Nothing (with optional message).
Public Function GetSelection(Optional ByVal minCount As Long = 1, _
                             Optional ByVal warn As Boolean = True) As ShapeRange
    On Error GoTo NoSel
    Dim sel As Selection
    Set sel = ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes And sel.Type <> ppSelectionText Then GoTo NoSel
    If sel.ShapeRange.Count < minCount Then GoTo NoSel
    Set GetSelection = sel.ShapeRange
    Exit Function
NoSel:
    If warn Then MsgBox "Please select at least " & minCount & " object(s).", vbExclamation, "Slide Aid"
    Set GetSelection = Nothing
End Function

' Master = last selected shape.
Public Function GetMaster(sr As ShapeRange) As Shape
    Set GetMaster = sr(sr.Count)
End Function

Public Function CurrentSlide() As Slide
    Set CurrentSlide = ActiveWindow.View.Slide
End Function

Public Function SlideW() As Single
    SlideW = ActivePresentation.PageSetup.SlideWidth
End Function

Public Function SlideH() As Single
    SlideH = ActivePresentation.PageSetup.SlideHeight
End Function

' ---------- Geometry ----------

Public Function ShpRight(s As Shape) As Single
    ShpRight = s.Left + s.Width
End Function

Public Function ShpBottom(s As Shape) As Single
    ShpBottom = s.Top + s.Height
End Function

Public Function ShpCenterX(s As Shape) As Single
    ShpCenterX = s.Left + s.Width / 2
End Function

Public Function ShpCenterY(s As Shape) As Single
    ShpCenterY = s.Top + s.Height / 2
End Function

' Horizontal / vertical overlap tests (do projections intersect?)
Public Function OverlapsH(a As Shape, b As Shape) As Boolean
    OverlapsH = (a.Left < ShpRight(b)) And (ShpRight(a) > b.Left)
End Function

Public Function OverlapsV(a As Shape, b As Shape) As Boolean
    OverlapsV = (a.Top < ShpBottom(b)) And (ShpBottom(a) > b.Top)
End Function

' ---------- Persistent input defaults ----------
' Dialogs remember the last-used value via a small
' key=value file in the store. Pass a prefKey to AskCm/AskInt.

Public Function PrefsPath() As String
    PrefsPath = StoreDir() & "/prefs.txt"
End Function

Public Function GetPref(ByVal key As String, ByVal dflt As String) As String
    GetPref = dflt
    If Dir(PrefsPath()) = "" Then Exit Function
    Dim f As Integer, ln As String, p As Long
    f = FreeFile
    On Error GoTo Done                ' Open failed: nothing to close
    Open PrefsPath() For Input As #f
    On Error GoTo CloseIt             ' from here on the file IS open
    Do While Not EOF(f)
        Line Input #f, ln
        p = InStr(ln, "=")
        If p > 1 Then
            If LCase$(Left$(ln, p - 1)) = LCase$(key) Then GetPref = Mid$(ln, p + 1)
        End If
    Loop
CloseIt:
    Close #f
Done:
    On Error GoTo 0
End Function

Public Sub SetPref(ByVal key As String, ByVal value As String)
    Dim keys(1 To 100) As String, vals(1 To 100) As String, n As Long
    Dim f As Integer, ln As String, p As Long, i As Long, hit As Boolean
    If Dir(PrefsPath()) <> "" Then
        f = FreeFile
        On Error GoTo ReadDone
        Open PrefsPath() For Input As #f
        On Error GoTo ReadClose
        Do While Not EOF(f)
            Line Input #f, ln
            p = InStr(ln, "=")
            If p > 1 And n < 100 Then
                n = n + 1
                keys(n) = Left$(ln, p - 1)
                vals(n) = Mid$(ln, p + 1)
            End If
        Loop
ReadClose:
        Close #f
ReadDone:
        On Error GoTo 0
    End If
    For i = 1 To n
        If LCase$(keys(i)) = LCase$(key) Then
            vals(i) = value
            hit = True
        End If
    Next i
    If Not hit And n < 100 Then
        n = n + 1: keys(n) = key: vals(n) = value
    End If
    On Error Resume Next              ' prefs are a convenience - never fail a tool
    EnsureStore
    f = FreeFile
    Open PrefsPath() For Output As #f
    Print #f, "# Slide Aid - remembered dialog defaults."
    For i = 1 To n
        Print #f, keys(i) & "=" & vals(i)
    Next i
    Close #f
    On Error GoTo 0
End Sub

' ---------- Input ----------

' Ask user for a length in cm (negative allowed, e.g. overlapping
' gaps). Returns points; sets ok=False on cancel/invalid input.
' With prefKey given, the last-used value becomes the next default.
Public Function AskCm(ByVal prompt As String, ByVal defaultCm As String, _
                      ByRef ok As Boolean, Optional ByVal prefKey As String = "") As Single
    ok = False
    If Len(prefKey) > 0 Then defaultCm = GetPref(prefKey, defaultCm)
    Dim s As String
    s = InputBox(prompt, "Slide Aid", defaultCm)
    If Len(Trim$(s)) = 0 Then Exit Function
    s = Replace(s, ",", ".")
    If Not IsNumeric(s) Then Exit Function
    ok = True
    If Len(prefKey) > 0 Then SetPref prefKey, s
    AskCm = CSng(Val(s)) * CM_TO_PT
End Function

' Ask user for an integer. Returns -1 on cancel/invalid.
' With prefKey given, the last-used value becomes the next default.
Public Function AskInt(ByVal prompt As String, ByVal defaultVal As String, _
                       Optional ByVal prefKey As String = "") As Long
    If Len(prefKey) > 0 Then defaultVal = GetPref(prefKey, defaultVal)
    Dim s As String
    s = InputBox(prompt, "Slide Aid", defaultVal)
    If Len(Trim$(s)) = 0 Or Not IsNumeric(s) Then
        AskInt = -1
        Exit Function
    End If
    AskInt = CLng(Val(s))
    If Len(prefKey) > 0 Then SetPref prefKey, CStr(AskInt)
End Function

' ---------- Environment ----------

' Sandboxed PowerPoint reports its container as HOME
' (~/Library/Containers/com.microsoft.Powerpoint/Data) - strip that
' to get the real user home.
Public Function RealHome() As String
    Dim h As String, i As Long
    h = Environ("HOME")
    i = InStr(h, "/Library/Containers/")
    If i > 0 Then h = Left$(h, i - 1)
    RealHome = h
End Function

' ---------- Timing ----------

' Wait while keeping PowerPoint responsive (used around flaky
' operations like SaveAs, a known Mac PowerPoint issue).
Public Sub Pause(ByVal seconds As Single)
    Dim t As Single
    t = Timer
    Do While Timer < t + seconds
        DoEvents
    Loop
End Sub

' ---------- Colors ----------

' Copy a color, preserving the theme link if the source has one.
' (Shared by modColors, modPainter.)
Public Sub CopyColorFormat(src As ColorFormat, dst As ColorFormat)
    On Error Resume Next
    If src.ObjectThemeColor <> msoNotThemeColor Then
        dst.ObjectThemeColor = src.ObjectThemeColor
        dst.TintAndShade = src.TintAndShade
        dst.Brightness = src.Brightness
    Else
        dst.RGB = src.RGB
    End If
    On Error GoTo 0
End Sub

' ---------- Sorting ----------

' Fill arr() with shape indices of sr sorted by position.
' axis = "H" (left-to-right) or "V" (top-to-bottom).
Public Sub SortIndicesByPosition(sr As ShapeRange, axis As String, arr() As Long)
    Dim n As Long, i As Long, j As Long, tmp As Long
    n = sr.Count
    ReDim arr(1 To n)
    For i = 1 To n: arr(i) = i: Next i
    For i = 1 To n - 1
        For j = 1 To n - i
            Dim a As Single, b As Single
            If axis = "H" Then
                a = sr(arr(j)).Left: b = sr(arr(j + 1)).Left
            Else
                a = sr(arr(j)).Top: b = sr(arr(j + 1)).Top
            End If
            If a > b Then
                tmp = arr(j): arr(j) = arr(j + 1): arr(j + 1) = tmp
            End If
        Next j
    Next i
End Sub
