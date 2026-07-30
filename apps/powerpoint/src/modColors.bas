Attribute VB_Name = "modColors"
' =====================================================================
' Slide Aid - Colors
'  * Palette follows the ACTIVE SLIDE's theme automatically.
'  * Generic fallback palette for extra, theme-independent colors.
'  * Theme -> RGB and RGB -> Theme conversion of the selection.
'  * Eyedropper: pick fill/line/font colors from the Master shape.
' =====================================================================
Option Explicit

' Generic fallback palette (edit freely: name, RGB)
Private Const PAL_COUNT As Long = 8
Private Function PaletteRGB(ByVal i As Long) As Long
    Select Case i
        Case 1: PaletteRGB = RGB(31, 73, 125)    ' Dark blue
        Case 2: PaletteRGB = RGB(79, 129, 189)   ' Blue
        Case 3: PaletteRGB = RGB(155, 187, 89)   ' Green
        Case 4: PaletteRGB = RGB(192, 80, 77)    ' Red
        Case 5: PaletteRGB = RGB(247, 150, 70)   ' Orange
        Case 6: PaletteRGB = RGB(128, 100, 162)  ' Purple
        Case 7: PaletteRGB = RGB(89, 89, 89)     ' Dark grey
        Case 8: PaletteRGB = RGB(217, 217, 217)  ' Light grey
    End Select
End Function

' ---------------------------------------------------------------
' APPLY a theme color to the selection.
' target: "F"=fill, "L"=line, "T"=text/font
' themeIdx: 1..10 -> msoThemeColorDark1, Light1, Dark2, Light2,
'           Accent1..Accent6
' Keeps the theme link, so shapes adapt when the template changes.
' ---------------------------------------------------------------
Public Sub ApplyThemeColor(ByVal target As String, ByVal themeIdx As Long)
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim tc As MsoThemeColorIndex
    tc = ThemeIndexFromOrdinal(themeIdx)

    Dim s As Shape
    For Each s In sr
        ApplyThemeColorToShape s, target, tc
    Next s
End Sub

' Apply a generic (RGB) palette color to the selection.
Public Sub ApplyPaletteColor(ByVal target As String, ByVal palIdx As Long)
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub
    If palIdx < 1 Or palIdx > PAL_COUNT Then Exit Sub

    Dim s As Shape
    For Each s In sr
        ApplyRGBToShape s, target, PaletteRGB(palIdx)
    Next s
End Sub

Private Function ThemeIndexFromOrdinal(ByVal i As Long) As MsoThemeColorIndex
    Select Case i
        Case 1: ThemeIndexFromOrdinal = msoThemeColorDark1
        Case 2: ThemeIndexFromOrdinal = msoThemeColorLight1
        Case 3: ThemeIndexFromOrdinal = msoThemeColorDark2
        Case 4: ThemeIndexFromOrdinal = msoThemeColorLight2
        Case Else: ThemeIndexFromOrdinal = msoThemeColorAccent1 + (i - 5)
    End Select
End Function

Private Sub ApplyThemeColorToShape(s As Shape, ByVal target As String, ByVal tc As MsoThemeColorIndex)
    On Error Resume Next
    If s.Type = msoGroup Then
        Dim g As Shape
        For Each g In s.GroupItems
            ApplyThemeColorToShape g, target, tc
        Next g
        Exit Sub
    End If
    Select Case target
        Case "F"
            s.Fill.Visible = msoTrue
            s.Fill.ForeColor.ObjectThemeColor = tc
        Case "L"
            s.Line.Visible = msoTrue
            s.Line.ForeColor.ObjectThemeColor = tc
        Case "T"
            If s.HasTextFrame Then
                s.TextFrame.TextRange.Font.Color.ObjectThemeColor = tc
            End If
    End Select
    On Error GoTo 0
End Sub

Private Sub ApplyRGBToShape(s As Shape, ByVal target As String, ByVal rgbVal As Long)
    On Error Resume Next
    If s.Type = msoGroup Then
        Dim g As Shape
        For Each g In s.GroupItems
            ApplyRGBToShape g, target, rgbVal
        Next g
        Exit Sub
    End If
    Select Case target
        Case "F"
            s.Fill.Visible = msoTrue
            s.Fill.ForeColor.RGB = rgbVal
        Case "L"
            s.Line.Visible = msoTrue
            s.Line.ForeColor.RGB = rgbVal
        Case "T"
            If s.HasTextFrame Then s.TextFrame.TextRange.Font.Color.RGB = rgbVal
    End Select
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------
' THEME -> RGB: freeze all theme-linked colors in the selection
' as plain RGB, making them independent of the slide master/theme.
' ---------------------------------------------------------------
Public Sub ThemeToRGB()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    Dim s As Shape
    For Each s In sr
        FreezeShapeColors s
    Next s
End Sub

Private Sub FreezeShapeColors(s As Shape)
    On Error Resume Next
    If s.Type = msoGroup Then
        Dim g As Shape
        For Each g In s.GroupItems
            FreezeShapeColors g
        Next g
        Exit Sub
    End If
    FreezeColorFormat s.Fill.ForeColor
    FreezeColorFormat s.Fill.BackColor
    FreezeColorFormat s.Line.ForeColor
    If s.HasTextFrame Then
        Dim r As TextRange
        For Each r In s.TextFrame.TextRange.Runs
            FreezeColorFormat r.Font.Color
        Next r
    End If
    On Error GoTo 0
End Sub

Private Sub FreezeColorFormat(cf As ColorFormat)
    On Error Resume Next
    If cf.Type = msoColorTypeScheme Or cf.ObjectThemeColor <> msoNotThemeColor Then
        Dim v As Long
        v = cf.RGB          ' final RGB incl. tint/shade
        cf.RGB = v          ' assigning RGB breaks the theme link
    End If
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------
' RGB -> THEME: re-link plain RGB colors in the selection to the
' current theme, when they exactly match a theme color.
' ---------------------------------------------------------------
Public Sub RGBToTheme()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub

    ' Theme colors of the active slide
    Dim themeRGB(1 To 10) As Long
    Dim i As Long
    On Error Resume Next
    For i = 1 To 10
        themeRGB(i) = CurrentSlide().ThemeColorScheme(ThemeIndexFromOrdinal(i)).RGB
    Next i
    On Error GoTo 0

    Dim s As Shape
    For Each s In sr
        RelinkShapeColors s, themeRGB
    Next s
End Sub

Private Sub RelinkShapeColors(s As Shape, themeRGB() As Long)
    On Error Resume Next
    If s.Type = msoGroup Then
        Dim g As Shape
        For Each g In s.GroupItems
            RelinkShapeColors g, themeRGB
        Next g
        Exit Sub
    End If
    RelinkColorFormat s.Fill.ForeColor, themeRGB
    RelinkColorFormat s.Line.ForeColor, themeRGB
    If s.HasTextFrame Then
        Dim r As TextRange
        For Each r In s.TextFrame.TextRange.Runs
            RelinkColorFormat r.Font.Color, themeRGB
        Next r
    End If
    On Error GoTo 0
End Sub

Private Sub RelinkColorFormat(cf As ColorFormat, themeRGB() As Long)
    On Error Resume Next
    If cf.ObjectThemeColor = msoNotThemeColor Then
        Dim i As Long
        For i = 1 To 10
            If cf.RGB = themeRGB(i) Then
                cf.ObjectThemeColor = ThemeIndexFromOrdinal(i)
                Exit For
            End If
        Next i
    End If
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------
' EYEDROPPER: copy the Master's (last selected) fill, line and
' font color to all other selected shapes.
' what: "F", "L", "T" or "ALL"
' ---------------------------------------------------------------
Public Sub PickColorsFromMaster(ByVal what As String)
    Dim sr As ShapeRange
    Set sr = GetSelection(2)
    If sr Is Nothing Then Exit Sub

    Dim m As Shape
    Set m = GetMaster(sr)

    Dim i As Long, s As Shape
    For i = 1 To sr.Count - 1
        Set s = sr(i)
        On Error Resume Next
        If what = "F" Or what = "ALL" Then
            If m.Fill.Visible Then
                s.Fill.Visible = msoTrue
                CopyColorFormat m.Fill.ForeColor, s.Fill.ForeColor
            End If
        End If
        If what = "L" Or what = "ALL" Then
            If m.Line.Visible Then
                s.Line.Visible = msoTrue
                CopyColorFormat m.Line.ForeColor, s.Line.ForeColor
            End If
        End If
        If what = "T" Or what = "ALL" Then
            If m.HasTextFrame And s.HasTextFrame Then
                CopyColorFormat m.TextFrame.TextRange.Font.Color, s.TextFrame.TextRange.Font.Color
            End If
        End If
        On Error GoTo 0
    Next i
End Sub

' Show the RGB / theme info of the Master's fill (quick inspector).
Public Sub ShowColorInfo()
    Dim sr As ShapeRange
    Set sr = GetSelection(1)
    If sr Is Nothing Then Exit Sub
    Dim m As Shape
    Set m = GetMaster(sr)

    Dim v As Long, msg As String
    On Error Resume Next
    v = m.Fill.ForeColor.RGB
    msg = "Fill RGB: " & (v And &HFF) & ", " & ((v \ &H100) And &HFF) & ", " & ((v \ &H10000) And &HFF) & _
          "   Hex: #" & Right$("0" & Hex((v And &HFF)), 2) & Right$("0" & Hex((v \ &H100) And &HFF), 2) & Right$("0" & Hex((v \ &H10000) And &HFF), 2)
    If m.Fill.ForeColor.ObjectThemeColor <> msoNotThemeColor Then
        msg = msg & vbCr & "Theme-linked color (index " & m.Fill.ForeColor.ObjectThemeColor & ")"
    Else
        msg = msg & vbCr & "Plain RGB color (not theme-linked)"
    End If
    On Error GoTo 0
    MsgBox msg, vbInformation, "Slide Aid - Color info"
End Sub
