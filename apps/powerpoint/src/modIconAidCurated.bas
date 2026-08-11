Attribute VB_Name = "modIconAidCurated"
' IconAid - editable vector icons for PowerPoint.
'
' The web sidebar (IconAid task pane) drops each chosen icon onto the slide as a
' PICTURE tagged "IconAid:<id>[:<#hexcolor>]". "Make Editable" on the Insert
' ribbon tab finds those tagged pictures and rebuilds each as a native, fully
' editable freeform - adjustable stroke color, line weight, and (for filled
' designs) fill color.
'
' Icon geometry lives in an external data file, icons.dat (54k+ lines of
' "id|name|category|tags|subpath1|subpath2|..."), NOT in this VBA project - far
' more data than Mac PowerPoint can hold in a project. icons.dat ships in the
' installer, is copied into PowerPoint's SlideAid folder, and is loaded into
' memory once here. Subpaths are pre-normalized to absolute M/L/C/Z by
' scripts/svg_normalize.py, so the parser below only needs to handle those.
Option Explicit

Private Const ICON_VIEWBOX As Single = 24

' In-memory cache of icon records, loaded once from icons.dat.
Private mData() As String        ' 1-based: mData(1..mCuratedCount)
Private mCuratedCount As Long
Private mCuratedLoaded As Boolean
Private mById As Collection      ' icon id -> index into mData

' Editable-icon style.
Public Type CuratedIconStyle
    StrokeColor As Long      ' Line color (RGB)
    FillColor As Long        ' Fill color (RGB, -1 for none)
    StrokeWidth As Single    ' Line width in points
    Size As Single           ' Icon size in points
End Type

' Fill a style with defaults. NOTE: VBA (especially on Mac) requires user-defined
' types to be passed ByRef and cannot return them from functions, so this fills
' the caller's style ByRef rather than returning one.
Public Sub SetDefaultCuratedStyle(ByRef style As CuratedIconStyle)
    style.StrokeColor = RGB(31, 41, 55)   ' dark gray
    style.FillColor = -1                  ' no fill
    style.StrokeWidth = 1.5
    style.Size = 48
End Sub

' =====================================================================
' Data layer: load icons.dat once, serve records from memory.
' =====================================================================

' icons.dat is placed in PowerPoint's SlideAid folder by install.command (and by
' tools/build.sh for development). Environ("HOME") is the sandbox container Data
' directory; also try the real-home container path as a fallback.
Private Function CuratedDataPath() As String
    Dim c1 As String, c2 As String
    c1 = Environ("HOME") & "/SlideAid/icons.dat"
    c2 = Environ("HOME") & "/Library/Containers/com.microsoft.Powerpoint/Data/SlideAid/icons.dat"
    If Len(Dir(c1)) > 0 Then
        CuratedDataPath = c1
    ElseIf Len(Dir(c2)) > 0 Then
        CuratedDataPath = c2
    Else
        CuratedDataPath = c1
    End If
End Function

Private Sub EnsureCuratedLoaded()
    If mCuratedLoaded Then Exit Sub
    mCuratedLoaded = True            ' set first so a failure doesn't retry every call
    mCuratedCount = 0
    Set mById = New Collection

    Dim p As String
    p = CuratedDataPath()
    If Len(Dir(p)) = 0 Then Exit Sub

    ' Read the whole file at once and Split - far faster than a Line Input loop.
    Dim f As Integer, whole As String
    On Error GoTo Fail
    f = FreeFile
    Open p For Input As #f
    whole = Input$(LOF(f), f)
    Close #f

    Dim arr() As String, i As Long, lineStr As String, cutPos As Long
    arr = Split(whole, vbLf)
    ReDim mData(1 To UBound(arr) + 2)
    For i = 0 To UBound(arr)
        lineStr = arr(i)
        If Len(lineStr) > 0 Then
            If Right$(lineStr, 1) = vbCr Then lineStr = Left$(lineStr, Len(lineStr) - 1)
        End If
        If Len(lineStr) > 0 Then
            mCuratedCount = mCuratedCount + 1
            mData(mCuratedCount) = lineStr
            cutPos = InStr(lineStr, "|")
            If cutPos > 1 Then
                On Error Resume Next
                mById.Add mCuratedCount, Left$(lineStr, cutPos - 1)   ' key = icon id
                On Error GoTo Fail
            End If
        End If
    Next i
    Exit Sub
Fail:
    On Error Resume Next
    Close #f
    On Error GoTo 0
End Sub

' Record "id|name|category|tags|path1|path2|..." at a 1-based index, or "".
Private Function GetCuratedIconData(ByVal idx As Long) As String
    EnsureCuratedLoaded
    If idx < 1 Or idx > mCuratedCount Then
        GetCuratedIconData = ""
    Else
        GetCuratedIconData = mData(idx)
    End If
End Function

' Index into mData for an icon id, or 0 if absent.
Private Function GetCuratedIndexById(ByVal iconId As String) As Long
    EnsureCuratedLoaded
    On Error Resume Next
    GetCuratedIndexById = mById(iconId)
    On Error GoTo 0
End Function

' =====================================================================
' Rendering: build one normalized M/L/C/Z subpath into a freeform shape.
' =====================================================================
Private Function CreateCuratedShape(ByVal sl As Slide, ByVal pathData As String, _
    ByVal baseLeft As Single, ByVal baseTop As Single, ByVal iconScale As Single, _
    ByRef style As CuratedIconStyle) As Shape

    Dim fb As FreeformBuilder
    Dim shp As Shape
    Dim curX As Single, curY As Single
    Dim startX As Single, startY As Single

    On Error GoTo ErrorHandler

    curX = 0: curY = 0
    startX = 0: startY = 0
    Set fb = Nothing

    Dim pos As Long, pathLen As Long
    Dim cmdChar As String, numStr As String
    Dim numList As Collection

    pathLen = Len(pathData)
    pos = 1

    Do While pos <= pathLen
        cmdChar = Mid(pathData, pos, 1)
        pos = pos + 1

        Do While pos <= pathLen And Mid(pathData, pos, 1) = " "
            pos = pos + 1
        Loop

        Set numList = New Collection
        numStr = ""
        Do While pos <= pathLen
            Dim c As String
            c = Mid(pathData, pos, 1)
            If c Like "[A-Za-z]" Then Exit Do
            If c = "," Or c = " " Then
                If Len(numStr) > 0 Then
                    numList.Add Val(numStr)
                    numStr = ""
                End If
            ElseIf c = "-" And Len(numStr) > 0 Then
                numList.Add Val(numStr)
                numStr = c
            Else
                numStr = numStr & c
            End If
            pos = pos + 1
        Loop
        If Len(numStr) > 0 Then
            numList.Add Val(numStr)
        End If

        Select Case UCase(cmdChar)
            Case "M"
                If numList.Count >= 2 Then
                    curX = numList(1) * iconScale
                    curY = numList(2) * iconScale
                    startX = curX: startY = curY
                    If fb Is Nothing Then
                        Set fb = sl.Shapes.BuildFreeform(msoEditingAuto, baseLeft + curX, baseTop + curY)
                    End If
                End If

            Case "L"
                If Not fb Is Nothing And numList.Count >= 2 Then
                    curX = numList(1) * iconScale
                    curY = numList(2) * iconScale
                    fb.AddNodes msoSegmentLine, msoEditingAuto, baseLeft + curX, baseTop + curY
                End If

            Case "C"
                If Not fb Is Nothing And numList.Count >= 6 Then
                    ' msoEditingCorner keeps the exact SVG control points (Auto
                    ' lets PowerPoint re-smooth them, which distorts the curve).
                    fb.AddNodes msoSegmentCurve, msoEditingCorner, _
                        baseLeft + numList(1) * iconScale, baseTop + numList(2) * iconScale, _
                        baseLeft + numList(3) * iconScale, baseTop + numList(4) * iconScale, _
                        baseLeft + numList(5) * iconScale, baseTop + numList(6) * iconScale
                    curX = numList(5) * iconScale
                    curY = numList(6) * iconScale
                End If

            Case "Z"
                If Not fb Is Nothing Then
                    If Abs(curX - startX) > 0.1 Or Abs(curY - startY) > 0.1 Then
                        fb.AddNodes msoSegmentLine, msoEditingAuto, baseLeft + startX, baseTop + startY
                    End If
                End If
        End Select
    Loop

    If Not fb Is Nothing Then
        Set shp = fb.ConvertToShape
        If style.FillColor >= 0 Then            ' filled design: solid fill, no outline
            shp.Fill.Solid
            shp.Fill.ForeColor.RGB = style.FillColor
            shp.Line.Visible = msoFalse
        Else                                     ' outline design: stroke, no fill
            shp.Fill.Visible = msoFalse
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = style.StrokeColor
            shp.Line.Weight = style.StrokeWidth
        End If
        Set CreateCuratedShape = shp
    End If
    Exit Function

ErrorHandler:
    Set CreateCuratedShape = Nothing
End Function

' =====================================================================
' "Make Editable": convert IconAid pictures dropped by the web sidebar
' (each tagged "IconAid:<id>" or "IconAid:<id>:<#hexcolor>") on the active
' slide into full editable freeforms - all of them, in one click.
' =====================================================================
Public Sub MakeIconsEditable()
    EnsureCuratedLoaded
    If mCuratedCount = 0 Then
        MsgBox "Icon data (icons.dat) was not found - reinstall Slide Aid.", _
               vbExclamation, "Slide Aid Icons"
        Exit Sub
    End If

    Dim sl As Slide
    On Error Resume Next
    Set sl = ActiveWindow.View.Slide
    On Error GoTo 0
    If sl Is Nothing Then
        MsgBox "Open a slide first.", vbExclamation, "Slide Aid Icons"
        Exit Sub
    End If

    ' Snapshot tagged pictures first (don't mutate the shape collection mid-loop).
    Dim doomed As New Collection, shp As Shape
    For Each shp In sl.Shapes
        If Len(IconTagOf(shp)) > 0 Then doomed.Add shp
    Next shp
    If doomed.Count = 0 Then
        MsgBox "No inserted icons on this slide to make editable." & vbCrLf & vbCrLf & _
               "Use Insert Icons on the Insert tab to add icons, then select " & _
               "them and click Make Editable.", _
               vbInformation, "Slide Aid Icons"
        Exit Sub
    End If

    Dim n As Long, notFound As Long, i As Long
    Dim iconTag As String, tagParts() As String, idx As Long, strokeCol As Long
    For i = 1 To doomed.Count
        Set shp = doomed(i)
        iconTag = IconTagOf(shp)                  ' "<id>" or "<id>:<#hex>"
        tagParts = Split(iconTag, ":")
        idx = GetCuratedIndexById(tagParts(0))
        strokeCol = -1
        If UBound(tagParts) >= 1 Then strokeCol = HexToRgb(tagParts(1))
        If idx = 0 Then
            notFound = notFound + 1
        ElseIf MaterializeIcon(sl, idx, shp.Left, shp.Top, shp.Width, shp.Height, strokeCol) Then
            shp.Delete
            n = n + 1
        End If
    Next i

    Dim msg As String
    msg = "Made " & n & " icon(s) editable. Recolor, restyle the outline, or " & _
          "reshape them like any PowerPoint shape."
    If notFound > 0 Then _
        msg = msg & vbCrLf & vbCrLf & notFound & " could not be matched in the local library."
    MsgBox msg, vbInformation, "Slide Aid Icons"
End Sub

' Returns "<id>[:<#hex>]" from an IconAid picture's tag, or "" if not tagged.
Private Function IconTagOf(ByVal shp As Shape) As String
    Dim s As String
    s = shp.Name
    If InStr(s, "IconAid:") <> 1 Then
        On Error Resume Next
        s = shp.AlternativeText          ' Office.js altTextDescription
        On Error GoTo 0
    End If
    If InStr(s, "IconAid:") = 1 Then IconTagOf = Mid$(s, 9)
End Function

' Build icon idx as an editable freeform fitted into (left0,top0,w,h).
Private Function MaterializeIcon(ByVal sl As Slide, ByVal idx As Long, _
    ByVal left0 As Single, ByVal top0 As Single, ByVal w As Single, ByVal h As Single, _
    ByVal strokeCol As Long) As Boolean
    Dim rec As String: rec = GetCuratedIconData(idx)
    If Len(rec) = 0 Then Exit Function
    Dim parts() As String: parts = Split(rec, "|")
    If UBound(parts) < 4 Then Exit Function

    Dim st As CuratedIconStyle: SetDefaultCuratedStyle st
    If strokeCol >= 0 Then st.StrokeColor = strokeCol
    Dim isFilled As Boolean: isFilled = IconIsFilled(parts(0))
    If isFilled Then st.FillColor = st.StrokeColor    ' filled design -> solid fill
    Dim boxSz As Single: boxSz = w: If h < boxSz Then boxSz = h
    st.Size = boxSz
    Dim sc As Single: sc = boxSz / ICON_VIEWBOX

    Dim grp As New Collection, i As Long, shp As Shape
    For i = 4 To UBound(parts)
        If Len(parts(i)) > 0 Then
            Set shp = CreateCuratedShape(sl, parts(i), left0, top0, sc, st)
            If Not shp Is Nothing Then grp.Add shp
        End If
    Next i
    If grp.Count = 0 Then Exit Function

    Dim result As Shape
    If grp.Count = 1 Then
        Set result = grp(1)
    Else
        Dim nm() As String, j As Long
        ReDim nm(1 To grp.Count)
        For j = 1 To grp.Count: nm(j) = grp(j).Name: Next j

        Dim merged As Boolean: merged = False
        If isFilled Then
            ' Combine overlapping contours so even-odd holes (e.g. the gap in a
            ' ring) survive - leaves ONE editable shape whose fill and outline you
            ' can color independently. MergeShapes / msoMergeCombine are NOT in
            ' every Mac PowerPoint's VBA type library; naming a missing constant
            ' would stop this whole module compiling ("compile error in hidden
            ' module"). So call it LATE-BOUND: the module always compiles, and
            ' where the method is absent the call fails at runtime and we fall
            ' back to a plain group (then holes render filled). 2 = msoMergeCombine.
            On Error Resume Next
            CallByName sl.Shapes.Range(nm), "MergeShapes", VbMethod, 2
            If Err.Number = 0 Then merged = True
            On Error GoTo 0
        End If
        If merged Then
            ' MergeShapes leaves the combined shape selected; grab it to name+lock.
            On Error Resume Next
            Set result = ActiveWindow.Selection.ShapeRange(1)
            On Error GoTo 0
        Else
            Set result = sl.Shapes.Range(nm).Group
        End If
    End If

    If Not result Is Nothing Then
        result.Name = "Icon_" & parts(0)
        ' Lock the aspect ratio so dragging any handle scales the icon uniformly
        ' (no squashing) - matches how PowerPoint's own inserted icons behave.
        On Error Resume Next
        result.LockAspectRatio = msoTrue
        On Error GoTo 0
    End If
    MaterializeIcon = True
End Function

' Bootstrap icons (and any heroicons solid/mini) are filled shapes; everything
' else is a stroke outline. Mirrors the sidebar's isFilled().
Private Function IconIsFilled(ByVal iconId As String) As Boolean
    Dim s As String: s = LCase$(iconId)
    IconIsFilled = (Left$(s, 10) = "bootstrap-") Or (Right$(s, 6) = "-solid") Or (Right$(s, 5) = "-mini")
End Function

' "#RRGGBB" (or "RRGGBB") -> RGB Long; -1 if invalid.
Private Function HexToRgb(ByVal hx As String) As Long
    HexToRgb = -1
    hx = Replace(hx, "#", "")
    If Len(hx) <> 6 Then Exit Function
    On Error GoTo Bad
    HexToRgb = RGB(CLng("&H" & Mid$(hx, 1, 2)), CLng("&H" & Mid$(hx, 3, 2)), CLng("&H" & Mid$(hx, 5, 2)))
    Exit Function
Bad:
    HexToRgb = -1
End Function
