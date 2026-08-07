Attribute VB_Name = "modImportHelper"
' =====================================================================
' Slide Aid - dev helper. Import ONLY this file into a blank
' presentation. Safe to leave in place (it ships inside the .ppam
' but has no ribbon buttons).
'
'   ImportAllModules - pull all src/*.bas into this presentation.
'                      Run ONCE per fresh presentation (running twice
'                      creates duplicate modules -> compile errors).
'   ExportAllModules - write this presentation's modules back to
'                      src/ (run after editing code in the VBE, so
'                      git stays the source of truth).
' =====================================================================
Option Explicit

' PowerPoint can always write below Environ("HOME"), which macOS maps
' to the app's sandbox Data directory. The external helper moves the
' completed add-in from here into the Git checkout.
Private Function BuildDir() As String
    BuildDir = Environ("HOME") & "/SlideAid/build"
End Function

Private Function BuildPptmPath() As String
    BuildPptmPath = BuildDir() & "/Slide Aid.pptm"
End Function

Private Sub EnsureBuildDir()
    Dim root As String
    root = Environ("HOME") & "/SlideAid"
    If Dir(root, vbDirectory) = "" Then MkDir root
    If Dir(BuildDir(), vbDirectory) = "" Then MkDir BuildDir()
End Sub

' tools/build.sh writes this developer-local path before invoking the
' macro. Keeping it outside the add-in avoids embedding a workstation
' username or checkout location in release artifacts.
Private Function ConfiguredRepoDir() As String
    Dim configPath As String, f As Integer, value As String
    configPath = Environ("HOME") & "/SlideAid/repo_path.txt"
    If Dir(configPath) = "" Then Exit Function

    On Error GoTo Done
    f = FreeFile
    Open configPath For Input As #f
    Line Input #f, value
    Close #f
    ConfiguredRepoDir = Trim$(value)
Done:
    On Error GoTo 0
End Function

' Prefer the folder this presentation is saved in (the dev .pptm lives
' in the repo); fall back to the configured checkout for unsaved blanks.
Private Function RepoDir() As String
    RepoDir = ActivePresentation.Path
    If Len(RepoDir) = 0 Then RepoDir = ConfiguredRepoDir()
End Function

' Gather every .bas path up front. CRITICAL: do not call VBComponents.Import
' while iterating Dir(). VBA keeps a single global Dir() search state, and
' Import disturbs it, so a fetch-then-import loop silently stops after the
' first file (the classic "only 1 module imported" bug). Collect first, then
' import from the returned list.
Private Function CollectBasFiles(ByVal folder As String) As Collection
    Dim names As New Collection
    Dim f As String
    f = Dir(folder & "/*.bas")
    Do While Len(f) > 0
        names.Add folder & "/" & f
        f = Dir()
    Loop
    Set CollectBasFiles = names
End Function

Private Function GetBaseName(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "/")
    If i > 0 Then GetBaseName = Mid$(p, i + 1) Else GetBaseName = p
End Function

' The dev helper / stub must never be re-imported into a presentation that
' is already running them, or the duplicate module breaks compilation.
Private Function IsHelperModule(ByVal p As String) As Boolean
    Dim b As String
    b = LCase$(GetBaseName(p))
    IsHelperModule = (b = "modimporthelper.bas" Or b = "modstub.bas")
End Function

Public Sub ImportAllModules()
    ' Guard against double import
    Dim comp As Object
    On Error Resume Next
    Set comp = ActivePresentation.VBProject.VBComponents("modHelpers")
    On Error GoTo 0
    If Not comp Is Nothing Then
        MsgBox "Modules are already imported - aborting to avoid duplicates. " & _
               "Use ExportAllModules / manual VBE editing instead.", _
               vbExclamation, "Slide Aid"
        Exit Sub
    End If

    ' Prefer this checkout's src/; fall back to the sandbox cache that
    ' tools/build.sh populates (readable even when the repo folder is not).
    Dim srcDir As String, files As Collection
    srcDir = RepoDir() & "/src"
    Set files = CollectBasFiles(srcDir)
    If files.Count = 0 Then
        srcDir = BuildDir() & "/src"
        Set files = CollectBasFiles(srcDir)
    End If

    Dim basPath As Variant, n As Long, failed As String
    For Each basPath In files
        If Not IsHelperModule(CStr(basPath)) Then
            On Error Resume Next
            Err.Clear
            ActivePresentation.VBProject.VBComponents.Import CStr(basPath)
            If Err.Number = 0 Then
                n = n + 1
            Else
                failed = failed & vbCr & "  " & GetBaseName(CStr(basPath)) & _
                         " (" & Err.Description & ")"
            End If
            On Error GoTo 0
        End If
    Next basPath

    If n = 0 Then
        MsgBox "Nothing imported from:" & vbCr & srcDir & vbCr & vbCr & _
               "Grant file access to the repo (or run tools/build.sh first) and retry.", _
               vbExclamation, "Slide Aid"
    ElseIf Len(failed) > 0 Then
        MsgBox n & " modules imported, but some FAILED:" & failed & vbCr & vbCr & _
               "Fix those and re-run in a fresh presentation.", vbExclamation, "Slide Aid"
    Else
        MsgBox n & " modules imported from:" & vbCr & srcDir & vbCr & vbCr & _
               "Next: Debug > Compile (VBE), then File > Save As -> a macro-enabled " & _
               ".pptm, and run tools/inject_ribbon.py --make-ppam on it.", _
               vbInformation, "Slide Aid"
    End If
End Sub

' =====================================================================
' ONE-COMMAND BUILD: fresh presentation <- src/*.bas -> Slide Aid.pptm
' -> inject_ribbon.py --make-ppam (via the SlideAidUI helper) ->
' Slide Aid.ppam. Requires: AccessVBOM enabled (once) and the helper
' compiled with the buildPpam handler (osacompile, see README).
'
' Run it via  tools/build.sh  from Terminal, or in PowerPoint:
' Tools > Macro > Macros... > type "BuildSlideAid" > Run.
' Restart PowerPoint afterwards to load the new build.
' =====================================================================
Public Sub BuildSlideAid()
    Dim repo As String
    repo = ConfiguredRepoDir()
    If Len(repo) = 0 Then
        MsgBox "Repository path is not configured. Run tools/build.sh " & _
               "from the PowerPoint app folder first.", _
               vbExclamation, "Slide Aid build"
        Exit Sub
    End If

    ' 1. fresh presentation with all modules. tools/build.sh copied the
    ' source cache into PowerPoint's sandbox before invoking this macro.
    Dim p As Presentation
    On Error GoTo Fail
    Set p = Presentations.Add(WithWindow:=msoTrue)
    ' A slideless presentation cannot be saved (SaveAs yields a 0-byte file
    ' and reports failure). Presentations.Add creates zero slides, so add one.
    If p.Slides.Count = 0 Then p.Slides.Add 1, ppLayoutBlank
    Dim n As Long, sourceDir As String
    sourceDir = BuildDir() & "/src"
    Dim files As Collection, basPath As Variant
    Set files = CollectBasFiles(sourceDir)
    For Each basPath In files
        p.VBProject.VBComponents.Import CStr(basPath)
        n = n + 1
    Next basPath
    If n = 0 Then
        p.Close
        MsgBox "No modules found in PowerPoint's SlideAid build cache. " & _
               "Run tools/build.sh from the repository first.", _
               vbExclamation, "Slide Aid build"
        Exit Sub
    End If

    ' 2. save as .pptm. Programmatic SaveAs is unreliable in the Mac sandbox,
    ' so try the container build folder then a no-space fallback, verify the
    ' file is non-empty, and on total failure leave the presentation OPEN so
    ' the modules can be saved via the native File > Save As dialog.
    Dim i As Long, ti As Long, saved As Boolean, savedPath As String, lastErr As String
    Dim targets(1 To 2) As String
    EnsureBuildDir
    targets(1) = BuildPptmPath()
    targets(2) = Environ("HOME") & "/SlideAidBuild.pptm"
    For ti = 1 To 2
        For i = 1 To 3
            On Error Resume Next
            Err.Clear
            If Len(Dir(targets(ti))) > 0 Then Kill targets(ti)   ' clear a 0-byte remnant
            Err.Clear
            DoEvents
            p.SaveAs targets(ti), ppSaveAsOpenXMLPresentationMacroEnabled
            If Err.Number = 0 And FileLen(targets(ti)) > 0 Then
                saved = True
                savedPath = targets(ti)
            Else
                lastErr = "#" & Err.Number & " " & Err.Description
            End If
            On Error GoTo Fail
            If saved Then Exit For
            LocalPause 0.7
        Next i
        If saved Then Exit For
    Next ti
    If Not saved Then
        MsgBox "Automatic save failed (" & lastErr & ")." & vbCr & vbCr & _
               "The new presentation with all " & n & " modules is still OPEN." & vbCr & _
               "Save it manually: File > Save As > 'PowerPoint Macro-Enabled " & _
               "Presentation (.pptm)', then run" & vbCr & _
               "tools/inject_ribbon.py --make-ppam on that file.", _
               vbExclamation, "Slide Aid build"
        Exit Sub
    End If
    p.Close

    ' 3. inject ribbon + icons, convert to .ppam
    Dim res As String
    On Error GoTo NoHelper
    res = AppleScriptTask("SlideAidUI.scpt", "buildPpam", repo & vbLf & savedPath)
    On Error GoTo Fail
    If Left$(res, 2) <> "OK" Then
        MsgBox "The injector reported an error:" & vbCr & res, _
               vbExclamation, "Slide Aid build"
        Exit Sub
    End If

    MsgBox "dist/Slide Aid.ppam rebuilt (" & n & " modules)." & vbCr & vbCr & _
           "Restart PowerPoint to load the new build.", _
           vbInformation, "Slide Aid build"
    Exit Sub

NoHelper:
    MsgBox "Built the temporary 'Slide Aid.pptm' (" & n & " modules), but the " & _
           "SlideAidUI helper is missing its buildPpam handler." & vbCr & vbCr & _
           "Recompile it (see README), or finish manually:" & vbCr & _
           "python3 apps/powerpoint/tools/inject_ribbon.py --make-ppam """ & _
           savedPath & """", _
           vbExclamation, "Slide Aid build"
    Exit Sub
Fail:
    MsgBox "Build failed: " & Err.Description, vbExclamation, "Slide Aid build"
End Sub

' Wait while keeping PowerPoint responsive. Local copy so this module
' also works standalone (without modHelpers) in a fresh presentation.
Private Sub LocalPause(ByVal seconds As Single)
    Dim t As Single
    t = Timer
    Do While Timer < t + seconds
        DoEvents
    Loop
End Sub

' =====================================================================
' FAST DEV LOOP (no rebuild): in your dev .pptm, replace all modules
' with the current src/*.bas. Edit code in any editor -> RefreshModules
' -> click the ribbon (stub add-in forwards to this file). Seconds
' instead of a full rebuild + restart.
' =====================================================================
Public Sub RefreshModules()
    Dim proj As Object, comp As Object
    Set proj = ActivePresentation.VBProject

    ' collect first - removing while iterating is unsafe
    Dim doomed As New Collection
    For Each comp In proj.VBComponents
        If comp.Type = 1 Then                       ' standard modules
            If comp.Name <> "modImportHelper" And comp.Name <> "modStub" Then
                doomed.Add comp
            End If
        End If
    Next comp
    For Each comp In doomed
        proj.VBComponents.Remove comp
    Next comp

    Dim n As Long
    Dim files As Collection, basPath As Variant
    Set files = CollectBasFiles(RepoDir() & "/src")
    For Each basPath In files
        If Not IsHelperModule(CStr(basPath)) Then
            proj.VBComponents.Import CStr(basPath)
            n = n + 1
        End If
    Next basPath
    MsgBox n & " modules refreshed from src/ - test away.", _
           vbInformation, "Slide Aid dev"
End Sub

Public Sub ExportAllModules()
    Dim comp As Object, n As Long, failed As String
    For Each comp In ActivePresentation.VBProject.VBComponents
        If comp.Type = 1 Then                       ' standard modules
            If comp.Name <> "modImportHelper" Then
                On Error Resume Next
                Err.Clear
                comp.Export RepoDir() & "/src/" & comp.Name & ".bas"
                If Err.Number = 0 Then
                    n = n + 1
                Else
                    failed = failed & vbCr & "  " & comp.Name & " (" & Err.Description & ")"
                End If
                On Error GoTo 0
            End If
        End If
    Next comp
    If Len(failed) > 0 Then
        MsgBox n & " modules exported to " & RepoDir() & "/src" & vbCr & _
               "FAILED:" & failed, vbExclamation, "Slide Aid"
    Else
        MsgBox n & " modules exported to " & RepoDir() & "/src", vbInformation, "Slide Aid"
    End If
End Sub
