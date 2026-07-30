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

    Dim f As String, n As Long
    f = Dir(RepoDir() & "/src/*.bas")
    Do While Len(f) > 0
        ActivePresentation.VBProject.VBComponents.Import RepoDir() & "/src/" & f
        n = n + 1
        f = Dir()
    Loop
    If n = 0 Then
        MsgBox "Nothing imported - grant file access to the repo and run again.", _
               vbExclamation, "Slide Aid"
    Else
        MsgBox n & " modules imported. Next: Debug > Compile, then File > Save As -> " & _
               "'Slide Aid' as PowerPoint Add-In (.ppam) into apps/powerpoint/dist.", _
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
    Dim f As String, n As Long, sourceDir As String
    sourceDir = BuildDir() & "/src"
    f = Dir(sourceDir & "/*.bas")
    Do While Len(f) > 0
        p.VBProject.VBComponents.Import sourceDir & "/" & f
        n = n + 1
        f = Dir()
    Loop
    If n = 0 Then
        p.Close
        MsgBox "No modules found in PowerPoint's SlideAid build cache. " & _
               "Run tools/build.sh from the repository first.", _
               vbExclamation, "Slide Aid build"
        Exit Sub
    End If

    ' 2. save as .pptm inside PowerPoint's writable sandbox
    Dim i As Long, saved As Boolean
    EnsureBuildDir
    For i = 1 To 3
        On Error Resume Next
        Err.Clear
        DoEvents
        p.SaveAs BuildPptmPath(), ppSaveAsOpenXMLPresentationMacroEnabled
        saved = (Err.Number = 0)
        On Error GoTo Fail
        If saved Then Exit For
        LocalPause 0.7
    Next i
    If Not saved Then
        p.Close
        MsgBox "SaveAs to PowerPoint's SlideAid build folder failed. " & _
               "Run BuildSlideAid again; if it repeats, restart PowerPoint.", _
               vbExclamation, "Slide Aid build"
        Exit Sub
    End If
    p.Close

    ' 3. inject ribbon + icons, convert to .ppam
    Dim res As String
    On Error GoTo NoHelper
    res = AppleScriptTask("SlideAidUI.scpt", "buildPpam", repo & vbLf & BuildPptmPath())
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
           BuildPptmPath() & """", _
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

    Dim f As String, n As Long
    f = Dir(RepoDir() & "/src/*.bas")
    Do While Len(f) > 0
        proj.VBComponents.Import RepoDir() & "/src/" & f
        n = n + 1
        f = Dir()
    Loop
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
