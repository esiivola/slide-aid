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

Private Const REPO As String = "/Users/eerosiivola/repos/slide-aid"

' Prefer the folder this presentation is saved in (the dev .pptm lives
' in the repo); fall back to the REPO constant for unsaved blanks.
Private Function RepoDir() As String
    RepoDir = ActivePresentation.Path
    If Len(RepoDir) = 0 Then RepoDir = REPO
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
               "'Slide Aid' as PowerPoint Add-In (.ppam) into the repo folder.", _
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
    repo = REPO                       ' const above; the build runs from
                                      ' the loaded add-in, so the active
                                      ' presentation's path is unrelated

    ' Pre-authorize sandbox file access (no per-file prompts).
    On Error Resume Next
    GrantAccessToMultipleFiles Array(repo & "/", repo & "/src/", repo & "/tools/")
    On Error GoTo 0

    ' 1. fresh presentation with all modules
    Dim p As Presentation
    On Error GoTo Fail
    Set p = Presentations.Add(WithWindow:=msoTrue)
    Dim f As String, n As Long
    f = Dir(repo & "/src/*.bas")
    Do While Len(f) > 0
        p.VBProject.VBComponents.Import repo & "/src/" & f
        n = n + 1
        f = Dir()
    Loop
    p.VBProject.VBComponents.Import repo & "/tools/import_helper.bas"
    If n = 0 Then
        p.Close
        MsgBox "No modules found in " & repo & "/src - check the REPO " & _
               "constant in modImportHelper.", vbExclamation, "Slide Aid build"
        Exit Sub
    End If

    ' 2. save as .pptm (retry - Mac SaveAs right after VBE work is flaky)
    Dim i As Long, saved As Boolean
    For i = 1 To 3
        On Error Resume Next
        Err.Clear
        DoEvents
        p.SaveAs repo & "/Slide Aid.pptm", ppSaveAsOpenXMLPresentationMacroEnabled
        saved = (Err.Number = 0)
        On Error GoTo Fail
        If saved Then Exit For
        LocalPause 0.7
    Next i
    If Not saved Then
        p.Close
        MsgBox "SaveAs to '" & repo & "/Slide Aid.pptm' failed - grant " & _
               "file access if macOS asked, then run BuildSlideAid again.", _
               vbExclamation, "Slide Aid build"
        Exit Sub
    End If
    p.Close

    ' 3. inject ribbon + icons, convert to .ppam
    Dim res As String
    On Error GoTo NoHelper
    res = AppleScriptTask("SlideAidUI.scpt", "buildPpam", repo)
    On Error GoTo Fail
    If Left$(res, 2) <> "OK" Then
        MsgBox "The injector reported an error:" & vbCr & res, _
               vbExclamation, "Slide Aid build"
        Exit Sub
    End If

    MsgBox "Slide Aid.ppam rebuilt (" & n + 1 & " modules)." & vbCr & vbCr & _
           "Restart PowerPoint to load the new build.", _
           vbInformation, "Slide Aid build"
    Exit Sub

NoHelper:
    MsgBox "Built 'Slide Aid.pptm' (" & n + 1 & " modules), but the " & _
           "SlideAidUI helper is missing its buildPpam handler." & vbCr & vbCr & _
           "Recompile it (see README), or finish manually:" & vbCr & _
           "python3 tools/inject_ribbon.py --make-ppam ""Slide Aid.pptm""", _
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
