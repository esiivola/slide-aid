Attribute VB_Name = "modStub"
' =====================================================================
' Slide Aid - STUB ADD-IN for development.
' Mac PowerPoint renders custom ribbons in documents but only FIRES
' their callbacks from loaded add-ins. This stub owns the ribbon and
' forwards every click to the dev presentation via Application.Run.
'
' Build once:
'   1. blank presentation -> VBE -> import this module
'   2. Save As "Slide Aid.ppam" into apps/powerpoint/dist
'   3. cd apps/powerpoint && python3 tools/inject_ribbon.py "dist/Slide Aid.ppam"
'   4. Tools > PowerPoint Add-ins > + > load it, restart PowerPoint
'
' Daily loop: keep slideAidDev.pptm open, edit code there, Cmd+S,
' click the ribbon. Rebuild the stub only when the ribbon XML changes.
' =====================================================================
Option Explicit

Private Const DEV_FILE As String = "slideAidDev.pptm"

Public Sub RB_Dispatch(control As IRibbonControl)
    On Error GoTo fail
    Application.Run DEV_FILE & "!RB_DispatchTag", control.Tag
    Exit Sub
fail:
    MsgBox "Open " & DEV_FILE & " first - the Slide Aid code lives there. (" _
           & Err.Description & ")", vbExclamation, "Slide Aid"
End Sub

Public Sub RB_GetLibraryMenu(control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    returnedVal = Application.Run(DEV_FILE & "!LibraryMenuXML")
End Sub

Public Sub RB_GetFormatsMenu(control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    returnedVal = Application.Run(DEV_FILE & "!FormatsMenuXML")
End Sub
