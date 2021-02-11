Attribute VB_Name = "lib_export"
Option Explicit

Sub export_all_vba_code()
'***************************************************************************
'Purpose: exports the vba code in separate text files to be used in source control
'Inputs
'Outputs:
'***************************************************************************
    Dim comp As VBComponent
    Dim codeFolder As String
    Dim FileName As String
    
    If Not CanAccessVBOM Then Exit Sub ' Exit if access to VB object model is not allowed
    
    If (ThisWorkbook.VBProject.VBE.ActiveWindow Is Nothing) Then
        Exit Sub ' Exit if VBA window is not open
    End If
    
    codeFolder = file_path(ThisWorkbook.path, "vba_export")
    create_folder (codeFolder)

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case vbext_ct_ClassModule
                FileName = file_path(codeFolder, comp.Name & ".cls")
                delete_file FileName
                comp.Export FileName
            Case vbext_ct_StdModule
                FileName = file_path(codeFolder, comp.Name & ".bas")
                delete_file FileName
                comp.Export FileName
            Case vbext_ct_MSForm
                FileName = file_path(codeFolder, comp.Name & ".frm")
                delete_file FileName
                comp.Export FileName
            Case vbext_ct_Document
                FileName = file_path(codeFolder, comp.Name & ".cls")
                delete_file FileName
                comp.Export FileName
        End Select
    Next
End Sub


Function CanAccessVBOM() As Boolean
'***************************************************************************
'Purpose: Check resgistry to see if we can access the VB object model
'Inputs
'Outputs:
'***************************************************************************
    Dim wsh As Object
    Dim str1 As String
    Dim AccessVBOM As Long

    Set wsh = CreateObject("WScript.Shell")
    str1 = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & _
        Application.Version & "\Excel\Security\AccessVBOM"
    On Error Resume Next
    AccessVBOM = wsh.RegRead(str1)
    Set wsh = Nothing
    CanAccessVBOM = (AccessVBOM = 1)
End Function
