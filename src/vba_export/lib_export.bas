Attribute VB_Name = "lib_export"
Option Explicit

Sub export_all_vba_code()
'***************************************************************************
'Purpose: exports the vba code in separate text files to be used in source control
'Inputs
'Outputs:
'***************************************************************************
' enable access to object model:
' 1. Start Microsoft Excel.
' 2. Open a workbook.
' 3. Click File and then Options.
' 4. In the navigation pane, select Trust Center.
' 5. Click Trust Center Settings....
' 6. In the navigation pane, select Macro Settings.
' 7. Ensure that Trust access to the VBA project object model is checked.
' 8. Click OK.

' enable required references
' 1. open VBE
' 2. click tools then references
' 3. Microsoft Visual Basic for Applications Extensibility 5.3
' 4. microsoft scripting runtime
    
    Dim comp As VBComponent
    Dim codeFolder As String
    Dim FileName As String
    
    If Not CanAccessVBOM Then
        MsgBox "Cannot access VBOM module - please activate this in Excel"
        Exit Sub ' Exit if access to VB object model is not allowed
    End If
    
'    If (ThisWorkbook.VBProject.VBE.ActiveWindow Is Nothing) Then
'        MsgBox "active window is nothing"
'        Exit Sub ' Exit if VBA window is not open
'    End If
    
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
    
    MsgBox "All VBA code has been exported as text files"
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
