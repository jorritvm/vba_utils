Attribute VB_Name = "lib_file_operations"
Option Explicit

Function file_path(ParamArray args() As Variant) As String
'***************************************************************************
'Purpose: creates a file path from pieces putting backslahses were needed
'Inputs:  argument array of file path pieces
'Outputs: proper filepath
'***************************************************************************
    Dim word As Variant
    
    file_path = ""
    
    For Each word In args
        word = CStr(word)
        If Left(word, 1) = "\" Then
            word = Right(word, Len(word) - 1)
        End If
        file_path = file_path & word
        If Right(word, 1) <> "\" Then
            file_path = file_path & "\"
        End If
    Next word
    
    file_path = Left(file_path, Len(file_path) - 1)
End Function


Function dirname(sPathFile As String) As String
'***************************************************************************
'Purpose: strip the filename off the end and return the path
'Inputs:  full filename path
'Outputs: only the directory path
'***************************************************************************
    Dim filesystem As New FileSystemObject
    dirname = filesystem.GetParentFolderName(sPathFile)
End Function


Function basename(path As String) As String
'***************************************************************************
'Purpose: strip the folder path and returns only the filename
'Inputs:  full filename path
'Outputs: basename
'***************************************************************************
    Dim filesystem As New FileSystemObject
    basename = filesystem.GetFileName(path)
End Function


Function file_exists(filepath As String) As Boolean
'***************************************************************************
'Purpose: check if a file exists, does not work on folders
'Inputs:  a filepath
'Outputs: true or false
'***************************************************************************
    If (filepath = "") Then
        file_exists = False
    Else
        Dim TestStr As String
        TestStr = ""
        On Error Resume Next
        TestStr = Dir(filepath)
        On Error GoTo 0
        If TestStr = "" Then
            file_exists = False
        Else
            file_exists = True
        End If
    End If
End Function


'***************************************************************************
'Purpose: deletes a file
'Inputs:  full path to file
'Outputs: -
'***************************************************************************
Sub delete_file(file_path As String)
    On Error Resume Next
    Kill file_path
End Sub


Sub create_folder(path As String)
'***************************************************************************
'Purpose: "create_folder" recusively creates folders until the destination
'         folder is reached
'Inputs:  path (string) folderpath to create
'Outputs: -
'***************************************************************************
    Dim elm As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    For Each elm In Split(path, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
End Sub


Function get_all_files_in_a_folder(ByVal folder_name As String)
'***************************************************************************
'Purpose: create an array containing the folder's content
'Inputs: a folder name as a string
'Outputs: a string array with 2 row, 1 is filename, 2 is path
'***************************************************************************
    Dim objFSO As Object
    Dim objFolder As Object
    Dim File As Object
    Dim i As Integer
    ReDim output_array(1 To 2, 1 To 1) As String
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.GetFolder(folder_name)
    i = 0
    'loops through each file in the directory and prints their names and path
    For Each File In objFolder.Files
        i = i + 1
        ReDim Preserve output_array(1 To 2, 1 To i)
        'print folder name
        output_array(1, i) = File.Name
        'print folder path
        output_array(2, i) = File.path
    Next File
    
    get_all_files_in_a_folder = output_array
End Function


Function get_all_subfolders_in_a_folder(ByVal folder_name As String)
'***************************************************************************
'Purpose: create an array containing the folder's subfolders
'Inputs: a folder name as a string
'Outputs: a string array with 2 row, 1 is subfoldername, 2 is path
'***************************************************************************
    Dim objFSO As Object
    Dim objFolder As Object
    Dim File As Object
    Dim i As Integer
    ReDim output_array(1 To 2, 1 To 1) As String
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object
    Set objFolder = objFSO.GetFolder(folder_name)
    i = 0
    'loops through each file in the directory and prints their names and path
    For Each File In objFolder.SubFolders
        i = i + 1
        ReDim Preserve output_array(1 To 2, 1 To i)
        'print folder name
        output_array(1, i) = File.Name
        'print folder path
        output_array(2, i) = File.path
    Next File
    
    get_all_subfolders_in_a_folder = output_array

End Function

Sub export_range_as_csv(filepath As String, range_to_export As Range)
'***************************************************************************
'Purpose: "export_range_as_csv" creates a CSV file containing the data in the provided Range
'
'Inputs:  range_to_export (Range) The Range containing the data to be exporterd
'         filename: (string) The name of the file to be created including extension
'         directory_path: (string) The path where the created CSV file should be saved
'Outputs: -
'***************************************************************************
    Dim export_wb As Workbook
    Dim directory_path As String
        
    directory_path = strip_filename(filepath)
    create_folder (directory_path)
    
    range_to_export.Copy
    
    Application.DisplayAlerts = False
    Set export_wb = Workbooks.Add
    export_wb.Sheets(1).Range("A1").PasteSpecial Paste:=xlPasteValues
    export_wb.SaveAs FileName:=filepath, FileFormat:=xlCSV
    export_wb.Close
    Application.DisplayAlerts = True
End Sub


Sub export_sheet_as_csv(filepath As String, sheet_name As String)
'***************************************************************************
'Purpose: "export_sheet_as_csv" exports a CSV file required by the R Tirole
'         output script
'Inputs:  filepath: (string) representing the full filepath to be written
'         sheet_name: (string) containing the sheetname to be exported
'Outputs: -
'***************************************************************************
    Dim MyFileName As String
    Dim TempWB As Workbook
    Dim directory_path As String

    directory_path = strip_filename(filepath)
    create_folder (directory_path)
    
    ThisWorkbook.Worksheets(sheet_name).UsedRange.Copy

    Application.DisplayAlerts = False
    Set TempWB = Application.Workbooks.Add(1)
    With TempWB.Sheets(1).Range("A1")
      .PasteSpecial xlPasteValues
      .PasteSpecial xlPasteFormats
    End With

    TempWB.SaveAs FileName:=filepath, FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    TempWB.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub


Sub copy_all_files_in_folder(source_folder_path As String, target_folder_path As String)
'***************************************************************************
'Purpose: This Subroutine copies all files in the source folder
'         to the target folder
'Inputs: source_folder_path (string) folderpath to copy
'        target_folder_path (string) folderpath to create
'Outputs: -
'***************************************************************************
    Dim MyFile2 As String
   
   MyFile2 = Dir(source_folder_path & "/*.*")
   Do While MyFile2 <> ""
        FileCopy source_folder_path & "/" & MyFile2, target_folder_path & "/" & MyFile2
        MyFile2 = Dir
    Loop
End Sub


Function get_newest_subfolder(ByRef path As String) As String
'***************************************************************************
'Purpose: "get_newest_subfolder" fetches the path to the latest antares ...
'         simulation output folder
'Inputs:  path: (string) path to output folder of antares model
'Outputs: (string) path to latest simulation output
'***************************************************************************
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oSubFldr As Object
    Dim strFolderPath As String
    Dim dteDate As Date

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(path)

    For Each oSubFldr In oFolder.SubFolders
        If oSubFldr.DateLastModified > dteDate Then
            dteDate = oSubFldr.DateLastModified
            strFolderPath = oSubFldr.path
        End If
    Next
    get_newest_subfolder = strFolderPath

    Set oSubFldr = Nothing
    Set oFolder = Nothing
    Set oFSO = Nothing
End Function


Sub export_worksheets_to_csv_files(Optional ByVal wb As String = "", _
                                   Optional ByVal path As String = "", _
                                   Optional ByVal do_screenupdating As Boolean = True, _
                                   Optional ByVal overwrite As Boolean = True)
'***************************************************************************
'Purpose: copy each of the workbooks' worksheets to a seperate csv file
'Inputs:  (optional) wb: name of an workbook to export - default: activeworkbook
'         (optional) path: path where the CSV files should be put - default: same as workbook
'         (optional) do_screenupdating: boolean to disable screenupdating during the routine execution - default: True
'         (optional) overwrite: boolean to set whether it should ask to overwrite existing files
'Outputs: CSV files (US style locale)
'***************************************************************************
    Dim ws_length As Integer
    Dim i As Integer
    Dim wb_source As Workbook, wb_new As Workbook
    Dim ws As String, fn As String, fp As String
    Dim displayalertstate As Boolean, screenupdateingstate As Boolean
    
    If (wb = "") Then
        Set wb_source = ActiveWorkbook
    Else
        Set wb_source = Workbooks(wb)
    End If
    
    If (do_screenupdating = False) Then
        screenupdateingstate = Application.screenupdating
        Application.screenupdating = False
    End If
    
    If (path = "") Then
        fp = wb_source.path
    Else
        fp = path
    End If
        
    If (overwrite) Then
        displayalertstate = Application.DisplayAlerts
        Application.DisplayAlerts = False
    End If
        
    ws_length = wb_source.Worksheets.Count
                
    For i = 1 To ws_length
        ws = LCase(wb_source.Worksheets(i).Name)
        
        'MsgBox ws
        Application.SheetsInNewWorkbook = 1
        Set wb_new = Application.Workbooks.Add
        wb_new.Sheets(1).Name = ws
        wb_source.Worksheets(i).Range("C12:AK8771").Copy wb_new.Worksheets(ws).Range("A1")

        fn = fp & ws & ".csv"
        wb_new.SaveAs FileName:=fn, FileFormat:=xlCSV, CreateBackup:=False
        wb_new.Close SaveChanges:=False
    Next i
    
    If (do_screenupdating = False) Then
        Application.screenupdating = screenupdateingstate
    End If
    
    If (overwrite) Then
        Application.DisplayAlerts = displayalertstate
    End If
    

End Sub


