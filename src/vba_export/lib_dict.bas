Attribute VB_Name = "lib_dict"
Option Explicit


Sub print_dict(Dict As scripting.dictionary)
'***************************************************************************
'Purpose: print content of dictionary to immediate window
'Inputs:  dict: a scripting.dictionary object containing string values
'Outputs: -
'***************************************************************************
    Dim key As Variant
    
    For Each key In Dict.Keys
    Debug.Print ("'" + CStr(key) + "': '" + CStr(Dict(key)) + "'")
    Next key
End Sub


Function deep_copy_dict(Dict)
'***************************************************************************
'Purpose: deep copy a scripting.dictionary object
'Inputs: input dictionary
'Outputs: output clone
'***************************************************************************
  Dim newDict
  Set newDict = CreateObject("Scripting.Dictionary")

  For Each key In Dict.Keys
    newDict.Add key, Dict(key)
  Next
  newDict.CompareMode = Dict.CompareMode

  Set deep_copy_dict = newDict
End Function


Function fetch_dict_from_range() As scripting.dictionary
'***************************************************************************
'Purpose: transforms all_param range into a dictionary so we can use
'         the params in our VBA functions
'Inputs: range object of 2 columns (key & value)
'Outputs: a dictionary containing the entries (key: value)
'***************************************************************************
    Dim all_param As Variant
    Dim s As scripting.dictionary
    
    Application.Calculate
    
    all_param = Application.Range("all_param")
    Set s = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(all_param)
        'stops when we encounter the first empty row
        If all_param(i, 1) = "" Then
            Exit For
        End If
        
        s.Add all_param(i, 1), all_param(i, 2)
    Next i
    
    Set fetch_dict_from_range = s
End Function


Function prepare_dict_for_R(s As scripting.dictionary)
'***************************************************************************
'Purpose: replace \ by / in all dict values
'         replace "" by "NULL" in all dict values
'Inputs: a scripting dict
'Outputs: a modified scriptin dict
'***************************************************************************
    Dim z As scripting.dictionary
    
    'deep copy the dict first
    Set z = deep_copy_dict(s)

    'now do some data manipulations before writing it to R
    For Each k In z.Keys
       ' replace backslashes
       z(k) = flip_backslash(z(k))
       
       ' replace empty strings
       v = z(k)
       If v = "" Then
           z.Remove k
           z.Add k, "NULL"
       End If
    Next k
    
    Set prepare_dict_for_R = z
End Function


Sub write_dict(s As scripting.dictionary, export_file)
'***************************************************************************
'Purpose: write the all params sheet to a csv file, everything in text format
'Inputs: s: a scripting.dictionary
'        export_file: a filepath
'Outputs: -
'***************************************************************************
    
    Dim fpfn As String
    Dim export_wb As Workbook
    
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    Set export_wb = Workbooks.Add
    i = 1
    For Each k In z.Keys
        export_wb.Sheets(1).Cells(i, 1) = k
        export_wb.Sheets(1).Cells(i, 2).NumberFormat = "@" 'make sure excel knows we are purely dealing with text
        export_wb.Sheets(1).Cells(i, 2) = z(k)
        i = i + 1
    Next k
    
    export_wb.SaveAs FileName:=export_file, FileFormat:=xlCSV
    export_wb.Close
    
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub



