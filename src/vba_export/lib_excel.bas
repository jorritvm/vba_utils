Attribute VB_Name = "lib_excel"
Option Explicit

Sub zoom_all_worksheets(zoomlevel)
'***************************************************************************
'Purpose: zooms all worksheets to the desired zoom level (
'Inputs:  zoomlevel as integer. E.g. 85 for 85% zoomlevel
'Outputs: -
'***************************************************************************
    Dim ws As Worksheet, wsinit As Worksheet
    Set wsinit = ActiveSheet
    Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        ActiveWindow.Zoom = zoomlevel
        Cells(1, 1).Select
    Next
    wsinit.Activate
    Application.ScreenUpdating = True
End Sub


Function convert_column_numeric2string(ByVal col_index As Integer) As String
'***************************************************************************
'Purpose: converts an excel numeric column number to alpha column number
'Inputs:  numeric value of the column
'Outputs: alpha value of the column
'***************************************************************************
    Dim col_string As String

    col_string = ThisWorkbook.Sheets(1).Cells(1, col_index).Address     'Address in $A$1 format
    col_string = Mid(col_string, 2, InStr(2, col_string, "$") - 2)  'Extracts just the letter
    
    convert_column_numeric2string = col_string
    
End Function
