Attribute VB_Name = "lib_string"
Option Explicit

Function flip_backslash(path) As String
'***************************************************************************
'Purpose: replace backslash by forward slash to be compatible with R
'Inputs:  full path
'Outputs: modified path
'***************************************************************************
    flip_backslash = Replace(path, "\", "/")
End Function
