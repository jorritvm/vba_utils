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


Function make_str_antares_safe(str As String) As String
'***************************************************************************
'Purpose: VBA port of the R function to translate unit name into a string safe to use in antares0
'Inputs:  any string
'Outputs: string that matches antares rules
'***************************************************************************

    str = LCase(str)
    str = Replace(str, "'", " ")
    str = Replace(str, "+", "")
    str = Replace(str, "\n", "")
    str = Replace(str, "ß", "s") ' estzet is a unicode character
    str = Trim(str)
    str = strip_accent(str)
    str = remove_punctuation(str)
    str = Replace(str, Space(2), Space(1))
    str = Replace(str, " ", "")
    
    make_str_antares_safe = str
End Function


Function strip_accent(str As String) As String
'***************************************************************************
'Purpose: replace accented characters by their non accented counterpart
'Inputs:  string with accented characters
'Outputs: modified string
'***************************************************************************

    Dim A As String * 1
    Dim B As String * 1
    Dim i As Integer
    Dim accented_chars As String
    Dim regular_chars As String

    '-- Add more chars to these 2 string as you want
    '-- You may have problem with unicode chars that has code > 255
    '-- such as some Vietnamese characters that are outside of ASCII code (0-255)
    accented_chars = "ŠšŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖÙÚÛÜİàáâãäåçèéêëìíîïğñòóôõöùúûüıÿ"
    regular_chars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
    
    For i = 1 To Len(accented_chars)
        A = Mid(accented_chars, i, 1)
        B = Mid(regular_chars, i, 1)
        str = Replace(str, A, B)
    Next
    
    strip_accent = str
End Function


Function remove_punctuation(str As String) As String
'***************************************************************************
'Purpose: removes all punctuation from a string
'Inputs:  string with punctuation characters
'Outputs: modified string
'***************************************************************************

    With CreateObject("VBScript.RegExp")
    .Pattern = "[^A-Z0-9 ]"
    .IgnoreCase = True
    .Global = True
    remove_punctuation = .Replace(str, "")
    End With
End Function

