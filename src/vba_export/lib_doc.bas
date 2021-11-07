Attribute VB_Name = "lib_doc"
Option Explicit
'code formatter:
'https://www.automateexcel.com/vba-code-indenter/

'***************************************************************************
'Purpose:
'Inputs
'Outputs:
'***************************************************************************


'-----------------------------------------------
' CHANGELOG
'-----------------------------------------------


'----------------
' INIT
'----------------


'----------------
' HOOKS
'----------------

'----------------
' HELPERS
'----------------


Sub create_vba_documentation()
'***************************************************************************
'Purpose: documents the workbook's VBA project into a 'doc' worksheet
'Inputs:  -
'Outputs: -
'***************************************************************************
    'declare
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBMod As VBIDE.CodeModule
       
    Dim doc_array(1 To 1000, 1 To 4) As String
    Dim i As Integer, j As Integer, i_array As Integer
    Dim docstring As String
    Dim looking_for_docstring As Boolean
    Dim doc_sheet As String
    Dim rng As Range
    
    'init
    Set VBProj = ThisWorkbook.VBProject
    
    doc_sheet = "doc" 'name of the workshet to put documentation
    
    doc_array(1, 1) = "Module"
    doc_array(1, 2) = "Routine name"
    doc_array(1, 3) = "Routine header"
    doc_array(1, 4) = "Docstring"
    
    i_array = 1
    
    looking_for_docstring = False
    
    ' loop excel objects and modules
    For Each VBComp In VBProj.VBComponents
        Debug.Print VBComp.Name
        Set VBMod = VBComp.CodeModule
        For i = 1 To VBMod.CountOfLines
            'Split the current line by space characters
            Dim code_line As String
            code_line = VBMod.Lines(i, 1)
            
            Dim code_line_words() As String
            code_line_words = Split(code_line, " ")
            
            ' gather docstring
            If looking_for_docstring Then
                If Left(code_line, 1) = "'" Then
                    If docstring = "" Then
                        docstring = code_line
                    Else
                        docstring = docstring + vbCrLf + code_line
                    End If
                Else
                    doc_array(i_array, 4) = docstring ' docstring
                    docstring = ""
                    looking_for_docstring = False
                End If
            End If
            
            ' loop through first three words of the codeline to look for 'sub' or 'function'
            For j = LBound(code_line_words) To WorksheetFunction.Min(3, UBound(code_line_words))
                If LCase(code_line_words(j)) = "sub" Or LCase(code_line_words(j)) = "function" Then
                    'This line is a subroutine declaration
                    If j + 1 <= UBound(code_line_words) Then
                        Dim routine_name As String
                        'find out if there's a parenthesis included in the declaration, if so get rid of it
                         If InStr(1, code_line_words(j + 1), "(") > 0 Then
                            routine_name = Left(code_line_words(j + 1), InStr(1, code_line_words(j + 1), "(") - 1)
                        Else
                            routine_name = code_line_words(j + 1)
                        End If
                        
                        i_array = i_array + 1
                        doc_array(i_array, 1) = VBComp.Name ' module name
                        doc_array(i_array, 2) = routine_name ' routine name
                        doc_array(i_array, 3) = code_line ' routine header
                        
                        looking_for_docstring = True
                    End If
                End If
            Next j

        Next i
    Next VBComp
    Set VBProj = Nothing
    Set VBMod = Nothing
    Set VBComp = Nothing
    Set VBProj = Nothing
    
    
    ' clear doc worksheet
    Sheets(doc_sheet).Activate
    Worksheets(doc_sheet).Cells.Clear
    
    ' get doc onto the doc worksheet
    Set rng = Worksheets(doc_sheet).Range("a1").Resize(UBound(doc_array, 1), UBound(doc_array, 2))
    rng.Value = doc_array
    
    ' format cell sizes
    Columns("A:D").EntireColumn.AutoFit
    Columns("C:C").ColumnWidth = WorksheetFunction.Min(Columns("C:C").ColumnWidth, 100) ' max widt = 100
    Columns("D:D").ColumnWidth = 100 ' fixed width  = 100
    Cells.Select
    Cells.EntireRow.AutoFit
    Cells.Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    'font
        Columns("A:D").Select
    With Selection.Font
        .Name = "Lucda consolas"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    ' header
    Range("A1:D1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    'add autofilter & sort
    Columns("A:D").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("doc").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("doc").AutoFilter.Sort.SortFields.Add key:=Range( _
        "A1:A1000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("doc").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' format as table
    'Columns("A:D").Select
    'Range("D1").Activate
    'ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$D"), , xlYes).Name = "doctable"
    'ActiveSheet.ListObjects("doctable").TableStyle = "TableStyleMedium9"
    
    ' reset selection
    Range("a1").Select
    MsgBox "Documentation updated."
   
End Sub
