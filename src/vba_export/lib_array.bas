Attribute VB_Name = "lib_array"
Option Explicit

Public Sub Array_Transpose(ByRef ar_Array As Variant)
'***********************************************************************************
'Sub Array_Transpose
'This sub transposes an array
'ar_Array : array to transpose
'***********************************************************************************
    Dim int_LB1, int_UB1, int_LB2, int_UB2, int_i, int_j As Integer

    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
    
    Dim ar_ArrayCopy As Variant
    ReDim ar_ArrayCopy(int_LB1 To int_UB1, int_LB2 To int_UB2)
    For int_i = int_LB1 To int_UB1
        For int_j = int_LB2 To int_UB2
            ar_ArrayCopy(int_i, int_j) = ar_Array(int_i, int_j)
        Next int_j
    Next int_i
        
    ReDim ar_Array(int_LB2 To int_UB2, int_LB1 To int_UB1)
    For int_i = int_LB1 To int_UB1
        For int_j = int_LB2 To int_UB2
            ar_Array(int_j, int_i) = ar_ArrayCopy(int_i, int_j)
        Next int_j
    Next int_i
    
    Erase ar_ArrayCopy
End Sub


Public Sub Array_Rename(ByRef ar_Array As Variant, ByVal ar_cols As Variant)
'***********************************************************************************
'Sub Array_Rename
'This sub renames the column header (header is always written in the first row of an array)
'ar_Array : array containing column header to rename
'ar_cols : array containing the pairwise columns with old & new name - Example : arCols=array(array("old_col1","newcol_1"),array("old_col2","newcol_2"),..)
'***********************************************************************************
Dim str_old_name, str_new_name As String
    Dim int_old_col As Integer
    Dim cols As Variant
    
    For Each cols In ar_cols
        str_old_name = cols(0)
        str_new_name = cols(1)
        int_old_col = find_row_col(ar_Array, str_old_name, 1)(2)
        ar_Array(1, int_old_col) = str_new_name
    Next cols
End Sub


Public Function Array_Slice(ByRef ar_Array As Variant, ByVal int_slice_num As Integer) As Variant
'***********************************************************************************
'Sub 'Array_Slice
'This fct returns the k-th slide or component of an array - Example : it will return x, y or z a for array(array("x(1)","y(1)","z(1)"),array("x(2)","y(2)","z(2)"))
'arArray : array containing data to slice
'slice_num : slice/component/dimension to extract
'***********************************************************************************
Dim int_LB, int_UB, int_i As Integer
    Dim ar_ArraySlice As Variant
    
    int_LB = LBound(ar_Array, 1)
    int_UB = UBound(ar_Array, 1)
    ReDim ar_ArraySlice(int_LB To int_UB)
    For int_i = int_LB To int_UB
        ar_ArraySlice(int_i) = ar_Array(int_i)(int_slice_num)
    Next int_i
    Array_Slice = ar_ArraySlice
End Function


Public Function Array_GetHeaders(ByRef ar_Array As Variant) As Variant
'***********************************************************************************
'Sint_UB Array_GetHeaders
'This fct will return the headers(i.e. the first row) of an array into a new 1D-array
'ar_Array : the array containg the headers name to extract (located in the first row)
'***********************************************************************************
Dim int_LB, int_UB, int_i As Integer
    
    int_LB = LBound(ar_Array, 2)
    int_UB = UBound(ar_Array, 2)
    ReDim Array_GetHeaders(int_LB To int_UB)
    For int_i = int_LB To int_UB
        Array_GetHeaders(int_i) = ar_Array(1, int_i)
    Next int_i
End Function

Public Sub Array_Keep(ByRef ar_Array As Variant, Optional ByVal int_col_nums = Empty, Optional ByVal ar_headers = Empty, _
Optional ByVal bo_order = False)
'***********************************************************************************
'Sub Array_Keep
'This sub keeps only the columns of a given array with header name in a specified list
'ar_Array : the array containing the ar_headers & the data
'int_col_nums : the column numbers to keep (only used if ar_headers is empty)
'ar_headers : the header names to keep
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2, int_Total, int_ToKeep, int_ToDrop As Integer
    Dim int_i, int_j, int_k As Integer
    Dim str_header As String
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
    
    int_Total = int_UB2 - int_LB2 + 1
    int_ToKeep = 1
    int_ToDrop = int_Total - int_ToKeep
    
    Dim ar_headersToDrop As Variant
    ReDim ar_headersToDrop(1 To int_ToDrop)
    
    int_k = 0
    For int_j = int_LB2 To int_UB2
        If (Not IsInArray(ar_Array(1, int_j), ar_headers)) Then
            If Not (IsInArray(ar_Array(1, int_j), ar_headersToDrop)) Then
                int_k = int_k + 1
                ar_headersToDrop(int_k) = ar_Array(1, int_j)
            End If
        End If
    Next int_j
    
    ReDim Preserve ar_headersToDrop(1 To int_k)
    
    For int_j = LBound(ar_headersToDrop) To UBound(ar_headersToDrop)
        str_header = ar_headersToDrop(int_j)
        Call Array_Drop(ar_Array, , str_header)
    Next int_j
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
    
    If bo_order Then
        Dim ar_ArrayCopy As Variant
        ReDim ar_ArrayCopy(int_LB1 To int_UB1, int_LB2 To int_UB2)
        For int_j = LBound(ar_headers) To UBound(ar_headers)
            str_header = ar_headers(int_j)
            int_col = find_row_col(ar_Array, str_header, 1)(2)
            For int_i = int_LB1 To int_UB1
                ar_ArrayCopy(int_i, int_j + LBound(ar_Array, 2)) = ar_Array(int_i, int_col)
            Next int_i
        Next int_j
        
        ReDim ar_Array(int_LB1 To int_UB1, int_LB2 To int_UB2)
        For int_i = int_LB1 To int_UB1
            For int_j = int_LB2 To int_UB2
                ar_Array(int_i, int_j) = ar_ArrayCopy(int_i, int_j)
            Next int_j
        Next int_i
    
    Erase ar_ArrayCopy
            
    End If

End Sub


Public Sub Array_Drop(ByRef ar_Array As Variant, Optional ByVal int_col_num = Empty, Optional ByVal str_header = Empty)
'***********************************************************************************
'Sub Array_Drop
'This sub drops one column of a given array based on a specified str_header name
'ar_Array : the array containing the str_headers & the data
'int_col_num : the column numbers to drop (only used if str_header is empty)
'str_header : the str_header name to drop
'***********************************************************************************
Dim int_LB1, int_LB2, int_UB1, int_UB2 As Integer
    Dim int_i, int_j, int_k As Integer
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
    
    Dim ar_ArrayCopy As Variant
    ReDim ar_ArrayCopy(int_LB1 To int_UB1, int_LB2 To (int_UB2 - 1))
    
    For int_i = int_LB1 To int_UB1
        int_k = int_LB2 - 1
        For int_j = int_LB2 To int_UB2
            If Not IsEmpty(str_header) Then
                If ar_Array(1, int_j) <> str_header Then
                    int_k = int_k + 1
                    ar_ArrayCopy(int_i, int_k) = ar_Array(int_i, int_j)
                End If
            ElseIf Not IsEmpty(int_col_num) Then
                If int_j <> int_col_num Then
                    int_k = int_k + 1
                    ar_ArrayCopy(int_i, int_k) = ar_Array(int_i, int_j)
                End If
            End If
        Next int_j
    Next int_i
    
    ReDim ar_Array(int_LB1 To int_UB1, int_LB2 To (int_UB2 - 1))
    For int_i = LBound(ar_Array, 1) To UBound(ar_Array, 1)
        For int_j = LBound(ar_Array, 2) To UBound(ar_Array, 2)
            ar_Array(int_i, int_j) = ar_ArrayCopy(int_i, int_j)
        Next int_j
    Next int_i
    
    Erase ar_ArrayCopy


End Sub


Public Sub Array_Append(ByRef ar_Array1 As Variant, ByRef ar_Array2 As Variant, Optional ByVal str_append_direction = "V", Optional ByVal bo_headers = False)
'***********************************************************************************
'Sub Array_Append
'This sub append a second array on a first array
'If vertical (resp. horizontal) append then the number of columns (resp.rows) of both arrays should be equal
'ar_Array1, : the base/first array
'ar_Array2, : the second array which will be appended on the base/first array
'str_append_direction : Vertical = "V", Horizontal ="H"
'bo_headers : If TRUE then the bo_headers of the first array will be preserved after appending
'***********************************************************************************
Dim int_LB1_array1, int_UB1_array1, int_LB2_array1, int_UB2_array1 As Integer
    Dim int_LB1_array2, int_UB1_array2, int_LB2_array2, int_UB2_array2 As Integer
    Dim int_i, int_j, int_lag As Integer
    Dim ar_ArrayCopy As Variant
    
    int_LB1_array1 = LBound(ar_Array1, 1)
    int_UB1_array1 = UBound(ar_Array1, 1)
    int_LB2_array1 = LBound(ar_Array1, 2)
    int_UB2_array1 = UBound(ar_Array1, 2)
    
    int_LB1_array2 = LBound(ar_Array2, 1)
    int_UB1_array2 = UBound(ar_Array2, 1)
    int_LB2_array2 = LBound(ar_Array2, 2)
    int_UB2_array2 = UBound(ar_Array2, 2)
        
    If str_append_direction = "V" Then
        If bo_headers Then int_lag = 1 Else int_lag = 0
        ReDim ar_ArrayCopy(int_LB1_array1 To (int_UB1_array1 + int_UB1_array2 + 1 - int_LB1_array2 - int_lag), int_LB2_array1 To int_UB2_array1)
        
        For int_j = int_LB2_array1 To int_UB2_array1
            For int_i = int_LB1_array1 To int_UB1_array1
                ar_ArrayCopy(int_i, int_j) = ar_Array1(int_i, int_j)
            Next int_i
        Next int_j
        
        For int_j = int_LB2_array2 To int_UB2_array2
            For int_i = int_LB1_array2 To (int_UB1_array2 - int_lag)
                ar_ArrayCopy(int_UB1_array1 + int_i - int_LB1_array2 + 1, int_j) = ar_Array2(int_i + int_lag, int_j)
            Next int_i
        Next int_j
    Else
       ReDim ar_ArrayCopy(int_LB1_array1 To int_UB1_array1, int_LB2_array1 To _
       int_UB2_array1 + int_UB2_array2 - int_LB2_array2 + 1)
    
        For int_j = int_LB2_array1 To int_UB2_array1
            For int_i = int_LB1_array1 To int_UB1_array1
                ar_ArrayCopy(int_i, int_j) = ar_Array1(int_i, int_j)
            Next int_i
        Next int_j
        
        For int_j = int_LB2_array2 To int_UB2_array2
            For int_i = int_LB1_array2 To int_UB1_array2
                ar_ArrayCopy(int_i, int_UB2_array1 + int_j - int_LB1_array2 + 1) = ar_Array2(int_i, int_j)
            Next int_i
        Next int_j
    End If
    
    ReDim ar_Array1(LBound(ar_ArrayCopy, 1) To UBound(ar_ArrayCopy, 1), LBound(ar_ArrayCopy, 2) To UBound(ar_ArrayCopy, 2))
    For int_i = LBound(ar_Array1, 1) To UBound(ar_Array1, 1)
        For int_j = LBound(ar_Array1, 2) To UBound(ar_Array1, 2)
            ar_Array1(int_i, int_j) = ar_ArrayCopy(int_i, int_j)
        Next int_j
    Next int_i
    Erase ar_ArrayCopy

End Sub


Public Sub Array_LeftJoin(ByRef ar_ArrayLeft As Variant, ByRef ar_ArrayRight As Variant, _
ar_Keys As Variant)
'***********************************************************************************
'Sub Array_LeftJoin
'This fct performs a left join between two arrays based on given key(s)
'Bith arrays should contains variable names/headers in their first row
'ar_Array1, : the left array
'ar_Array2, : the right array
'ar_Keys : array containing the pairwise keys of first array and second array - Example : ar_Keys=array(array("key1_array1","key1_array2"),array("key2_array1","key2_array2"),...)
'***********************************************************************************
Dim int_LB1_left, int_UB1_left, int_LB2_left, int_UB2_left As Integer
    Dim int_LB1_right, int_UB1_right, int_LB2_right, int_UB2_right As Integer
    Dim int_i, int_j, int_ii, int_ColLeft, int_ColRight As Integer
    Dim str_key As Variant
    Dim str_key_left, str_key_right As String
    
    ' Initialize array dimensions
    int_LB1_left = LBound(ar_ArrayLeft, 1)
    int_UB1_left = UBound(ar_ArrayLeft, 1)
    int_LB2_left = LBound(ar_ArrayLeft, 2)
    int_UB2_left = UBound(ar_ArrayLeft, 2)
    
    int_LB1_right = LBound(ar_ArrayRight, 1)
    int_UB1_right = UBound(ar_ArrayRight, 1)
    int_LB2_right = LBound(ar_ArrayRight, 2)
    int_UB2_right = UBound(ar_ArrayRight, 2)

    'Start with a copy of ar_ArrayLeft
    Dim ar_ArrayCopy As Variant
    ReDim ar_ArrayCopy(int_LB1_left To int_UB1_left, int_LB2_left To (int_UB2_left + int_UB2_right))
    
    For int_j = int_LB2_left To int_UB2_left
        For int_i = int_LB1_left To int_UB1_left
            ar_ArrayCopy(int_i, int_j) = ar_ArrayLeft(int_i, int_j)
        Next int_i
    Next int_j

    'Add headers
    For int_j = int_LB2_right To int_UB2_right
        ar_ArrayCopy(1, int_UB2_left + int_j) = ar_ArrayRight(1, int_j)
    Next int_j
    
    Dim ar_RightRow() As Variant
    ReDim ar_RightRow(int_LB1_right To int_UB1_right)
    
    For int_i = LBound(ar_ArrayCopy, 1) To UBound(ar_ArrayCopy, 1)
    
        For int_ii = LBound(ar_RightRow) To UBound(ar_RightRow)
            ar_RightRow(int_ii) = True
        Next int_ii
        
        For Each str_key In ar_Keys
            str_key_left = str_key(0)
            str_key_right = str_key(1)
            int_ColLeft = find_row_col(ar_ArrayCopy, str_key_left)(2)
            int_ColRight = find_row_col(ar_ArrayRight, str_key_right)(2)
            For int_ii = int_LB1_right To int_UB1_right
                If ar_ArrayLeft(int_i, int_ColLeft) <> ar_ArrayRight(int_ii, int_ColRight) Then ar_RightRow(int_ii) = False
            Next int_ii
        Next str_key
        
        For int_ii = int_LB1_right To int_UB1_right
            If ar_RightRow(int_ii) Then
                For int_j = int_LB2_right To int_UB2_right
                    ar_ArrayCopy(int_i, int_UB2_left + int_j) = ar_ArrayRight(int_ii, int_j)
                Next int_j
                Exit For
            End If
        Next int_ii
    Next int_i
    
    ReDim ar_ArrayLeft(LBound(ar_ArrayCopy, 1) To UBound(ar_ArrayCopy, 1), LBound(ar_ArrayCopy, 2) To UBound(ar_ArrayCopy, 2))
    For int_i = LBound(ar_ArrayLeft, 1) To UBound(ar_ArrayLeft, 1)
        For int_j = LBound(ar_ArrayLeft, 2) To UBound(ar_ArrayLeft, 2)
            ar_ArrayLeft(int_i, int_j) = ar_ArrayCopy(int_i, int_j)
        Next int_j
    Next int_i
    Erase ar_ArrayCopy

End Sub


Public Sub Array_LoadConstant(ByRef ar_Array As Variant, str_const_value, Optional ByVal int_start_row = Empty, Optional ByVal int_col = Empty)
'***********************************************************************************
'Sub Array_LoadConstant
'This sub writes a constant value in a given column of an array
'ar_Array : array which will contain the data
'str_const_value : constant to be written
'int_start_row : first row starting the writing
'int_col : column number in ar_Array - i.e. where to put the data
'***********************************************************************************
Dim int_LB, int_UB As Integer
    Dim int_i As Integer
    
    int_LB = LBound(ar_Array, 1)
    int_UB = UBound(ar_Array, 1)
    
    If IsEmpty(int_start_row) Then int_start_row = int_LB
    
    If IsEmpty(int_col) Then
        For int_i = int_start_row To int_UB
            ar_Array(int_i) = str_const_value
        Next int_i
    Else
        For int_i = int_start_row To int_UB
            ar_Array(int_i, int_col) = str_const_value
        Next int_i
    End If

End Sub


Public Sub Array_LoadRow(ByRef ar_Array As Variant, ar_data As Variant, Optional ByVal int_row_array As Integer, Optional ByVal int_row_data As Integer, _
Optional ar_exitValues As Variant = Empty, Optional ByVal int_start_col_data = Empty, Optional ByVal int_exit_row = Empty)
'***********************************************************************************
'Sub Array_LoadRow
'This sub loads one row of data from a given array into a second one with specific conditions
'ar_Array : array which will contain the loaded row
'ar_Data : array containing the row-data
'int_row_array : row number in ar_Array - where to put the data
'int_row_data : row number in ar_Data - where to find the data to be loaded
'ar_exitValues : array containing the characters triggering the exit of the loading process - Example :  ar_exitValues=Array("", Empty)
'int_start_col_data : first row used in ar_Data for loading the data
'int_exit_row : row used in ar_Data for stopping the loading process
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2, int_LB_data, int_UB_data, int_j As Integer
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB_data = LBound(ar_data, 2)
    int_UB_data = UBound(ar_data, 2)
        
    If int_row_array = 0 Then
        If IsEmpty(int_start_col_data) Then int_start_col_data = int_LB_data
        For int_j = int_start_col_data To int_UB_data
            If Not IsEmpty(int_exit_row) Then If IsInArray(ar_data(int_exit_row, int_j), ar_exitValues) Then Exit For
            ar_Array(int_LB1 + int_j - int_start_col_data) = ar_data(int_row_data, int_j)
        Next int_j
    Else
        int_LB2 = LBound(ar_Array, 2)
        int_UB2 = UBound(ar_Array, 2)
        If IsEmpty(int_start_col_data) Then int_start_col_data = int_LB_data
        For int_j = int_start_col_data To int_UB_data
            If Not IsEmpty(int_exit_row) Then If IsInArray(ar_data(int_exit_row, int_j), ar_exitValues) Then Exit For
            ar_Array(int_row_array, int_LB2 + int_j - int_start_col_data) = ar_data(int_row_data, int_j)
        Next int_j
    End If
End Sub


Public Sub Array_LoadColumnFromRow(ByRef ar_Array As Variant, ar_data As Variant, ByVal int_col_array As Integer, ByVal int_row_data As Integer, _
ar_exitValues As Variant, Optional ByVal int_start_col_data = Empty, Optional ByVal int_exit_row = Empty)
'***********************************************************************************
'Sub Array_LoadColumnFromRow
'This sub loads one row of data from a given array into a column of a second one with specific conditions
'ar_Array : array which will contain the loaded column
'ar_Data : array containing the row-data
'int_col_array : column number in ar_Array - where to put the data
'int_row_data : row number in ar_Data - where to find the data to be loaded
'ar_exitValues : array containing the characters triggering the exit of the loading process - Example :  ar_exitValues=Array("", Empty)
'int_start_col_data : first col used in ar_Data for loading the data
'int_exit_row : row used in ar_Data for stopping the loading process
'***********************************************************************************
Dim int_LB, int_UB, int_i As Integer
    
    int_LB = LBound(ar_data, 2)
    int_UB = UBound(ar_data, 2)
    
    If IsEmpty(int_start_col_data) Then int_start_col_data = int_LB
    
    For int_i = int_start_col_data To int_UB
        If Not IsEmpty(int_exit_row) Then If IsInArray(ar_data(int_exit_row, int_i), ar_exitValues) Then Exit For
        ar_Array(LBound(ar_Array, 1) + int_i - int_start_col_data, int_col_array) = ar_data(int_row_data, int_i)
    Next int_i
End Sub


Public Sub Array_LoadColumn(ByRef ar_Array As Variant, ByVal ar_data As Variant, ByVal int_col_array As Integer, ByVal int_col_data As Integer, _
ar_exitValues As Variant, Optional ByVal int_start_row_data = Empty, Optional ByVal int_exit_col = Empty, Optional ByVal bo_transpose = False)
'***********************************************************************************
'Sub Array_LoadColumn
'This sub loads one column of data from a given array into a second one with specific conditions
'ar_Array : array which will contain the loaded column
'ar_Data : array containing the column-data
'int_col_array : column number in ar_Array - where to put the data
'int_col_data : column number in ar_Data - where to find the data to be loaded
'ar_exitValues : array containing the characters triggering the exit of the loading process - Example :  ar_exitValues=Array("", Empty)
'int_start_row_data : first row used in ar_Data for loading the data
'int_exit_col : column used in ar_Data for stopping the loading process
'bo_transpose : if TRUE then ar_Data is translated before the loading
'***********************************************************************************
Dim int_LB, int_UB, int_i As Integer
    
    int_LB = LBound(ar_data, 1)
    int_UB = UBound(ar_data, 1)

    If IsEmpty(int_start_row_data) Then int_start_row_data = int_LB
    
    If bo_transpose Then Call Array_Transpose(ar_data)
    
    For int_i = int_start_row_data To int_UB
        If Not IsEmpty(int_exit_col) Then If IsInArray(ar_data(int_i, int_exit_col), ar_exitValues) Then Exit For
        ar_Array(LBound(ar_Array, 1) + int_i - int_start_row_data, int_col_array) = ar_data(int_i, int_col_data)
    Next int_i
End Sub


Public Sub Array_LoadData(ByRef ar_Array As Variant, ar_data As Variant, _
ar_exitValues As Variant, Optional ByVal int_exit_row = Empty, Optional ByVal int_exit_col = Empty, _
Optional ByVal int_start_row_data = Empty, Optional ByVal int_start_col_data = Empty, Optional ByVal int_header_row = Empty)
'***********************************************************************************
'Sub Array_LoadData
'This sub loads all the data from a given array into a second one with specific conditions
'ar_Array : array which will contain the loaded data
'ar_Data : array containing the data
'ar_exitValues : array containing the characters triggering the exit of the loading process - Example :  ar_exitValues=Array("", Empty)
'int_exit_row : row used in ar_Data for stopping the loading process
'int_exit_col : column used in ar_Data for stopping the loading process
'int_start_row_data : first row used in ar_Data for loading the data
'int_start_col_data : first column used in ar_Data for loading the data
'int_header_row : row used in ar_Data for defining the header loaded in ar_Array
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_LB1_data, int_UB1_data, int_LB2_data, int_UB2_data As Integer
    Dim int_i, int_j, int_lag As Integer
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
    
    int_LB1_data = LBound(ar_data, 1)
    int_UB1_data = UBound(ar_data, 1)
    int_LB2_data = LBound(ar_data, 2)
    int_UB2_data = UBound(ar_data, 2)

    If IsEmpty(int_start_row_data) Then int_start_row_data = int_LB1_data
    If IsEmpty(int_start_col_data) Then int_start_col_data = int_LB2_data
    
    'Detect the number of columns in the data set
    If Not IsEmpty(int_exit_row) Then
        For int_j = int_start_col_data To int_UB2_data
            If IsInArray(ar_data(int_exit_row, int_j), ar_exitValues) Then Exit For
        Next int_j
        int_end_col_data = int_j - 1
    Else
        int_end_col_data = int_start_col_data + int_UB2 - 1
    End If
    
    'Detect the number of rows in the data set
    If Not IsEmpty(int_exit_col) Then
        For int_i = int_start_row_data To int_UB1_data
            If IsInArray(ar_data(int_i, int_exit_col), ar_exitValues) Then Exit For
        Next int_i
        int_end_row_data = int_i - 1
    Else
        int_end_row_data = int_start_row_data + int_UB1 - 1
    End If
    
    If (int_end_col_data < int_start_col_data) Or (int_end_row_data < int_start_row_data) Then Exit Sub
    
    If Not IsEmpty(int_header_row) Then
        For int_j = int_start_col_data To int_end_col_data
            ar_Array(int_LB1_data, int_LB2_data + int_j - int_start_col_data) = ar_data(int_header_row, int_j)
        Next int_j
        int_lag = 1
    Else
        int_lag = 0
    End If
        
    For int_i = int_start_row_data To int_end_row_data
        For int_j = int_start_col_data To (int_end_col_data)
            ar_Array(int_LB1_data + int_lag + int_i - int_start_row_data, int_LB2_data + int_j - int_start_col_data) = ar_data(int_i, int_j)
        Next int_j
    Next int_i
End Sub


Public Sub Array_Copy(ByRef ar_ArrayToCopy As Variant, ByRef ar_Array As Variant)
'***********************************************************************************
'Sub Array_Copy
'This sub copy an array to another one
'ar_ArrayToCopy : array to be copied
'ar_Array : array containing the copied data
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_i, int_j As Integer
    Dim ar_ArrayCopy As Variant
    
    int_LB1 = LBound(ar_ArrayToCopy, 1)
    int_UB1 = UBound(ar_ArrayToCopy, 1)
    int_LB2 = LBound(ar_ArrayToCopy, 2)
    int_UB2 = UBound(ar_ArrayToCopy, 2)
    
    ReDim ar_ArrayCopy(int_LB1 To int_UB1, int_LB2 To int_UB2)
    For int_i = int_LB1 To int_UB1
        For int_j = int_LB2 To int_UB2
            ar_ArrayCopy(int_i, int_j) = ar_ArrayToCopy(int_i, int_j)
        Next int_j
    Next int_i
    
    ReDim ar_Array(int_LB1 To int_UB1, int_LB2 To int_UB2)
    For int_i = int_LB1 To int_UB1
        For int_j = int_LB2 To int_UB2
            ar_Array(int_i, int_j) = ar_ArrayCopy(int_i, int_j)
        Next int_j
    Next int_i
    Erase ar_ArrayCopy
End Sub


Public Sub Array_Cut1D(ByRef ar_Array As Variant, ar_exitValues As Variant)
'***********************************************************************************
'Sub Array_Cut1D
'This sub reduces an 1D array to the first non-empty block of elements
'ar_Array : array containing the elements and subject to reduction
'ar_exitValues : array containing the characters triggering the cutting process - Example :  ExitValues=Array("", Empty)
'***********************************************************************************
Dim int_LB, int_UB As Integer
    Dim int_i As Integer
    
    int_LB = LBound(ar_Array)
    int_UB = UBound(ar_Array)
    
    For int_i = int_LB To int_UB
        If IsInArray(ar_Array(int_i), ar_exitValues) Then Exit For
    Next int_i
    
    Dim ar_ArrayCopy As Variant
    ReDim ar_ArrayCopy(int_LB To (int_i - 1))
    
    For int_i = LBound(ar_ArrayCopy, 1) To UBound(ar_ArrayCopy, 1)
        ar_ArrayCopy(int_i) = ar_Array(int_i)
    Next int_i
    
    ReDim ar_Array(LBound(ar_ArrayCopy, 1) To UBound(ar_ArrayCopy, 1))
    For int_i = LBound(ar_Array, 1) To UBound(ar_Array, 1)
        ar_Array(int_i) = ar_ArrayCopy(int_i)
    Next int_i
    Erase ar_ArrayCopy
End Sub


Public Sub Array_Cut(ByRef ar_Array As Variant, ar_exitValues As Variant)
'***********************************************************************************
'Sub Array_Cut
'This sub reduces an array to the first non-empty block of elements
'ar_Array : array containing the elements and subject to reduction
'ar_exitValues : array containing the characters triggering the cutting process - Example :  ExitValues=Array("", Empty)
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_i As Integer
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
    
    For int_i = int_LB1 To int_UB1
        If IsInArray(ar_Array(int_i), ar_exitValues) Then Exit For
    Next int_i
    
    Dim ar_ArrayCopy As Variant
    ReDim ar_ArrayCopy(int_LB1 To (int_i - 1))
    
    For int_i = LBound(ar_ArrayCopy, 1) To UBound(ar_ArrayCopy, 1)
        ar_ArrayCopy(int_i) = ar_Array(int_i)
    Next int_i
    
    ReDim ar_Array(LBound(ar_ArrayCopy, 1) To UBound(ar_ArrayCopy, 1))
    For int_i = LBound(ar_Array, 1) To UBound(ar_Array, 1)
        ar_Array(int_i) = ar_ArrayCopy(int_i)
    Next int_i
    Erase ar_ArrayCopy
End Sub


Public Sub Array_CutAtCol(ByRef ar_Array As Variant, int_col As Integer)
'***********************************************************************************
'This sub reduces an array to a given number of columns
'ar_Array : array containing the data and subject to column reduction
'int_col : number or colums to keep or starting cut column number
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_i, int_j As Integer
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)

    Dim ar_ArrayCopy As Variant
    ReDim ar_ArrayCopy(int_LB1 To int_UB1, int_LB2 To int_col)
    int_UB2 = UBound(ar_ArrayCopy, 2)
    
    For int_j = int_LB2 To int_UB2
        For int_i = int_LB1 To int_UB1
            ar_ArrayCopy(int_i, int_j) = ar_Array(int_i, int_j)
        Next int_i
    Next int_j
    
    ReDim ar_Array(int_LB1 To int_UB1, int_LB2 To int_UB2)
    For int_i = int_LB1 To int_UB1
        For int_j = int_LB2 To int_UB2
            ar_Array(int_i, int_j) = ar_ArrayCopy(int_i, int_j)
        Next int_j
    Next int_i
    Erase ar_ArrayCopy
End Sub

Public Sub Array_CutAtLastCol(ByRef ar_Array As Variant, ar_exitValues As Variant, int_row As Integer)
'***********************************************************************************
'Sub Array_CutAtLastCol
'This sub reduces an array to the first non-empty block of columns
'ar_Array : array containing the data and subject to column reduction
'ExitValues : array containing the characters triggering the cutting process - Example :  ExitValues=Array("", Empty)
'int_row : int_row number used in ar_Array for testing the cutting process
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_i, int_j As Integer
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
    
    For int_j = int_LB2 To int_UB2
        If IsInArray(ar_Array(int_row, int_j), ar_exitValues) Then Exit For
    Next int_j
    
    Dim ar_ArrayCopy As Variant
    If int_j = int_LB2 Then int_j = int_j + 1
    ReDim ar_ArrayCopy(int_LB1 To int_UB1, int_LB2 To (int_j - 1))
    int_UB2 = UBound(ar_ArrayCopy, 2)
    
    For int_j = int_LB2 To int_UB2
        For int_i = int_LB1 To int_UB1
            ar_ArrayCopy(int_i, int_j) = ar_Array(int_i, int_j)
        Next int_i
    Next int_j
    ReDim ar_Array(int_LB1 To int_UB1, int_LB2 To int_UB2)
    For int_i = int_LB1 To int_UB1
        For int_j = int_LB2 To int_UB2
            ar_Array(int_i, int_j) = ar_ArrayCopy(int_i, int_j)
        Next int_j
    Next int_i
    Erase ar_ArrayCopy
End Sub


Public Sub Array_CutAtRow(ByRef ar_Array As Variant, ByVal int_row As Integer)
'***********************************************************************************
'Sub Array_CutAtRow
'This sub reduces an array to a given number of int_rows
'ar_Array : array containing the data and subject to int_row reduction
'int_row : number or rows to keep or starting cut column number
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_i, int_j As Integer
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
        
    Dim ar_ArrayCopy As Variant
    ReDim ar_ArrayCopy(int_LB1 To int_row, int_LB2 To int_UB2)
    int_UB1 = UBound(ar_ArrayCopy, 1)
    
    For int_i = int_LB1 To int_UB1
        For int_j = int_LB2 To int_UB2
            ar_ArrayCopy(int_i, int_j) = ar_Array(int_i, int_j)
        Next int_j
    Next int_i
    
    ReDim ar_Array(int_LB1 To int_UB1, int_LB2 To int_UB2)
    For int_i = int_LB1 To int_UB1
        For int_j = int_LB2 To int_UB2
            ar_Array(int_i, int_j) = ar_ArrayCopy(int_i, int_j)
        Next int_j
    Next int_i
    Erase ar_ArrayCopy
End Sub


Public Sub Array_RemoveRowCol(ByRef ar_Array As Variant, ByVal int_row_num As Integer, Optional ByVal bo_row = True)
'***********************************************************************************
'Sub Array_RemoveRowCol
'This sub removes one given row from an array
'ar_Array : array containing the data and subject to row reduction
'int_row_num : the row/col number to remove
'bo_row : if TRUE then it removes row (otherwise column)
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_row, int_col As Integer
    Dim ar_ArrayCopy As Variant
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
    
    ReDim ar_ArrayCopy(int_LB1 To int_UB1, int_LB2 To int_UB2)
    For int_row = int_LB1 To int_UB1
        For int_col = int_LB2 To int_UB2
           ar_ArrayCopy(int_row, int_col) = ar_Array(int_row, int_col)
        Next int_col
    Next int_row
    
    If bo_row Then int_idx = 1 Else int_idx = 0
    ReDim ar_Array(int_LB1 To (int_UB1 - int_idx), int_LB2 To (int_UB2 - (1 - int_idx)))
    For int_row = int_LB1 To int_UB1
        For int_col = int_LB2 To int_UB2
           If bo_row Then
                If int_row > int_row_num Then int_lag = 1 Else int_lag = 0
                If int_row <> int_row_num Then ar_Array(int_row - int_lag, int_col) = ar_ArrayCopy(int_row, int_col)
            Else
                If int_col > int_row_num Then int_lag = 1 Else int_lag = 0
                If int_col <> int_row_num Then ar_Array(int_row, int_col - int_lag) = ar_ArrayCopy(int_row, int_col)
            End If
        Next int_col
    Next int_row
End Sub


Public Sub Array_RemoveRowIf(ByRef ar_Array As Variant, ByVal int_col_num As Integer, ar_exitValues As Variant)
'***********************************************************************************
'Sub Array_RemoveRowIf
'This sub removes one given row from an array if the value within a given colmun is in exit values
'ar_Array : array containing the data and subject to row reduction
'int_col_num : the row number to remove
'ar_exitValues : array containing the characters triggering the cutting process
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim ar_ArrayCopy As Variant
    Dim int_row, int_col As Integer
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
    
    ReDim ar_ArrayCopy(int_LB1 To int_UB1, int_LB2 To int_UB2)
    For int_row = int_LB1 To int_UB1
        For int_col = int_LB2 To int_UB2
           ar_ArrayCopy(int_row, int_col) = ar_Array(int_row, int_col)
        Next int_col
    Next int_row
    
    
    For int_row = int_UB1 To int_LB1 Step -1
        If IsInArray(ar_ArrayCopy(int_row, int_col_num), ar_exitValues) Then Call Array_RemoveRowCol(ar_Array, int_row)
    Next int_row
End Sub


Public Sub Array_RemoveEmptyCols(ByRef ar_Array As Variant, Optional ByVal bo_header = Empty)
'***********************************************************************************
'Sub Array_RemoveEmptyCols
'This sub reduces an array to non-empty columns
'ar_Array : array containing the data and subject to column reduction
'bo_header : if TRUE then first row contains bo_headers
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_k, int_row, int_col, int_lag As Integer
    
    Dim ar_ArrayCopy As Variant
    Dim ar_flag_lag_empty_col() As Boolean
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
    
    ReDim ar_flag_lag_empty_col(int_LB2 To int_UB2)
    
    If IsEmpty(bo_header) Then int_lag = 0 Else int_lag = 1
    
    'Init
    For int_col = int_LB2 To int_UB2
        ar_flag_lag_empty_col(int_col) = True
    Next int_col
    
    'Identify the empty cols
    For int_col = int_LB2 To int_UB2
        For int_row = (int_LB1 + int_lag) To int_UB1
            If ar_Array(int_row, int_col) <> "" And Not (IsEmpty(ar_Array(int_row, int_col))) Then ar_flag_lag_empty_col(int_col) = False
        Next int_row
    Next int_col
    
    'Number of non-empty cols
    int_k = 0
    For int_col = int_LB2 To int_UB2
        If (ar_flag_lag_empty_col(int_col)) = False Then int_k = int_k + 1
    Next int_col
    
    ' Fill an array with non-empty cols
    If int_k = 0 Then
        Exit Sub
    Else
        ReDim ar_ArrayCopy(int_LB1 To int_UB1, int_LB2 To (int_LB2 + int_k - 1))
        int_k = 0
        For int_col = int_LB2 To int_UB2
            If (ar_flag_lag_empty_col(int_col)) = False Then
                int_k = int_k + 1
                For int_row = int_LB1 To int_UB1
                    ar_ArrayCopy(int_row, int_LB2 + int_k - 1) = ar_Array(int_row, int_col)
                Next int_row
            End If
        Next int_col
        
        'Copy to ar_Array
        ReDim ar_Array(int_LB1 To int_UB1, int_LB2 To UBound(ar_ArrayCopy, 2))
        int_UB2 = UBound(ar_Array, 2)
        For int_col = int_LB2 To int_UB2
            For int_row = int_LB1 To int_UB1
                ar_Array(int_row, int_col) = ar_ArrayCopy(int_row, int_col)
            Next int_row
        Next int_col
    End If
End Sub


Public Sub Array_CutAtLastRow(ByRef ar_Array As Variant, ar_exitValues As Variant, ByVal int_col As Integer)
'***********************************************************************************
'Sub Array_CutAtLastRow
'This sub reduces an array to the first non-empty block of rows
'ar_Array : array containing the data and subject to column reduction
'ar_exitValues : array containing the characters triggering the cutting process - Example :  ExitValues=Array("", Empty)
'int_col : column number used in ar_Array for testing the cutting process
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_i, int_j As Integer
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
        
    For int_i = int_LB1 To int_UB1
        If IsInArray(ar_Array(int_i, int_col), ar_exitValues) Then Exit For
    Next int_i
    
    If int_i = int_LB1 Then int_i = int_LB1 + 1

    Dim ar_ArrayCopy As Variant
    ReDim ar_ArrayCopy(int_LB1 To (int_i - 1), int_LB2 To int_UB2)
    int_UB1 = UBound(ar_ArrayCopy, 1)
    
    For int_i = LBound(ar_ArrayCopy, 1) To UBound(ar_ArrayCopy, 1)
        For int_j = LBound(ar_ArrayCopy, 2) To UBound(ar_ArrayCopy, 2)
            ar_ArrayCopy(int_i, int_j) = ar_Array(int_i, int_j)
        Next int_j
    Next int_i
    
    ReDim ar_Array(LBound(ar_ArrayCopy, 1) To UBound(ar_ArrayCopy, 1), LBound(ar_ArrayCopy, 2) To UBound(ar_ArrayCopy, 2))
    For int_i = LBound(ar_Array, 1) To UBound(ar_Array, 1)
        For int_j = LBound(ar_Array, 2) To UBound(ar_Array, 2)
            ar_Array(int_i, int_j) = ar_ArrayCopy(int_i, int_j)
        Next int_j
    Next int_i
    Erase ar_ArrayCopy
End Sub


Public Sub Array_AddHeader(ByRef ar_Array As Variant, ar_headers As Variant, Optional ByVal int_col_headers = Empty, Optional ByVal bo_shift = False)
'***********************************************************************************
'Sub Array_AddHeader
'This sub add headers/variable name as first row to a given array
'ar_Array : array containing the data
'ar_headers : array containing the headers name
'int_col_headers:  Print
'bo_shift : if TRUE then first row is bo_shifted before adding the headers as first row
'***********************************************************************************
Dim int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_i, int_j As Integer
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
    Dim ar_ArrayCopy As Variant
    
    ReDim ar_ArrayCopy(int_LB1 To (int_UB1 + 1), int_LB2 To int_UB2)
    
    If bo_shift Then
        For int_i = int_LB1 To int_UB1
            For int_j = int_LB2 To int_UB2
                ar_ArrayCopy(int_i + 1, int_j) = ar_Array(int_i, int_j)
            Next int_j
        Next int_i
        int_UB1 = UBound(ar_ArrayCopy, 1)
        ReDim ar_Array(int_LB1 To int_UB1, int_LB2 To int_UB2)
        For int_i = int_LB1 To int_UB1
            For int_j = int_LB2 To int_UB2
                ar_Array(int_i, int_j) = ar_ArrayCopy(int_i, int_j)
            Next int_j
        Next int_i
        Erase ar_ArrayCopy
    End If
        
    If IsEmpty(int_col_headers) Then
        For int_j = LBound(ar_headers, 1) To UBound(ar_headers, 1)
            If int_LB2 + int_j <= int_UB2 Then ar_Array(int_LB1, int_LB2 + int_j - LBound(ar_headers, 1)) = ar_headers(int_j)
        Next int_j
    Else
        For int_j = LBound(ar_headers, 1) To UBound(ar_headers, 1)
            If int_LB2 + int_j <= int_UB2 Then ar_Array(int_LB1, int_LB2 + int_j - LBound(ar_headers, 1)) = ar_headers(int_j)(int_col_headers)
        Next int_j
    End If
End Sub


Function find_row_col(ar_Array As Variant, ByVal str_text_to_find As String, Optional ByVal int_row = Empty, Optional ByVal int_col = Empty, Optional ByVal int_start_row = 1, Optional ByVal int_start_col = 1) As Variant
'***********************************************************************************
' Function find_row_col
' This fct the position of a given string within an array : row & column
' ar_Array : Input array
' str_text_to_find : String for which position (row,col) is searched
' int_row : if provided then the string position is only searched on this row (then only the column is searched)
' int_column : if provided then the string position is only searched on this column (then only the row is searched)
' int_start_row : starting row for string position searching
' int_start_col : starting col for string position searching
' bo1D : if TRUE then consider the arrays as 1-dimensional
'***********************************************************************************
'Declare
    Dim int_irow, int_icol As Integer
    Dim bo_flg As Boolean
    Dim ar_row_col(1 To 2) As Integer

    'Search the string according the row & column user specifications
    If IsEmpty(int_row) And IsEmpty(int_col) Then
        bo_flg = False
        For int_irow = int_start_row To UBound(ar_Array, 1)
            For int_icol = int_start_col To UBound(ar_Array, 2)
                If LCase(ar_Array(int_irow, int_icol)) = LCase(str_text_to_find) Then
                    bo_flg = True
                    Exit For
                End If
            Next int_icol
            If bo_flg = True Then Exit For
        Next int_irow
    ElseIf IsEmpty(int_row) Then
        For int_irow = int_start_row To UBound(ar_Array, 1)
            If LCase(ar_Array(int_irow, int_col)) = LCase(str_text_to_find) Then Exit For
            
        Next int_irow
    ElseIf IsEmpty(int_col) Then
        For int_icol = int_start_col To UBound(ar_Array, 2)
            If LCase(ar_Array(int_row, int_icol)) = LCase(str_text_to_find) Then Exit For
        Next int_icol
    Else
        int_irow = int_row
        int_icol = int_col
    End If

    'Return the row & column position in a 2-element array
    ar_row_col(1) = int_irow
    ar_row_col(2) = int_icol
    find_row_col = ar_row_col()
End Function


Public Function IsInArray(str_text_to_find As Variant, ar_Array As Variant, Optional ByVal int_col As Integer = 0, Optional ByVal int_row As Integer = 0) As Boolean
'***********************************************************************************
' Function IsInArray
' This fct checks whether a string belongs to an ar_Arrayay/list
' str_text_to_find : string to be found in input ar_Arrayay
' ar_Array : input ar_Arrayay
' int_col : if provided then the string sreaching will be performed only in the given column
' int_row : if provided then the string sreaching will be performed only in the given row
'***********************************************************************************
'Declare
    Dim int_i, int_LB, int_UB As Integer
    Dim str
    
    'Search the string according the row & column user specifications
    If int_row = 0 Then
        int_LB = LBound(ar_Array)
        int_UB = UBound(ar_Array)
    Else
        int_LB = LBound(ar_Array, 2)
        int_UB = UBound(ar_Array, 2)
    End If
    For int_i = int_LB To int_UB
        If int_col = 0 Then
            If int_row = 0 Then
                str = ar_Array(int_i)
            Else
                str = ar_Array(int_row, int_i)
            End If
        Else
            str = ar_Array(int_i, int_col)
        End If
        If str = str_text_to_find Then
            IsInArray = True
            Exit Function
        End If
    Next int_i
    IsInArray = False
End Function


Public Function Array_Max(ar_Array As Variant, Optional ByVal int_col As Integer = 0, Optional ByVal int_row As Integer = 0, Optional ByVal int_start_col As Integer = 0, Optional ByVal int_start_row As Integer = 0, Optional ByVal str_header_val As String = "") As Double
'***********************************************************************************
' Function Array_Max(ar_Array As Variant, Optional ByVal int_col As Integer = 0, Optional ByVal int_row As Integer = 0, Optional ByVal int_start_col As Integer = 0, Optional ByVal int_start_row As Integer = 0)
' This function returns the maximulm value on a given row or given column of an array
' ar_Array : input Array
' int_col : if provided then the maximum will be performed only in the given column
' int_row : if provided then the maximum  will be performed only in the given row
' int_start_col : if provided then the max. finding will be eprformed only after the given column
' int_start_row : if provided then the max. finding will be eprformed only after the given row
' str_header_val : if provided the max will be computed only for the headers containing this string value
'***********************************************************************************
Dim int_LB, int_UB, int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_i, int_j As Integer
    Dim double_max As Double
    
    double_max = -100000
    
    If str_header_val <> "" Then
        int_LB1 = LBound(ar_Array, 1)
        int_UB1 = UBound(ar_Array, 1)
        int_LB2 = LBound(ar_Array, 2)
        int_UB2 = UBound(ar_Array, 2)
        For int_i = WorksheetFunction.Max((int_LB1 + int_start_row), 2) To int_UB1
            For int_j = (int_LB2 + int_start_col) To int_UB2
                If InStr(1, ar_Array(1, int_j), str_header_val) <> 0 Then
                    If ar_Array(int_i, int_j) > double_max Then double_max = ar_Array(int_i, int_j)
                End If
            Next int_j
        Next int_i
    ElseIf int_row = 0 And int_col = 0 Then
        int_LB = LBound(ar_Array)
        int_UB = UBound(ar_Array)
        For int_i = int_LB To int_UB
            If ar_Array(int_i) > double_max Then double_max = ar_Array(int_i)
        Next int_i
    ElseIf int_row = 0 Then
        int_LB = LBound(ar_Array, 1)
        int_UB = UBound(ar_Array, 1)
        For int_i = (int_LB + int_start_row) To int_UB
            If ar_Array(int_i, int_col) > double_max Then double_max = ar_Array(int_i, int_col)
        Next int_i
    ElseIf int_col = 0 Then
        int_LB = LBound(ar_Array, 2)
        int_UB = UBound(ar_Array, 2)
        For int_i = (int_LB + int_start_col) To int_UB
            If ar_Array(int_row, int_i) > double_max Then double_max = ar_Array(int_row, int_i)
        Next int_i
    Else
        int_LB1 = LBound(ar_Array, 1)
        int_UB1 = UBound(ar_Array, 1)
        int_LB2 = LBound(ar_Array, 2)
        int_UB2 = UBound(ar_Array, 2)
        For int_i = (int_LB1 + int_start_row) To int_UB1
            For int_j = (int_LB2 + int_start_col) To int_UB2
                If ar_Array(int_i, int_j) > double_max Then double_max = ar_Array(int_i, int_j)
            Next int_j
        Next int_i
    End If
    
    Array_Max = WorksheetFunction.Max(double_max, 0)
End Function


Public Function Array_Mean(ar_Array As Variant, Optional ByVal int_col As Integer = 0, Optional ByVal int_row As Integer = 0, Optional ByVal int_start_col As Integer = 0, Optional ByVal int_start_row As Integer = 0, Optional ByVal str_header_val As String = "", Optional ByVal str_type As String = "") As Double
'***********************************************************************************
' Function Array_Mean(ar_Array As Variant, Optional ByVal int_col As Integer = 0, Optional ByVal int_row As Integer = 0, Optional ByVal int_start_col As Integer = 0, Optional ByVal int_start_row As Integer = 0)
' This fct returns the mean value on a given row or given column of an array
' ar_Array : input Array
' int_col : if provided then the mean will be performed only in the given column
' int_row : if provided then the mean will be performed only in the given row
' int_start_col : if provided then the mean. finding will be eprformed only after the given column
' int_start_row : if provided then the mean. finding will be eprformed only after the given row
' str_header_val : if provided the mean will be computed only for the headers containing this string value
'***********************************************************************************
Dim int_LB, int_UB, int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_nb, int_i, int_j As Integer
    Dim double_mean As Double
    
    double_mean = 0
    int_nb = 0
    
    If str_header_val <> "" Then
        int_LB1 = LBound(ar_Array, 1)
        int_UB1 = UBound(ar_Array, 1)
        int_LB2 = LBound(ar_Array, 2)
        int_UB2 = UBound(ar_Array, 2)
        For int_i = WorksheetFunction.Max((int_LB1 + int_start_row), 2) To int_UB1
            For int_j = (int_LB2 + int_start_col) To int_UB2
                If ar_Array(1, int_j) = str_header_val <> 0 Then
                    double_mean = double_mean + ar_Array(int_i, int_j)
                    int_nb = int_nb + 1
                End If
            Next int_j
        Next int_i
    ElseIf int_row = 0 And int_col = 0 Then
        int_LB = LBound(ar_Array)
        int_UB = UBound(ar_Array)
        For int_i = int_LB To int_UB
            double_mean = double_mean + ar_Array(int_i)
            int_nb = int_nb + 1
        Next int_i
    ElseIf int_row = 0 Then
        int_LB = LBound(ar_Array, 1)
        int_UB = UBound(ar_Array, 1)
        For int_i = (int_LB + int_start_row) To int_UB
            double_mean = double_mean + ar_Array(int_i, int_col)
            int_nb = int_nb + 1
        Next int_i
    ElseIf int_col = 0 Then
        int_LB = LBound(ar_Array, 2)
        int_UB = UBound(ar_Array, 2)
        For int_i = (int_LB + int_start_col) To int_UB
            double_mean = double_mean + ar_Array(int_row, int_i)
            int_nb = int_nb + 1
        Next int_i
    Else
        int_LB1 = LBound(ar_Array, 1)
        int_UB1 = UBound(ar_Array, 1)
        int_LB2 = LBound(ar_Array, 2)
        int_UB2 = UBound(ar_Array, 2)
        For int_i = (int_LB1 + int_start_row) To int_UB1
            For int_j = (int_LB2 + int_start_col) To int_UB2
                double_mean = double_mean + ar_Array(int_i, int_j)
                int_nb = int_nb + 1
            Next int_j
        Next int_i
    End If
    
    If int_nb = 0 Then
        Array_Mean = 0
    Else
        If str_type = "" Then
            Array_Mean = double_mean / int_nb
        ElseIf UCase(str_type) = "SUM" Then
            Array_Mean = double_mean
        ElseIf UCase(str_type = "COUNT") Then
            Array_Mean = int_nb
        End If
    End If
End Function


Public Function Array_Unique(ar_Array As Variant, ar_exitValues As Variant, Optional ByVal int_col As Integer = 0, Optional ByVal int_row As Integer = 0, Optional ByVal int_start_col As Integer = 0, Optional ByVal int_start_row As Integer = 0) As Variant
'***********************************************************************************
' Fct Array_Unique(ar_Array As Variant, Optional ByVal int_col As Integer = 0, Optional ByVal int_row As Integer = 0, Optional ByVal int_start_col As Integer = 0, Optional ByVal int_start_row As Integer = 0)
' This fct returns an array with the unique elements of ar_Array specified column or row
' ar_Array : input Array
' ar_exitValues : array containing the characters triggering the exit of the process - Example :  ar_exitValues=Array("", Empty)
' int_col : if provided then the unique elements will be extracted only in the given column
' int_row : if provided then the unique elements will be extracted only in the given row
' int_start_col : if provided then the unique elements will be extracted only after the given column
' int_start_row : if provided then the unique elements will be extracted only after the given row
'***********************************************************************************
    LB1 = LBound(ar_Array, 1)
    LB2 = LBound(ar_Array, 2)
    UB1 = UBound(ar_Array, 1)
    UB2 = UBound(ar_Array, 2)
    int_k = 0
    If int_col > 0 Then
        ReDim ar_Array_Unique(1 To (UB1 + (1 - LB1)))
        For int_i = (LB1 + int_start_row) To UB1
            If IsInArray(ar_Array(int_i, int_col), ar_exitValues) Then Exit For
            If int_k = 0 Then
                int_k = int_k + 1
                ar_Array_Unique(int_k) = ar_Array(int_i, int_col)
            Else
                If Not IsInArray(ar_Array(int_i, int_col), ar_Array_Unique) Then
                    int_k = int_k + 1
                    ar_Array_Unique(int_k) = ar_Array(int_i, int_col)
                End If
            End If
        Next int_i
    Else
        ReDim ar_Array_Unique(1 To (UB2 + (1 - LB2)))
        For int_i = (LB2 - int_start_col) To UB2
            If IsInArray(ar_Array(int_row, int_i), ar_exitValues) Then Exit For
            If int_k = 0 Then
                int_k = int_k + 1
                ar_Array_Unique(int_k) = ar_Array(int_row, int_i)
            Else
                If Not IsInArray(ar_Array(int_row, int_i), ar_Array_Unique) Then
                    int_k = int_k + 1
                    ar_Array_Unique(int_k) = ar_Array(int_row, int_i)
                End If
            End If
        Next int_i
    End If
    
    Call Array_Cut1D(ar_Array_Unique, ar_exitValues)
    Array_Unique = ar_Array_Unique
End Function


Public Sub Array_ShiftDim(ByRef ar_Array As Variant, Optional ByVal int_row_shift = 0, Optional ByVal int_col_shift = 0)
'***********************************************************************************
'Sub Array_ShiftDim
'This sub shift the dimensions index of an array
'ar_Array : array containing the data
'int_row_shift : number of row dimension to shift
'int_col_headers:  number of col dimension to shift
'***********************************************************************************
    Dim int_LB1, int_UB1, int_LB2, int_UB2 As Integer
    Dim int_i, int_j As Integer
    
    int_LB1 = LBound(ar_Array, 1)
    int_UB1 = UBound(ar_Array, 1)
    int_LB2 = LBound(ar_Array, 2)
    int_UB2 = UBound(ar_Array, 2)
    Dim ar_ArrayCopy As Variant
    
    ReDim ar_ArrayCopy(int_LB1 To int_UB1, int_LB2 To int_UB2)

    For int_i = int_LB1 To int_UB1
        For int_j = int_LB2 To int_UB2
            ar_ArrayCopy(int_i, int_j) = ar_Array(int_i, int_j)
        Next int_j
    Next int_i
    
    ReDim ar_Array((int_LB1 + int_row_shift) To (int_UB1 + int_row_shift), (int_LB2 + int_col_shift) To (int_UB2 + int_col_shift))
                
    For int_i = int_LB1 To int_UB1
        For int_j = int_LB2 To int_UB2
            ar_Array(int_i + int_row_shift, int_j + int_col_shift) = ar_ArrayCopy(int_i, int_j)
        Next int_j
    Next int_i
End Sub




