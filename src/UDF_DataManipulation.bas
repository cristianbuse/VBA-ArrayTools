Attribute VB_Name = "UDF_DataManipulation"
'''=============================================================================
''' VBA ArrayTools
'''-----------------------------------------------
''' https://github.com/cristianbuse/VBA-ArrayTools
'''-----------------------------------------------
'''
''' Copyright (c) 2012 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to deal
''' in the Software without restriction, including without limitation the rights
''' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
''' copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in all
''' copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
''' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
''' SOFTWARE.
'''=============================================================================

Option Explicit

'*******************************************************************************
''Excel Data Manipulation Module
''Functions in this module are dynamic array formulas that allow easy data
''  manipulation within the Excel interface using the ArrayTools library
''In newer versions of Excel some of the functions presented below are native to
''  Excel (FILTER, SORT, SEQUENCE, UNIQUE and others) but the functionality is
''  slightly different. It was announced in fall 2018 for Office 365 users
''  For example:
''    - DM_FILTER allows filtering on the returned result of another function
''    - DM_SORT function allows case sensitive texts and more
''    - DM_UNIQUE function allows to choose the columns used for uniquifying
''  The functions below are capable to 'spill' on newer Excel versions
'*******************************************************************************

''Important!
'*******************************************************************************
''This module is intended to be used in Microsoft Excel only!
''Call the User-Defined-Functions (UDFs) in this module from Excel Ranges only
''  DO NOT call these functions from VBA! If you need any of the functions below
''  directly in VBA then use their equivalent from the LibArrayTools module
'*******************************************************************************

''Requires:
''  - LibArrayTools: library module with Array/Collection tools

''Exposed Excel UDFs:
''  - DM_ARRAY
''  - DM_FILTER
''  - DM_INSERT
''  - DM_INSERT2
''  - DM_MERGE
''  - DM_REVERSE
''  - DM_SEQUENCE
''  - DM_SLICE
''  - DM_SORT
''  - DM_UNIQUE

'*******************************************************************************
'Turn the below compiler constant to True if you are using the LibUDFs library
'https://github.com/cristianbuse/VBA-FastExcelUDFs
#Const USE_LIB_FAST_UDFS = False
'*******************************************************************************

'###############################################################################
'Register/Unregister Function Help for the Excel Function Arguments Dialog
'Notes:
'   - a dummy parameter is used to hide the methods from the Excel Macro Dialog
'   - ArgumentDescriptions not available for older versions of Excel e.g. 2007
'###############################################################################
Public Sub RegisterDMFunctions(Optional ByVal dummy As Boolean)
    RegisterDMArray
    RegisterDMFilter
    RegisterDMInsert
    RegisterDMInsert2
    RegisterDMMerge
    RegisterDMReverse
    RegisterDMSequence
    RegisterDMSlice
    RegisterDMSort
    RegisterDMUnique
End Sub
Public Sub UnregisterDMFunctions(Optional ByVal dummy As Boolean)
    UnregisterDMArray
    UnregisterDMFilter
    UnRegisterDMInsert
    UnRegisterDMInsert2
    UnregisterDMMerge
    UnregisterDMReverse
    UnregisterDMSequence
    UnregisterDMSlice
    UnregisterDMSort
    UnregisterDMUnique
End Sub

'*******************************************************************************
'Returns specified value(s) in a 1D or 2D array format
'Parameters:
'   - columnsCount: the number of columns the output 2D array will have
'     if 0 then a 1D array will be returned
'   - values: the value(s) to be returned
'Notes:
'   - uses LibArrayTools functions:
'       * CollectionTo1DArray
'       * CollectionTo2Darray
'       * ReplaceEmptyInArray
'       * ValuesToCollection
'   - can be used to easily 'JOIN' ranges/arrays of values
'*******************************************************************************
Public Function DM_ARRAY(ByVal columnsCount As Long _
    , ParamArray values() As Variant _
) As Variant
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    If columnsCount < 0 Then GoTo FailInput
    '
    'Get values to Collection
    Dim v As Variant: v = values
    Dim coll As Collection
    Set coll = LibArrayTools.ValuesToCollection(v, nestNone, rowWise)
    '
    'Return 1D or 2D array
    If columnsCount = 0 Then
        DM_ARRAY = LibArrayTools.CollectionTo1DArray(coll)
    Else
        'If the number of elements in the collection is not divisible by the
        '   columns count then add #N/A to fill the last row
        Dim remainders As Long: remainders = coll.Count Mod columnsCount
        '
        If remainders > 0 Then
            Dim errNA As Variant: errNA = VBA.CVErr(xlErrNA)
            Dim i As Long
            '
            For i = 1 To columnsCount - remainders
                coll.Add errNA
            Next i
        End If
        DM_ARRAY = LibArrayTools.CollectionTo2DArray(coll, columnsCount)
    End If
    '
    'Replace the special value Empty with empty String so it is not returned
    '   as 0 (zero) in the caller Range
    LibArrayTools.ReplaceEmptyInArray DM_ARRAY, vbNullString
Exit Function
FailInput:
    DM_ARRAY = VBA.CVErr(xlErrValue)
End Function

'###############################################################################
'Help for the Function Arguments Dialog in Excel - DM_ARRAY
'###############################################################################
Private Sub RegisterDMArray()
    Dim arg1 As String
    Dim arg2 As String
    '
    arg1 = "the number of columns the output 2D array will have" & vbNewLine _
        & "Use 0 (zero) to return a 1D Array"
    arg2 = "any value (Range, Named Range, Array, number, text etc.)"
    '
    Application.MacroOptions Macro:="DM_ARRAY" _
        , Description:="Returns specified value(s) in a joined 1D or 2D array" _
        , ArgumentDescriptions:=Array(arg1, arg2)
End Sub
Private Sub UnregisterDMArray()
    Application.MacroOptions Macro:="DM_ARRAY", Description:=Empty _
        , ArgumentDescriptions:=Array(Empty, Empty)
End Sub

'*******************************************************************************
'Filters a 2D array (or a 1-Area Range) vertically by the specified column index
'Returns:
'   - the filtered array or an Excel #VALUE! or #CALC! error
'Parameters:
'   - columnIndex: the index of the column to be used for filtering
'   - arr: the 2D array that needs filtering
'   - filters: pairs of Operator and Comparison Value to be used for filtering
'     Ex. {">=", 3, "=<", 17, "NOT IN", MyRange} would return numbers from 'arr'
'         between 3 and 17 that are not in the values of MyRange
'Notes:
'   - uses LibArrayTools functions:
'       * CreateFiltersArray
'       * Filter2DArray
'       * GetArrayElemCount
'       * OneDArrayTo2DArray
'       * ReplaceEmptyInArray
'   - single values are converted to 1-element 1D array
'   - 1D arrays are converted to 1-row 2D arrays
'   - accepted operators (as Strings - see 'GetCondOperatorFromText' method)
'       * comparison operators: =, <, >, <=, >=, <>
'       * inclusion operators: IN , NOT IN
'         accepts a list (array/range) as the comparison value
'       * pattern matching operators: LIKE, NOT LIKE
'         accepts pattern as the comparison value. For available patterns check
'         the help for the VBA LIKE operator
'*******************************************************************************
Public Function DM_FILTER(ByVal columnIndex As Long, ByRef arr As Variant _
    , ParamArray filters() As Variant _
) As Variant
Attribute DM_FILTER.VB_Description = "Filters a 2D array/range by the specified column index"
Attribute DM_FILTER.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    'Only accept 1-Area Ranges. This could alternatively be changed to ignore
    '   the extra Areas by arr = arr.Areas(1).Value2 instead of the 2 lines
    If VBA.TypeName(arr) = "Range" Then
        If arr.Areas.Count > 1 Then GoTo FailInput
        arr = arr.Value2
    End If
    '
    'Convert single value to 1-element 1D array
    If Not VBA.IsArray(arr) Then arr = Array(arr)
    '
    Select Case LibArrayTools.GetArrayDimsCount(arr)
    Case 2
        'Continue
    Case 1 'Convert to 1-row 2D array and adjust column index
        Dim colsCount As Long: colsCount = UBound(arr) - LBound(arr) + 1
        arr = LibArrayTools.OneDArrayTo2DArray(arr, colsCount)
        columnIndex = columnIndex + LBound(arr, 2) - 1
    Case Else
        GoTo FailInput 'Should not happen if called from Excel
    End Select
    '
    On Error GoTo ErrorHandler
    DM_FILTER = LibArrayTools.Filter2DArray(arr, columnIndex _
        , LibArrayTools.CreateFiltersArray(filters))
    '
    If LibArrayTools.GetArrayElemCount(DM_FILTER) = 0 Then
        DM_FILTER = VBA.CVErr(xlErrCalc)
    Else
        'Replace the special value Empty with empty String so it is not
        '   returned as 0 (zero) in the caller Range
        LibArrayTools.ReplaceEmptyInArray DM_FILTER, vbNullString
    End If
Exit Function
ErrorHandler:
FailInput:
    DM_FILTER = VBA.CVErr(xlErrValue)
End Function

'###############################################################################
'Help for the Function Arguments Dialog in Excel - DM_FILTER
'###############################################################################
Private Sub RegisterDMFilter()
    Dim arg1 As String
    Dim arg2 As String
    Dim arg3 As String
    '
    arg1 = "the index of the column to be used for filtering"
    arg2 = "the 2D array that needs filtering"
    arg3 = "'text operator' and 'comparison value(s)' pairs" & vbNewLine _
        & "text operator: =, <, >, <=, >=, <>, IN , NOT IN, LIKE, NOT LIKE" _
        & vbNewLine & "comparison value(s). LIKE and NOT LIKE accept arrays" _
    '
    Application.MacroOptions Macro:="DM_FILTER" _
        , Description:="Filters a 2D array/range by the specified column index" _
        , ArgumentDescriptions:=Array(arg1, arg2, arg3)
End Sub
Private Sub UnregisterDMFilter()
    Application.MacroOptions Macro:="DM_FILTER", Description:=Empty _
        , ArgumentDescriptions:=Array(Empty, Empty, Empty)
End Sub

'*******************************************************************************
'Inserts rows in a 2D Array or a 1-Area Range before the specified row index
'Parameters:
'   - arr: a 2D array to insert into
'   - rowsCount: the number of rows to insert
'   - beforeRow: the index of the row before which rows will be inserted
'Notes:
'   - uses LibArrayTools functions:
'       * GetArrayDimsCount
'       * InsertRowsAtIndex
'       * OneDArrayTo2DArray
'       * ReplaceEmptyInArray
'   - single values are converted to 1-element 1D array
'   - 1D arrays are converted to 1-row 2D arrays
'*******************************************************************************
Public Function DM_INSERT(ByRef arr As Variant, ByVal rowsCount As Long _
    , ByVal beforeRow As Long _
) As Variant
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    'Only accept 1-Area Ranges. This could alternatively be changed to ignore
    '   the extra Areas by arr = arr.Areas(1).Value2 instead of the 2 lines
    If VBA.TypeName(arr) = "Range" Then
        If arr.Areas.Count > 1 Then GoTo FailInput
        arr = arr.Value2
    End If
    '
    'Convert single value to 1-element 1D array
    If Not VBA.IsArray(arr) Then arr = Array(arr)
    '
    Select Case LibArrayTools.GetArrayDimsCount(arr)
    Case 2
        'Continue
    Case 1 'Convert to 1-row 2D array and adjust beforeRow index
        Dim colsCount As Long: colsCount = UBound(arr) - LBound(arr) + 1
        arr = LibArrayTools.OneDArrayTo2DArray(arr, colsCount)
        beforeRow = beforeRow + LBound(arr, 1) - 1
    Case Else
        GoTo FailInput 'Should not happen if called from Excel
    End Select
    On Error GoTo ErrorHandler
    DM_INSERT = LibArrayTools.InsertRowsAtIndex(arr, rowsCount, beforeRow)
    '
    'Replace the special value Empty with empty String so it is not returned
    '   as 0 (zero) in the caller Range
    LibArrayTools.ReplaceEmptyInArray DM_INSERT, vbNullString
Exit Function
ErrorHandler:
FailInput:
    DM_INSERT = VBA.CVErr(xlErrValue)
End Function

'###############################################################################
'Help for the Function Arguments Dialog in Excel - DM_INSERT
'###############################################################################
Private Sub RegisterDMInsert()
    Dim arg1 As String
    Dim arg2 As String
    Dim arg3 As String
    '
    arg1 = "a 2D array to insert into. 1D arrays are considered 1-row 2D"
    arg2 = "the number of rows to insert"
    arg3 = "the index of the row before which rows will be inserted"
    '
    Application.MacroOptions Macro:="DM_INSERT" _
        , Description:="Inserts rows in a 2D Array" _
        , ArgumentDescriptions:=Array(arg1, arg2, arg3)
End Sub
Private Sub UnRegisterDMInsert()
    Application.MacroOptions Macro:="DM_INSERT", Description:=Empty _
        , ArgumentDescriptions:=Array(Empty, Empty, Empty)
End Sub

'*******************************************************************************
'Inserts rows in a 2D Array between rows with different values (on the specified
'   column) and optionally at the top and/or bottom of the array
'Parameters:
'   - arr: a 2D array to insert into
'   - rowsCount: the number of rows to insert at each value change
'   - columnIndex: the index of the column used for row comparison
'   - [topRowsCount]: number of rows to insert before array. Default is 0
'   - [bottomRowsCount]: number of rows to insert after array. Default is 0
'Notes:
'   - uses LibArrayTools functions:
'       * GetArrayDimsCount
'       * InsertRowsAtValChange
'       * OneDArrayTo2DArray
'       * ReplaceEmptyInArray
'   - single values are converted to 1-element 1D array
'   - 1D arrays are converted to 1-row 2D arrays
'*******************************************************************************
Public Function DM_INSERT2(ByRef arr As Variant _
    , ByVal rowsCount As Long, ByVal columnIndex As Long _
    , Optional ByVal topRowsCount As Long = 0 _
    , Optional ByVal bottomRowsCount As Long = 0 _
) As Variant
Attribute DM_INSERT2.VB_Description = "Inserts rows in a 2D Array between rows with different values (on the specified column) and optionally at the top and/or bottom of the array"
Attribute DM_INSERT2.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    'Only accept 1-Area Ranges. This could alternatively be changed to ignore
    '   the extra Areas by arr = arr.Areas(1).Value2 instead of the 2 lines
    If VBA.TypeName(arr) = "Range" Then
        If arr.Areas.Count > 1 Then GoTo FailInput
        arr = arr.Value2
    End If
    '
    'Convert single value to 1-element 1D array
    If Not VBA.IsArray(arr) Then arr = Array(arr)
    '
    Select Case LibArrayTools.GetArrayDimsCount(arr)
    Case 2 'Continue
    Case 1 'Convert to 1-row 2D array and adjust column index
        Dim colsCount As Long: colsCount = UBound(arr) - LBound(arr) + 1
        arr = LibArrayTools.OneDArrayTo2DArray(arr, colsCount)
        columnIndex = columnIndex + LBound(arr, 2) - 1
    Case Else
        GoTo FailInput 'Should not happen if called from Excel
    End Select
    On Error GoTo ErrorHandler
    DM_INSERT2 = LibArrayTools.InsertRowsAtValChange( _
        arr, rowsCount, columnIndex, topRowsCount, bottomRowsCount)
    '
    'Replace the special value Empty with empty String so it is not returned
    '   as 0 (zero) in the caller Range
    LibArrayTools.ReplaceEmptyInArray DM_INSERT2, vbNullString
Exit Function
ErrorHandler:
FailInput:
    DM_INSERT2 = VBA.CVErr(xlErrValue)
End Function

'###############################################################################
'Help for the Function Arguments Dialog in Excel - DM_INSERT2
'###############################################################################
Private Sub RegisterDMInsert2()
    Dim arg1 As String
    Dim arg2 As String
    Dim arg3 As String
    Dim arg4 As String
    Dim arg5 As String
    '
    arg1 = "a 2D array to insert into. 1D arrays are considered 1-row 2D"
    arg2 = "the number of rows to insert at each value change"
    arg3 = "the index of the column used for row comparison"
    arg4 = "[Optional]" & vbNewLine _
        & "the number of rows to insert at the top of the array. Default is 0"
    arg5 = "[Optional]" & vbNewLine _
        & "the number of rows to insert at the bottom of the array. Default is 0"
    '
    Application.MacroOptions Macro:="DM_INSERT2" _
        , Description:="Inserts rows in a 2D Array between rows with " _
            & "different values (on the specified column) and optionally at " _
            & "the top and/or bottom of the array" _
        , ArgumentDescriptions:=Array(arg1, arg2, arg3, arg4, arg5)
End Sub
Private Sub UnRegisterDMInsert2()
    Application.MacroOptions Macro:="DM_INSERT2", Description:=Empty _
        , ArgumentDescriptions:=Array(Empty, Empty, Empty, Empty, Empty)
End Sub

'*******************************************************************************
'Merges/Combines two 1D/2D Arrays or 1-Area Ranges
'Returns:
'   - the merged array or the Excel #VALUE! error
'Parameters:
'   - arr1: the first 1D/2D Array
'   - arr2: the second 1D/2D Array
'   - [verticalMerge]:
'       * TRUE - arrays are combined vertically
'       * FALSE - arrays are combined horizontally (default)
'Notes:
'   - single values are converted to 1-element 1D array
'   - 1D arrays are converted to 1-row 2D arrays
'   - uses LibArrayTools functions:
'       * GetArrayDimsCount
'       * Merge2DArrays
'       * OneDArrayTo2DArray
'       * ReplaceEmptyInArray
'*******************************************************************************
Public Function DM_MERGE(ByRef arr1 As Variant, ByRef arr2 As Variant _
    , Optional ByVal verticalMerge As Boolean = False _
) As Variant
Attribute DM_MERGE.VB_Description = "Merges/Combines two 1D/2D Arrays"
Attribute DM_MERGE.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    'Only accept 1-Area Ranges. This could alternatively be changed to ignore
    '   the extra Areas by arr = arr.Areas(1).Value2 instead of the 2 lines
    If VBA.TypeName(arr1) = "Range" Then
        If arr1.Areas.Count > 1 Then GoTo FailInput
        arr1 = arr1.Value2
    End If
    If VBA.TypeName(arr2) = "Range" Then
        If arr2.Areas.Count > 1 Then GoTo FailInput
        arr2 = arr2.Value2
    End If
    '
    'Convert single value to 1-element 1D array
    If Not VBA.IsArray(arr1) Then arr1 = Array(arr1)
    If Not VBA.IsArray(arr2) Then arr2 = Array(arr2)
    '
    Dim columnsCount As Long
    '
    'Convert 1D arrays to 1-row 2D arrays
    If LibArrayTools.GetArrayDimsCount(arr1) = 1 Then
        columnsCount = UBound(arr1) - LBound(arr1) + 1
        arr1 = LibArrayTools.OneDArrayTo2DArray(arr1, columnsCount)
    End If
    If LibArrayTools.GetArrayDimsCount(arr2) = 1 Then
        columnsCount = UBound(arr2) - LBound(arr2) + 1
        arr2 = LibArrayTools.OneDArrayTo2DArray(arr2, columnsCount)
    End If
    '
    On Error GoTo ErrorHandler
    DM_MERGE = LibArrayTools.Merge2DArrays(arr1, arr2, verticalMerge)
    '
    'Replace the special value Empty with empty String so it is not returned
    '   as 0 (zero) in the caller Range
    LibArrayTools.ReplaceEmptyInArray DM_MERGE, vbNullString
Exit Function
ErrorHandler:
FailInput:
    DM_MERGE = VBA.CVErr(xlErrValue)
End Function

'###############################################################################
'Help for the Function Arguments Dialog in Excel - DM_MERGE
'###############################################################################
Private Sub RegisterDMMerge()
    Dim arg1 As String
    Dim arg2 As String
    Dim arg3 As String
    '
    arg1 = "the first 1D/2D Array. 1D arrays are considered 1-row 2D"
    arg2 = "the second 1D/2D Array. 1D arrays are considered 1-row 2D"
    arg3 = "[Optional]" & vbNewLine & "True - arrays are combined vertically" _
        & vbNewLine & "False - arrays are combined horizontally (Default)"
    '
    Application.MacroOptions Macro:="DM_MERGE" _
        , Description:="Merges/Combines two 1D/2D Arrays" _
        , ArgumentDescriptions:=Array(arg1, arg2, arg3)
End Sub
Private Sub UnregisterDMMerge()
    Application.MacroOptions Macro:="DM_MERGE", Description:=Empty _
        , ArgumentDescriptions:=Array(Empty, Empty, Empty)
End Sub

'*******************************************************************************
'Reverses (in groups) a 1D/2D Array or a 1-Area Range
'Returns:
'   - the reversed array or the Excel #VALUE! error
'Parameters:
'   - arr: a 1D/2D Array or Range of values to be reversed
'   - [groupSize]: the number of values in each group
'   - [verticalFlip]:
'       * TRUE - reverse vertically
'       * FALSE - reverse horizontally (default)
'Examples:
'   - arr = [1,2,3,4], groupSize = 1, verticalFlip = True  > return is [1,2,3,4]
'   - arr = [1,2,3,4], groupSize = 1, verticalFlip = False > return is [4,3,2,1]
'   - arr = [1,2,3,4], groupSize = 2, verticalFlip = False > return is [3,4,1,2]
'   - arr = [1,2,3,4], groupSize = 2, verticalFlip = False > return is [3,4,1,2]
'           [5,6,7,8]                                                  [7,8,5,6]
'   - arr = [1,2,3,4], groupSize = 1, verticalFlip = True  > return is [5,6,7,8]
'           [5,6,7,8]                                                  [1,2,3,4]
'Notes:
'   - single values are converted to 1-element 1D array
'   - uses LibArrayTools functions:
'       * GetArrayDimsCount
'       * ReplaceEmptyInArray
'       * Reverse1DArray
'       * Reverse2DArray
'*******************************************************************************
Public Function DM_REVERSE(ByRef arr As Variant _
    , Optional ByVal groupSize As Long = 1 _
    , Optional ByVal verticalFlip As Boolean = False _
) As Variant
Attribute DM_REVERSE.VB_Description = "Reverses (in groups) a 2D Array or a 1-Area Range"
Attribute DM_REVERSE.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    'Only accept 1-Area Ranges. This could alternatively be changed to ignore
    '   the extra Areas by arr = arr.Areas(1).Value2 instead of the 2 lines
    If VBA.TypeName(arr) = "Range" Then
        If arr.Areas.Count > 1 Then GoTo FailInput
        arr = arr.Value2
    End If
    '
    'Convert single value to 1-element 1D array
    If Not VBA.IsArray(arr) Then arr = Array(arr)
    '
    On Error GoTo ErrorHandler
    Select Case LibArrayTools.GetArrayDimsCount(arr)
    Case 1
        If Not verticalFlip Then LibArrayTools.Reverse1DArray arr, groupSize
    Case 2
        LibArrayTools.Reverse2DArray arr, groupSize, verticalFlip
    Case Else
        GoTo FailInput 'Should not happen if called from Excel
    End Select
    '
    'Replace the special value Empty with empty String so it is not returned
    '   as 0 (zero) in the caller Range
    LibArrayTools.ReplaceEmptyInArray arr, vbNullString
    '
    DM_REVERSE = arr
Exit Function
ErrorHandler:
FailInput:
    DM_REVERSE = VBA.CVErr(xlErrValue)
End Function

'###############################################################################
'Help for the Function Arguments Dialog in Excel - DM_REVERSE
'###############################################################################
Private Sub RegisterDMReverse()
    Dim arg1 As String
    Dim arg2 As String
    Dim arg3 As String
    '
    arg1 = "a 1D/2D Array or Range of values to be reversed" & vbNewLine _
        & "1D arrays are considered 1-row 2D"
    arg2 = "[Optional]" & vbNewLine _
        & "the number of values in each group. Default is 1"
    arg3 = "[Optional]" & vbNewLine & "True - reverse vertically" _
        & vbNewLine & "False - reverse horizontally (Default)"
    '
    Application.MacroOptions Macro:="DM_REVERSE" _
        , Description:="Reverses (in groups) a 2D Array or a 1-Area Range" _
        , ArgumentDescriptions:=Array(arg1, arg2, arg3)
End Sub
Private Sub UnregisterDMReverse()
    Application.MacroOptions Macro:="DM_REVERSE", Description:=Empty _
        , ArgumentDescriptions:=Array(Empty, Empty, Empty)
End Sub

'*******************************************************************************
'Creates an arithmetic progression sequence as a 2D array
'Returns:
'   - the arithmetic progression or the Excel #VALUE! error
'Parameters:
'   - rowsCount: the number of rows in the output 2D Array
'   - [columnsCount]: the number of columns in the output 2D Array
'   - [initialTerm]: the value of the first term
'   - [commonDifference]: the difference between any 2 consecutive terms
'Notes:
'   - uses LibArrayTools functions:
'       * Sequence2D
'*******************************************************************************
Public Function DM_SEQUENCE(ByVal rowsCount As Long _
    , Optional ByVal columnsCount As Long = 1 _
    , Optional ByVal initialTerm As Double = 1 _
    , Optional ByVal commonDifference As Double = 1 _
) As Variant
Attribute DM_SEQUENCE.VB_Description = "Returns an arithmetic progression sequence as 2D array"
Attribute DM_SEQUENCE.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    On Error GoTo ErrorHandler
    DM_SEQUENCE = LibArrayTools.Sequence2D(rowsCount * columnsCount, initialTerm _
        , commonDifference, columnsCount)
Exit Function
ErrorHandler:
    DM_SEQUENCE = VBA.CVErr(xlErrValue)
End Function

'###############################################################################
'Help for the Function Arguments Dialog in Excel - DM_SEQUENCE
'###############################################################################
Private Sub RegisterDMSequence()
    Dim arg1 As String
    Dim arg2 As String
    Dim arg3 As String
    Dim arg4 As String
    '
    arg1 = "number of rows in the output 2D Array"
    arg2 = "[Optional]" & vbNewLine & "number of output columns. Default is 1"
    arg3 = "[Optional]" & vbNewLine & "value of the first term. Default is 1"
    arg4 = "[Optional]" & vbNewLine _
        & "difference between any 2 consecutive terms. Default is 1"
    '
    Application.MacroOptions Macro:="DM_SEQUENCE " _
        , Description:="Returns an arithmetic progression sequence as 2D array" _
        , ArgumentDescriptions:=Array(arg1, arg2, arg3, arg4)
End Sub
Private Sub UnregisterDMSequence()
    Application.MacroOptions Macro:="DM_SEQUENCE ", Description:=Empty _
        , ArgumentDescriptions:=Array(Empty, Empty, Empty, Empty)
End Sub

'*******************************************************************************
'Slices a 1D/2D Array or a 1-Area Range
'Returns:
'   - the array slice or an Excel #VALUE! or #REF! error
'Parameters:
'   - arr: a 1D/2D array to slice
'   - startRow: the index of the first row to be added to result
'   - startColumn: the index of the first column to be added to result
'   - [height_]: the number of rows to be returned. Default is 1.
'                Use 0 to get all rows starting from startRow
'   - [width_]: the number of columns to be returned. Default is 1.
'               Use 0 to get all columns starting from startColumn
'Notes:
'   - excess height or width is ignored
'   - uses LibArrayTools functions:
'       * GetArrayDimsCount
'       * ReplaceEmptyInArray
'       * Slice1DArray
'       * Slice2DArray
'   - single values are converted to 1-element 1D array
'*******************************************************************************
Public Function DM_SLICE(ByRef arr As Variant _
    , ByVal startRow As Long, ByVal startColumn As Long _
    , Optional ByVal height_ As Long = 1, Optional ByVal width_ As Long = 1 _
) As Variant
Attribute DM_SLICE.VB_Description = "Slices a 1D/2D Array or a 1-Area Range"
Attribute DM_SLICE.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    'Only accept 1-Area Ranges. This could alternatively be changed to ignore
    '   the extra Areas by arr = arr.Areas(1).Value2 instead of the 2 lines
    If VBA.TypeName(arr) = "Range" Then
        If arr.Areas.Count > 1 Then GoTo FailInput
        arr = arr.Value2
    End If
    '
    'Convert single value to 1-element 1D array
    If Not VBA.IsArray(arr) Then
        arr = Array(arr)
        startColumn = startColumn + LBound(arr) - 1
    End If
    '
    On Error GoTo FailReference
    Select Case LibArrayTools.GetArrayDimsCount(arr)
    Case 1
        If startRow <> 1 Or height_ < 0 Then GoTo FailReference
        If width_ = 0 Then width_ = UBound(arr) - startColumn + 1
        DM_SLICE = LibArrayTools.Slice1DArray(arr, startColumn, width_)
    Case 2
        If height_ = 0 Then height_ = UBound(arr, 1) - startRow + 1
        If width_ = 0 Then width_ = UBound(arr, 2) - startColumn + 1
        DM_SLICE = LibArrayTools.Slice2DArray(arr, startRow, startColumn _
            , height_, width_)
    Case Else
        GoTo FailInput 'Should not happen if called from Excel
    End Select
    '
    'Replace the special value Empty with empty String so it is not returned
    '   as 0 (zero) in the caller Range
    LibArrayTools.ReplaceEmptyInArray DM_SLICE, vbNullString
Exit Function
FailInput:
    DM_SLICE = VBA.CVErr(xlErrValue)
Exit Function
FailReference:
    DM_SLICE = VBA.CVErr(xlErrRef)
End Function

'###############################################################################
'Help for the Function Arguments Dialog in Excel - DM_SLICE
'###############################################################################
Private Sub RegisterDMSlice()
    Dim arg1 As String
    Dim arg2 As String
    Dim arg3 As String
    Dim arg4 As String
    Dim arg5 As String
    '
    arg1 = "a 1D/2D array to slice. 1D arrays can be viewed as 1-row 2D"
    arg2 = "the index of the first row to be added to result"
    arg3 = "the index of the first column to be added to result"
    arg4 = "[Optional]" & vbNewLine _
        & "the number of rows to be returned. Default is 1. " _
        & "Use 0 to get all rows starting from startRow"
    arg5 = "[Optional]" & vbNewLine _
        & "the number of columns to be returned. Default is 1" _
        & "Use 0 to get all columns starting from startColumn"
    '
    Application.MacroOptions Macro:="DM_SLICE" _
        , Description:="Slices a 1D/2D Array or a 1-Area Range" _
        , ArgumentDescriptions:=Array(arg1, arg2, arg3, arg4, arg5)
End Sub
Private Sub UnregisterDMSlice()
    Application.MacroOptions Macro:="DM_SLICE", Description:=Empty _
        , ArgumentDescriptions:=Array(Empty, Empty, Empty, Empty, Empty)
End Sub

'*******************************************************************************
'Sorts a 1D/2D Array or a 1-Area Range
'Returns:
'   - the sorted array or the Excel #VALUE! error
'Parameters:
'   - arr: a 1D/2D Array or Range of values to sort
'   - [sortIndex]: the index of the column/row used for sorting
'   - [sortAscending]:
'       * TRUE - Ascending (default)
'       * FALSE - Descending
'   - [sortTextNumberAsNumber]:
'       * TRUE - numbers stored as texts are considered numbers (default)
'       * FALSE - numbers stored as texts are considered texts
'   - [caseSensitiveTexts]:
'       * TRUE - compare texts as case-sensitive
'       * FALSE - ignore case when comparing texts (default)
'   - [verticalSort]:
'       * TRUE - sorts vertically (by column) (default)
'       * FALSE - sorts horizontally (by row)
'Notes:
'   - single values are converted to 1-element 1D array
'   - when sorting a 1D array vertically, the array is regarded as a 1-row 2D
'     array and is returned as-is
'   - when sorting a 1D array the sortIndex is not used but must be 1 or omitted
'   - when sorting a 2D array horizontally, the array is first transposed, then
'     sorted by column and finally transposed back with the end result being
'     that the array was actually sorted by row
'   - uses LibArrayTools functions:
'       * GetArrayDimsCount
'       * ReplaceEmptyInArray
'       * Sort1DArray
'       * Sort2DArray
'       * TransposeArray
'*******************************************************************************
Public Function DM_SORT(ByRef arr As Variant _
    , Optional ByVal sortIndex As Long = 1 _
    , Optional ByVal sortAscending As Boolean = True _
    , Optional ByVal sortTextNumberAsNumber As Boolean = True _
    , Optional ByVal caseSensitiveTexts As Boolean = False _
    , Optional ByVal verticalSort As Boolean = True _
) As Variant
Attribute DM_SORT.VB_Description = "Sorts a 1D/2D Array or a 1-Area Range"
Attribute DM_SORT.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    'Only accept 1-Area Ranges. This could alternatively be changed to ignore
    '   the extra Areas by arr = arr.Areas(1).Value2 instead of the 2 lines
    If VBA.TypeName(arr) = "Range" Then
        If arr.Areas.Count > 1 Then GoTo FailInput
        arr = arr.Value2
    End If
    '
    'Convert single value to 1-element 1D array
    If Not VBA.IsArray(arr) Then
        arr = Array(arr)
        sortIndex = sortIndex + LBound(arr) - 1
    End If
    '
    Select Case LibArrayTools.GetArrayDimsCount(arr)
    Case 1
        If verticalSort Then
            'Still check if column index is valid
            If sortIndex < LBound(arr) Then GoTo FailInput
            If sortIndex > UBound(arr) Then GoTo FailInput
        Else
            'You can only sort by row one since it's the only row
            If sortIndex <> 1 Then GoTo FailInput
            LibArrayTools.Sort1DArray arr, sortAscending _
                , sortTextNumberAsNumber, caseSensitiveTexts
        End If
    Case 2
        'Transpose twice for horizontal sort
        If Not verticalSort Then arr = LibArrayTools.TransposeArray(arr)
        On Error GoTo ErrorHandler
        LibArrayTools.Sort2DArray arr, sortIndex, sortAscending _
            , sortTextNumberAsNumber, caseSensitiveTexts
        If Not verticalSort Then arr = LibArrayTools.TransposeArray(arr)
    Case Else
        GoTo FailInput 'Should not happen if called from Excel
    End Select
    '
    'Replace the special value Empty with empty String so it is not returned
    '   as 0 (zero) in the caller Range
    LibArrayTools.ReplaceEmptyInArray arr, vbNullString
    '
    DM_SORT = arr
Exit Function
ErrorHandler:
FailInput:
    DM_SORT = VBA.CVErr(xlErrValue)
End Function

'###############################################################################
'Help for the Function Arguments Dialog in Excel - DM_SORT
'###############################################################################
Private Sub RegisterDMSort()
    Dim arg1 As String
    Dim arg2 As String
    Dim arg3 As String
    Dim arg4 As String
    Dim arg5 As String
    Dim arg6 As String
    '
    arg1 = "the array/range to sort"
    arg2 = "[Optional]" & vbNewLine & "the column/row to sort by. Default is 1"
    arg3 = "[Optional]" & vbNewLine & "True - Ascending (Default)" _
        & vbNewLine & "False - Descending"
    arg4 = "[Optional]" & vbNewLine _
        & "True - Sort anything that looks as a number, as a number (Default)" _
        & vbNewLine & "False - Sort numbers stored as text as text"
    arg5 = "[Optional]" & vbNewLine & "True - Case Sensitive Texts" _
        & vbNewLine & "False - Case Insensitive Texts (Default)"
    arg6 = "[Optional]" & vbNewLine & "True - Sort Vertically (Default)" _
        & vbNewLine & "False - Sort Horizontally"
    '
    Application.MacroOptions Macro:="DM_SORT" _
        , Description:="Sorts a 1D/2D Array or a 1-Area Range" _
        , ArgumentDescriptions:=Array(arg1, arg2, arg3, arg4, arg5, arg6)
End Sub
Private Sub UnregisterDMSort()
    Application.MacroOptions Macro:="DM_SORT", Description:=Empty _
        , ArgumentDescriptions:=Array(Empty, Empty, Empty, Empty, Empty, Empty)
End Sub

'*******************************************************************************
'Returns:
'   - a 1D/2D (based on input) Array of unique values
'Parameters:
'   - arr: a 1D/2D Array or Range of values to uniquify
'   - [indexes]: the column/row index(es) to be used. 0 (default) - use all
'   - [byColumn]:
'       * TRUE - uniquifies vertically (by column) (default)
'       * FALSE - uniquifies horizontally (by row)
'Notes:
'   - when uniquifing a 1D array vertically (by column), the array is regarded
'     as a 1-row 2D array and is returned as-is
'   - when uniquifing a 1D array by row, indexes must be 0, 1 or omitted (0)
'   - when uniquifing a 2D array by row, the array is first transposed, then
'     uniquified by column and finally transposed back with the end result being
'     that the array was actually uniquified by row
'   - uses LibArrayTools functions
'       * GetArrayDimsCount
'       * GetUniqueIntegers
'       * GetUniqueRows
'       * GetUniqueValues
'       * IntegerRange1D
'       * ReplaceEmptyInArray
'       * TransposeArray
'       * ValuesToCollection
'*******************************************************************************
Public Function DM_UNIQUE(ByRef arr As Variant _
    , Optional ByVal indexes As Variant = 0 _
    , Optional ByVal byColumn As Boolean = True _
) As Variant
Attribute DM_UNIQUE.VB_Description = "Returns a 1D/2D Array of unique values"
Attribute DM_UNIQUE.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.Volatile False
    #If USE_LIB_FAST_UDFS Then
        LibUDFs.TriggerFastUDFCalculation
    #End If
    '
    'Only accept 1-Area Ranges. This could alternatively be changed to ignore
    '   the extra Areas by arr = arr.Areas(1).Value2 instead of the 2 lines
    If VBA.TypeName(arr) = "Range" Then
        If arr.Areas.Count > 1 Then GoTo FailInput
        arr = arr.Value2
    End If
    '
    'Convert single value to 1-element 1D array with lower bound set to 1
    If Not VBA.IsArray(arr) Then
        Dim tempArr(1 To 1) As Variant
        tempArr(1) = arr
        arr = tempArr
    End If
    '
    Dim dimensions As Long: dimensions = LibArrayTools.GetArrayDimsCount(arr)
    Dim minIndex As Long
    Dim maxIndex As Long
    '
    'Establish the minimum and maximum allowed indexes
    Select Case dimensions
    Case 1
        If byColumn Then minIndex = LBound(arr) Else minIndex = 1
        If byColumn Then maxIndex = UBound(arr) Else maxIndex = 1
    Case 2
        If byColumn Then minIndex = LBound(arr, 2) Else minIndex = LBound(arr, 1)
        If byColumn Then maxIndex = UBound(arr, 2) Else maxIndex = UBound(arr, 1)
    Case Else
        GoTo FailInput 'Should not happen if called from Excel
    End Select
    '
    Dim useAllIndexes As Boolean
    Dim uniqueIndexes() As Long
    '
    'Check if all indexes are used
    If VBA.IsNumeric(indexes) Then useAllIndexes = (indexes = 0)
    '
    'Create array of integer indexes
    On Error GoTo ErrorHandler
    If useAllIndexes Then
        uniqueIndexes = LibArrayTools.IntegerRange1D(minIndex, maxIndex)
    Else
        uniqueIndexes = LibArrayTools.GetUniqueIntegers( _
            LibArrayTools.ValuesToCollection(indexes, nestNone, columnWise) _
            , minIndex, maxIndex)
    End If
    '
    'Get Unique Rows/Values
    If dimensions = 1 Then
        If byColumn Then
            DM_UNIQUE = arr
        Else
            DM_UNIQUE = LibArrayTools.GetUniqueValues(arr)
        End If
    Else '2 dimensions
        If Not byColumn Then arr = LibArrayTools.TransposeArray(arr)
        DM_UNIQUE = LibArrayTools.GetUniqueRows(arr, uniqueIndexes)
        If Not byColumn Then DM_UNIQUE = LibArrayTools.TransposeArray(DM_UNIQUE)
    End If
    '
    'Replace the special value Empty with empty String so it is not returned
    '   as 0 (zero) in the caller Range
    LibArrayTools.ReplaceEmptyInArray DM_UNIQUE, vbNullString
Exit Function
ErrorHandler:
FailInput:
    DM_UNIQUE = VBA.CVErr(xlErrValue)
End Function

'###############################################################################
'Help for the Function Arguments Dialog in Excel - DM_UNIQUE
'###############################################################################
Private Sub RegisterDMUnique()
    Dim arg1 As String
    Dim arg2 As String
    Dim arg3 As String
    '
    arg1 = "a 1D/2D Array or Range of values to uniquify"
    arg2 = "[Optional]" & vbNewLine & _
        " the column/row index(es) to be used. Default is 0 (use all)" _
        & vbNewLine & "Can be a list of more indexes (array/range)"
    arg3 = "[Optional]" & vbNewLine _
        & "True - uniquifies vertically (by column) (Default)" & vbNewLine _
        & "False - uniquifies horizontally (by row)"
    '
    Application.MacroOptions Macro:="DM_UNIQUE" _
        , Description:="Returns a 1D/2D Array of unique values" _
        , ArgumentDescriptions:=Array(arg1, arg2, arg3)
End Sub
Private Sub UnregisterDMUnique()
    Application.MacroOptions Macro:="DM_UNIQUE", Description:=Empty _
        , ArgumentDescriptions:=Array(Empty, Empty, Empty)
End Sub
