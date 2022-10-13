Attribute VB_Name = "LibArrayTools"
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
Option Private Module
Option Compare Text 'See Like operator in 'IsValuePassingFilter' method

'*******************************************************************************
'' Functions in this library module allow Array/Collection manipulation in VBA
'' regardless of:
''  - the host Application (Excel, Word, AutoCAD etc.)
''  - the operating system (Mac, Windows)
''  - application environment (x32, x64)
'' No extra library references are needed (e.g. Microsoft Scripting Runtime)
'' Main features:
''  - conversions: array-array, array-collection, collection-array
''  - sorting
''  - filtering
''  - reversing
''  - slicing
''  - uniquifying
''  - getting Array/Collection information
'*******************************************************************************

'' Public/Exposed methods (40+):
''    - Collection
''    - CollectionHasKey
''    - CollectionTo1DArray
''    - CollectionTo2DArray
''    - CreateFilter
''    - CreateFiltersArray
''    - Filter1DArray
''    - Filter2DArray
''    - FilterCollection (in-place)
''    - GetArrayDimsCount
''    - GetArrayElemCount
''    - GetConditionOperator
''    - GetConditionOperatorText
''    - GetUniqueIntegers
''    - GetUniqueRows
''    - GetUniqueValues
''    - InsertRowsAtIndex
''    - InsertRowsAtValChange
''    - IntegerRange1D
''    - Is2DArrayRowEmpty
''    - IsIterable
''    - IsValuePassingFilter
''    - Merge1DArrays
''    - Merge2DArrays
''    - NDArrayTo1DArray (row-wise or column-wise)
''    - NDArrayToCollections
''    - OneDArrayTo2DArray
''    - OneDArrayToCollection
''    - ReplaceEmptyInArray (in-place)
''    - Reverse1DArray (in-place)
''    - Reverse2DArray (in-place)
''    - ReverseCollection (in-place)
''    - Sequence1D
''    - Sequence2D
''    - ShallowCopyCollection
''    - Slice1DArray
''    - Slice2DArray
''    - SliceCollection
''    - Sort1DArray (in-place)
''    - Sort2DArray (in-place)
''    - SortCollection (in-place)
''    - SwapValues (in-place)
''    - TextArrayToIndex
''    - TransposeArray
''    - ValuesToCollection

'*******************************************************************************
''Note that the 'Set' keyword is not needed when assigning a Variant variable of
''  type VbVarType.vbDataObject to another Variant (e.g. v1 = v2) but it is
''  needed when Variants of type vbDataObject are returned from methods or
''  object properties (e.g. Set v1 = collection.Item(1))
'*******************************************************************************

'Used for raising errors
Private Const MODULE_NAME As String = "LibArrayTools"

'Ways of traversing an array
Public Enum ARRAY_TRAVERSE_TYPE 'In VBA, arrays are stored in columnMajorOrder
    columnWise = 0
    rowWise = 1
End Enum

'Custom Structure to describe an array Dimension
Private Type ARRAY_DIMENSION
    index As Long 'the ordinal of the dimension 1-1st, 2-2nd ...
    size As Long  'the number of elements in the dimension
    depth As Long 'the product of lower dimension (higher index) sizes
End Type

'Custom Structure to describe the dimensions and elements of an array
Private Type ARRAY_DESCRIPTOR
    dimsCount As Long
    dimensions() As ARRAY_DIMENSION
    elemCount As Long
    elements1D As Variant
    traverseType As ARRAY_TRAVERSE_TYPE
    rowMajorIndexes() As Long
End Type

'Vector type (see QuickSortVector method)
Private Enum VECTOR_TYPE
    vecArray = 1
    vecCollection = 2
End Enum

'Rank for categorizing different data types (used for comparing values/sorting)
Private Enum DATA_TYPE_RANK
    rankEmpty = 1
    rankUDT = 2
    rankObject = 3
    rankArray = 4
    rankNull = 5
    rankError = 6
    rankBoolean = 7
    rankText = 8
    rankNumber = 9
End Enum

'The result of a comparison between 2 values (for the purpose of sorting)
Private Type COMPARE_RESULT
    mustSwap As Boolean
    areEqual As Boolean
End Type

'Struct used for passing compare-related options between relevant methods
Private Type COMPARE_OPTIONS
    compAscending As Boolean
    useTextNumberAsNumber As Boolean
    compareMethod As VbCompareMethod
End Type

'Struct used for storing a quick sort pivot's value and corresponding index
Private Type SORT_PIVOT
    index As Long
    value_ As Variant
End Type

'Available Operators for testing conditions (see FILTER_PAIR struct)
Public Enum CONDITION_OPERATOR
    opNone = 0
    opEqual = 1
    [_opMin] = 1
    opSmaller = 2
    opBigger = 3
    opSmallerOrEqual = 4
    opBiggerOrEqual = 5
    opNotEqual = 6
    opin = 7
    opNotIn = 8
    opLike = 9
    opNotLike = 10
    [_opMax] = 10
End Enum

'Struct used for filtering (see FILTER_PAIR struct)
Private Type COMPARE_VALUE
    value_ As Variant
    rank_ As DATA_TYPE_RANK
    isIterable_ As Boolean
    textKeys_ As Collection 'used when 'value_' is iterable (Array/Collection)
End Type

'Struct used for filtering (see Filter2DArray method)
'*******************************************************************************
'If used as a parameter in a class method or as a return value of a class method
'   then make sure to declare the class method as Friend. Otherwise the only
'   other way to make the code compile is to remove the Option Private Module
'   from the top of this module. Note that removing Option Private Module would
'   expose the methods of this module (for example in Excel they can be seen as
'   custom functions in the Excel interface - which is undesirable as they are
'   not intended as UDFs)
'*******************************************************************************
Public Type FILTER_PAIR
    cOperator As CONDITION_OPERATOR
    compValue As COMPARE_VALUE
End Type

'Struct used in ValuesToCollection
Public Enum NESTING_TYPE
    nestNone = 0
    [_nMin] = nestNone
    nestMultiItemsOnly = 1
    nestAll = 2
    [_nMax] = 2
End Enum

'vbLongLong is a VbVarType available in x64 systems only
'Create the value for x32 for convenience in writing Select Case logic
#If Mac Then
    Const vbLongLong As Long = 20 'Apparently missing for x64 on Mac
#Else
    #If Win64 = 0 Then
        Const vbLongLong As Long = 20
    #End If
#End If

'Structs needed for ZeroLengthArray method
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type TagVariant
    vt As Integer
    wReserved1 As Integer
    wReserved2 As Integer
    wReserved3 As Integer
    #If VBA7 Then
        ptr As LongPtr
    #Else
        ptr As Long
    #End If
End Type

'Win APIs needed for ZeroLengthArray method
#If Mac Then
#Else
    #If VBA7 Then
        Private Declare PtrSafe Function SafeArrayCreate Lib "OleAut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As SAFEARRAYBOUND) As LongPtr
        Private Declare PtrSafe Function VariantCopy Lib "OleAut32.dll" (pvargDest As Any, pvargSrc As Any) As Long
        Private Declare PtrSafe Function SafeArrayDestroy Lib "OleAut32.dll" (ByVal psa As LongPtr) As Long
    #Else
        Private Declare Function SafeArrayCreate Lib "OleAut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As SAFEARRAYBOUND) As Long
        Private Declare Function VariantCopy Lib "OleAut32.dll" (pvargDest As Variant, pvargSrc As Any) As Long
        Private Declare Function SafeArrayDestroy Lib "OleAut32.dll" (ByVal psa As Long) As Long
    #End If
#End If

'*******************************************************************************
'Returns a new collection containing the specified values
'Similar with [_HiddenModule].Array but returns a collection
'Parameters:
'   - values: a ParamArray Variant containing values to be added to collection
'Does not raise errors
'Examples:
'   - Collection() returns [] (a New empty collection)
'   - Collection(1,2,3) returns [1,2,3] (a collection with 3 integers)
'   - Collection(1,2,Collection(3,4)) returns [1,2,[3,4]]
'*******************************************************************************
Public Function Collection(ParamArray values() As Variant) As Collection
    Dim v As Variant
    Dim coll As Collection
    '
    Set coll = New Collection
    For Each v In values
        coll.Add v
    Next v
    '
    Set Collection = coll
End Function

'*******************************************************************************
'Returns a boolean indicating if a Collection has a specific key
'Parameters:
'   - coll: a collection to check for key
'   - keyValue: the key being searched for
'Does not raise errors
'*******************************************************************************
Public Function CollectionHasKey(ByVal coll As Collection _
                               , ByRef keyValue As String) As Boolean
    On Error Resume Next
    coll.Item keyValue
    CollectionHasKey = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Returns a 1D array based on values contained in the specified collection
'Parameters:
'   - coll: collection that contains the values to be used
'   - [outLowBound]: the start index of the result array. Default is 0
'Raises error:
'   - 91: if Collection Object is not set
'*******************************************************************************
Public Function CollectionTo1DArray(ByVal coll As Collection _
                                  , Optional ByVal outLowBound As Long = 0) As Variant()
    Const fullMethodName As String = MODULE_NAME & ".CollectionTo1DArray"
    '
    'Check Input
    If coll Is Nothing Then
        Err.Raise 91, fullMethodName, "Collection not set"
    ElseIf coll.Count = 0 Then
        CollectionTo1DArray = ZeroLengthArray()
        Exit Function
    End If
    '
    Dim res() As Variant: ReDim res(outLowBound To outLowBound + coll.Count - 1)
    Dim i As Long: i = outLowBound
    Dim v As Variant
    '
    'Populate array
    For Each v In coll
        If IsObject(v) Then Set res(i) = v Else res(i) = v
        i = i + 1
    Next v
    '
    CollectionTo1DArray = res
End Function

'*******************************************************************************
'Returns a 2D array based on values contained in the specified collection
'Parameters:
'   - coll: collection that contains the values to be used
'   - columnsCount: the number of columns that the result 2D array will have
'   - [outLowRow]: the start index of the result's 1st dimension. Default is 0
'   - [outLowCol]: the start index of the result's 2nd dimension. Default is 0
'Raises error:
'   - 91: if Collection Object is not set
'   -  5: if the number of columns is less than 1
'Notes:
'   - if the total Number of values is not divisible by columnsCount then the
'     extra values (last row) of the array are by default the value Empty
'Examples:
'   - coll = [1,2,3,4] and columnsCount = 1 >> returns [1]
'                                                      [2]
'                                                      [3]
'                                                      [4]   (4 rows 1 column)
'   - coll = [1,2,3,4] and columnsCount = 2 >> returns [1,2]
'                                                      [3,4] (2 rows 2 columns)
'   - coll = [1,2,3]   and columnsCount = 2 >> returns [1,2]
'                                                      [3,Empty]
'*******************************************************************************
Public Function CollectionTo2DArray(ByVal coll As Collection _
                                  , ByVal columnsCount As Long _
                                  , Optional ByVal outLowRow As Long = 0 _
                                  , Optional ByVal outLowCol As Long = 0) As Variant()
    Const fullMethodName As String = MODULE_NAME & ".CollectionTo2DArray"
    '
    'Check Input
    If coll Is Nothing Then
        Err.Raise 91, fullMethodName, "Collection not set"
    ElseIf columnsCount < 1 Then
        Err.Raise 5, fullMethodName, "Invalid Columns Count"
    ElseIf coll.Count = 0 Then
        CollectionTo2DArray = ZeroLengthArray()
        Exit Function
    End If
    '
    Dim rowsCount As Long: rowsCount = -Int(-coll.Count / columnsCount)
    Dim arr() As Variant: ReDim arr(outLowRow To outLowRow + rowsCount - 1 _
                                  , outLowCol To outLowCol + columnsCount - 1)
    Dim i As Long: i = 0
    Dim r As Long
    Dim c As Long
    Dim v As Variant
    '
    'Populate array
    For Each v In coll
        r = outLowRow + i \ columnsCount
        c = outLowCol + i Mod columnsCount
        If IsObject(v) Then Set arr(r, c) = v Else arr(r, c) = v
        i = i + 1
    Next v
    '
    CollectionTo2DArray = arr
End Function

'*******************************************************************************
'Creates a FILTER_PAIR struct from values
'Parameters:
'   - cOperator - see CONDITION_OPERATOR enum and GetConditionOperator method
'   - compareValue: any value
'Notes:
'   - see 'GetConditionOperator' for the operator conversion
'*******************************************************************************
Public Function CreateFilter(ByVal cOperator As CONDITION_OPERATOR _
                           , ByVal compareValue As Variant) As FILTER_PAIR
    If cOperator >= [_opMin] And cOperator <= [_opMax] Then
        CreateFilter.cOperator = cOperator
    End If
    With CreateFilter.compValue
        .rank_ = GetDataTypeRank(compareValue)
        If .rank_ = rankArray Or .rank_ = rankObject Then
            .isIterable_ = IsIterable(compareValue)
            If .isIterable_ Then Set .textKeys_ = CreateTextKeys(compareValue)
        End If
        If .rank_ = rankObject Then
            Set .value_ = compareValue
        Else
            .value_ = compareValue
        End If
    End With
End Function

'*******************************************************************************
'Creates a collection with text keys corresponding to the received values.
'Utility for 'CreateFilter' method
'*******************************************************************************
Private Function CreateTextKeys(ByRef values As Variant) As Collection
    Dim collResult As Collection: Set collResult = New Collection
    Dim keyValue As String
    Dim v As Variant
    '
    On Error Resume Next 'Ignore duplicates
    For Each v In values
        keyValue = GetUniqueTextKey(v)
        If LenB(keyValue) > 0 Then collResult.Add keyValue, keyValue
    Next v
    On Error GoTo 0
    Set CreateTextKeys = collResult
End Function

'*******************************************************************************
'Creates an array of FILTER_PAIR structs from a variant of valuePairs
'Parameters:
'   - valuePairs: a ParamArray Variant containing any number of filter pairs:
'       * operator:
'           + textOperator (see 'GetConditionOperator' method):
'               ~ comparison operators: =, <, >, <=, >=, <>
'               ~ inclusion operators: IN , NOT IN
'               ~ pattern matching operators: LIKE, NOT LIKE
'           + enumOperator (see CONDITION_OPERATOR enum)
'       * compareValue: any value
'Raises error:
'   - 5 if:
'       * no value is provided
'       * number of elements in 'valuePairs' is not divisible by 2
'       * operator is invalid (wrong data type or not supported)
'Notes:
'   - this method can be used to quickly create an array of FILTER_PAIRs to be
'     used with the 'Filter...' methods
'*******************************************************************************
Public Function CreateFiltersArray(ParamArray valuePairs() As Variant) As FILTER_PAIR()
    Const fullMethodName As String = MODULE_NAME & ".CreateFiltersArray"
    '
    Dim collFilterPairs As Collection
    Dim v As Variant: v = valuePairs
    '
    'Check Input
    Set collFilterPairs = ValuesToCollection(v, nestMultiItemsOnly, rowWise)
    If collFilterPairs.Count = 0 Then
        Err.Raise 5, fullMethodName, "Expected at least one filter"
    ElseIf collFilterPairs.Count Mod 2 <> 0 Then
        Err.Raise 5, fullMethodName, "Expected filter pairs of operator & value"
    End If
    '
    Dim arr() As FILTER_PAIR: ReDim arr(0 To collFilterPairs.Count / 2 - 1)
    Dim filter As FILTER_PAIR
    Dim cOperator As CONDITION_OPERATOR
    Dim isOperator As Boolean: isOperator = True
    Dim i As Long:             i = 0
    '
    For Each v In collFilterPairs
        If isOperator Then
            Select Case VarType(v)
                Case vbLong:   cOperator = v
                Case vbString: cOperator = GetConditionOperator(v)
                Case Else:     cOperator = opNone
            End Select
        Else
            filter = CreateFilter(cOperator, v)
            If filter.cOperator = opNone Then
                Err.Raise 5, fullMethodName, "Invalid operator"
            End If
            arr(i) = filter
            i = i + 1
        End If
        isOperator = Not isOperator
    Next v
    '
    CreateFiltersArray = arr
End Function

'*******************************************************************************
'Returnes the number of dimensions for an array of FILTER_PAIR UDTs
'Note that the 'GetArrayDimsCount' method cannot be used because UDTs canot be
'   assigned to a Variant
'Utility for 'Filter...' methods
'*******************************************************************************
Private Function GetFiltersDimsCount(ByRef arr() As FILTER_PAIR) As Long
    Const MAX_DIMENSION As Long = 60
    Dim dimension As Long
    Dim tempBound As Long
    '
    On Error GoTo FinalDimension
    For dimension = 1 To MAX_DIMENSION
        tempBound = LBound(arr, dimension)
    Next dimension
FinalDimension:
    GetFiltersDimsCount = dimension - 1
End Function

'*******************************************************************************
'Filters a 1D Array
'Parameters:
'   - arr: a 1D array to be filtered
'   - filters: an array of FILTER_PAIR structs (operator/compareValue pairs)
'   - [outLowBound]: the start index of the result array. Default is 0
'Raises Error:
'   - 5 if:
'       * 'arr' is not 1D
'       * values are incompatible (see 'IsValuePassingFilter' method)
'Notes:
'   - use the 'CreateFiltersArray' method to quickly create filters
'Examples:
'   - arr = [1,2,3,4] and filters = [">",2,"<=",4] >> returns [3,4]
'   - arr = [1,3,6,9] and filters = ["IN",[1,2,3,4]] >> returns [1,3]
'   - arr = ["test","hes","et"] and filters = ["LIKE","*es?"] > returns ["test"]
'*******************************************************************************
Public Function Filter1DArray(ByRef arr As Variant _
                            , ByRef filters() As FILTER_PAIR _
                            , Optional ByVal outLowBound As Long = 0) As Variant()
    Const fullMethodName As String = MODULE_NAME & ".Filter1DArray"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 1 Then
        Err.Raise 5, fullMethodName, "Expected 1D Array for filtering"
    ElseIf GetFiltersDimsCount(filters) <> 1 Then
        Err.Raise 5, fullMethodName, "Expected 1D Array of filters"
    End If
    '
    Dim i As Long
    Dim collIndexes As New Collection
    '
    'Create collecton of indexes with keys for easy removal
    For i = LBound(arr, 1) To UBound(arr, 1)
        collIndexes.Add i, CStr(i)
    Next
    '
    Dim filter As FILTER_PAIR
    Dim v As Variant
    '
    'Remove indexes for values that do NOT pass filters
    On Error GoTo ErrorHandler
    For i = LBound(filters, 1) To UBound(filters, 1)
        filter = filters(i)
        For Each v In collIndexes
            If Not IsValuePassingFilter(arr(v), filter) Then
                collIndexes.Remove CStr(v)
            End If
        Next v
        If collIndexes.Count = 0 Then
            Filter1DArray = ZeroLengthArray()
            Exit Function
        End If
    Next i
    '
    Dim res() As Variant
    Dim j As Long
    '
    'Copy values to the result array
    ReDim res(outLowBound To outLowBound + collIndexes.Count - 1)
    i = outLowBound
    For Each v In collIndexes
        j = CLng(v)
        If IsObject(arr(j)) Then Set res(i) = arr(j) Else res(i) = arr(j)
        i = i + 1
    Next v
    '
    Filter1DArray = res
Exit Function
ErrorHandler:
    Err.Raise Err.Number _
            , Err.Source & vbNewLine & fullMethodName _
            , Err.Description & vbNewLine & "Invalid filters or values"
End Function

'*******************************************************************************
'Filters a 2D Array by a specified column
'Parameters:
'   - arr: a 2D array to be filtered
'   - byColumn: the index of the column used for filtering
'   - filters: an array of FILTER_PAIR (operator/compareValue pairs)
'   - [outLowRow]: start index of the result array's 1st dimension. Default is 0
'Raises Error:
'   - 5 if:
'       * 'arr' is not 2D
'       * 'filters' is not 1D
'       * 'byColumn' is out of bounds
'       * values are incompatible (see 'IsValuePassingFilter' method)
'Notes:
'   - column lower bound is preserved (same as the original array)
'*******************************************************************************
Public Function Filter2DArray(ByRef arr As Variant _
                            , ByVal byColumn As Long _
                            , ByRef filters() As FILTER_PAIR _
                            , Optional ByVal outLowRow As Long = 0) As Variant()
    Const fullMethodName As String = MODULE_NAME & ".Filter2DArray"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 2 Then
        Err.Raise 5, fullMethodName, "Expected 2D Array for filtering"
    ElseIf byColumn < LBound(arr, 2) Or byColumn > UBound(arr, 2) Then
        Err.Raise 5, fullMethodName, "Invalid column index"
    ElseIf GetFiltersDimsCount(filters) <> 1 Then
        Err.Raise 5, fullMethodName, "Expected 1D Array of filters"
    End If
    '
    Dim i As Long
    Dim collRows As New Collection
    '
    'Create collecton of row indexes with keys for easy removal
    For i = LBound(arr, 1) To UBound(arr, 1)
        collRows.Add i, CStr(i)
    Next
    '
    Dim filter As FILTER_PAIR
    Dim v As Variant
    '
    'Remove row indexes for values that do NOT pass filters
    On Error GoTo ErrorHandler
    For i = LBound(filters, 1) To UBound(filters, 1)
        filter = filters(i)
        For Each v In collRows
            If Not IsValuePassingFilter(arr(v, byColumn), filter) Then
                collRows.Remove CStr(v)
            End If
        Next v
        If collRows.Count = 0 Then
            Filter2DArray = ZeroLengthArray()
            Exit Function
        End If
    Next i
    '
    Dim lowCol As Long: lowCol = LBound(arr, 2)
    Dim uppCol As Long: uppCol = UBound(arr, 2)
    Dim res() As Variant: ReDim res(outLowRow To outLowRow + collRows.Count - 1 _
                                  , lowCol To uppCol)
    Dim needSet As Boolean
    Dim r As Long
    Dim j As Long
    '
    'Copy rows to the result array
    i = outLowRow
    For Each v In collRows
        r = CLng(v)
        For j = lowCol To uppCol
            needSet = IsObject(arr(r, j))
            If needSet Then Set res(i, j) = arr(r, j) Else res(i, j) = arr(r, j)
        Next j
        i = i + 1
    Next v
    '
    Filter2DArray = res
Exit Function
ErrorHandler:
    Err.Raise Err.Number _
            , Err.Source & vbNewLine & fullMethodName _
            , Err.Description & vbNewLine & "Invalid filters or values"
End Function

'*******************************************************************************
'Filters a Collection in-place
'Parameters:
'   - coll: a collection to be filtered
'   - filters: an array of FILTER_PAIR (operator/compareValue pairs)
'Notes:
'   - the collection is modified in place so it is optional to use the return
'     value of the function
'Raises Error:
'   - 91: if Collection Object is not set
'   -  5 if:
'        * 'filters' is not 1D
'        * values are incompatible (see 'IsValuePassingFilter' method)
'*******************************************************************************
Public Function FilterCollection(ByVal coll As Collection _
                               , ByRef filters() As FILTER_PAIR) As Collection
    Const fullMethodName As String = MODULE_NAME & ".FilterCollection"
    '
    'Check Input
    If coll Is Nothing Then
        Err.Raise 91, fullMethodName, "Collection not set"
    ElseIf GetFiltersDimsCount(filters) <> 1 Then
        Err.Raise 5, fullMethodName, "Expected 1D Array of filters"
    ElseIf coll.Count = 0 Then
        Set FilterCollection = coll
        Exit Function
    End If
    '
    Dim filter As FILTER_PAIR
    Dim v As Variant
    Dim i As Long
    Dim j As Long
    '
    'Remove values that do NOT pass filters
    On Error GoTo ErrorHandler
    For i = LBound(filters, 1) To UBound(filters, 1)
        filter = filters(i)
        j = 1
        For Each v In coll
            If IsValuePassingFilter(v, filter) Then
                j = j + 1
            Else
                coll.Remove j
            End If
        Next v
    Next i
    '
    Set FilterCollection = coll 'Useful for method chaining
Exit Function
ErrorHandler:
    Err.Raise Err.Number _
            , Err.Source & vbNewLine & fullMethodName _
            , Err.Description & vbNewLine & "Invalid filters or values"
End Function

'*******************************************************************************
'Returns the Number of dimensions for an input array
'Returns 0 if array is uninitialized or input not an array
'Note that a zero-length array has 1 dimension! Ex. Array() bounds are (0 to -1)
'*******************************************************************************
Public Function GetArrayDimsCount(ByRef arr As Variant) As Long
    Const MAX_DIMENSION As Long = 60 'VB limit
    Dim dimension As Long
    Dim tempBound As Long
    '
    On Error GoTo FinalDimension
    For dimension = 1 To MAX_DIMENSION
        tempBound = LBound(arr, dimension)
    Next dimension
FinalDimension:
    GetArrayDimsCount = dimension - 1
End Function

'*******************************************************************************
'Returns the Number of elements for an input array
'Returns 0 if array is uninitialized/zero-length or if input is not an array
'*******************************************************************************
Public Function GetArrayElemCount(ByRef arr As Variant) As Long
    On Error Resume Next
    GetArrayElemCount = GetArrayDescriptor(arr, rowWise, False).elemCount
    On Error GoTo 0
End Function

'*******************************************************************************
'Converts a string representation of an operator to it's corresponding Enum:
'   * comparison operators: =, <, >, <=, >=, <>
'   * inclusion operators: IN, NOT IN
'   * pattern matching operators: LIKE, NOT LIKE
'A Static Keyed Collection is used for fast retrieval (instead of Select Case)
'Does not raise errors
'*******************************************************************************
Public Function GetConditionOperator(ByVal textOperator As String) As CONDITION_OPERATOR
    Static collOperators As Collection
    '
    If collOperators Is Nothing Then
        Set collOperators = New Collection
        Dim i As Long
        For i = [_opMin] To [_opMax]
            collOperators.Add i, GetConditionOperatorText(i)
        Next i
    End If
    '
    On Error Resume Next
    GetConditionOperator = collOperators.Item(textOperator)
    On Error GoTo 0
End Function

'*******************************************************************************
'Converts a CONDITION_OPERATOR enum value to it's string representation
'   * comparison operators: =, <, >, <=, >=, <>
'   * inclusion operators: IN , NOT IN
'   * pattern matching operators: LIKE, NOT LIKE
'*******************************************************************************
Public Function GetConditionOperatorText(ByVal cOperator As CONDITION_OPERATOR) As String
    Static arrOperators([_opMin] To [_opMax]) As String
    Static isSet As Boolean
    '
    If Not isSet Then
        arrOperators(opEqual) = "="
        arrOperators(opSmaller) = "<"
        arrOperators(opBigger) = ">"
        arrOperators(opSmallerOrEqual) = "<="
        arrOperators(opBiggerOrEqual) = ">="
        arrOperators(opNotEqual) = "<>"
        arrOperators(opin) = "IN"
        arrOperators(opNotIn) = "NOT IN"
        arrOperators(opLike) = "LIKE"
        arrOperators(opNotLike) = "NOT LIKE"
        isSet = True
    End If
    '
    If cOperator < [_opMin] Or cOperator > [_opMax] Then Exit Function
    GetConditionOperatorText = arrOperators(cOperator)
End Function

'*******************************************************************************
'Receives an iterable list of integers via a variant and returns a 1D array
'   containing all the unique integer values within specified limits
'Parameters:
'   - iterableList: an array, collection or other object that can be iterated
'                   using a For Each... Next loop
'   - [minValue]: the minimum integer value allowed. Default is -2147483648
'   - [maxValue]: the maximum integer value allowed. Default is +2147483647
'Raises error:
'   - 5 if:
'       * 'iterableList' does not support For Each... Next loop
'       * if any value from list is not numeric
'       * if any value from list is outside specified limits
'Notes:
'   - numbers with decimal places are floored using the Int function
'*******************************************************************************
Public Function GetUniqueIntegers(ByRef iterableList As Variant _
                                , Optional ByVal minAllowed As Long = &H80000000 _
                                , Optional ByVal maxAllowed As Long = &H7FFFFFFF) As Long()
    Const fullMethodName As String = MODULE_NAME & ".GetUniqueIntegers"
    '
    'Check Input
    If Not IsIterable(iterableList) Then
        Err.Raise 5, fullMethodName, "Variant doesn't support For Each... Next"
    ElseIf minAllowed > maxAllowed Then
        'Swapping minAllowed with maxAllowed could lead to unwanted results
        Err.Raise 5, fullMethodName, "Invalid limits"
    End If
    '
    Dim v As Variant
    Dim collUnique As New Collection
    '
    On Error Resume Next 'Ignore duplicates
    For Each v In iterableList
        'Check data type and floor numbers with decimal places
        Select Case VarType(v)
        Case vbByte, vbInteger, vbLong, vbLongLong
            'Integer. Do nothing
        Case vbCurrency, vbDecimal, vbDouble, vbSingle, vbDate
            v = Int(v)
        Case Else
            On Error GoTo 0
            Err.Raise 5, fullMethodName, "Invalid data type. Expected numeric"
        End Select
        '
        If v < minAllowed Or v > maxAllowed Then
            On Error GoTo 0
            Err.Raise 5, fullMethodName, "Value is outside limits"
        End If
        collUnique.Add v, CStr(v)
    Next v
    On Error GoTo 0
    '
    If collUnique.Count = 0 Then Exit Function
    '
    'Copy unique integers to a result array
    Dim res() As Long: ReDim res(0 To collUnique.Count - 1)
    Dim i As Long: i = 0
    '
    For Each v In collUnique
        res(i) = v
        i = i + 1
    Next v
    '
    GetUniqueIntegers = res
End Function

'*******************************************************************************
'Receives a 2D Array and returns unique rows based on chosen columns
'Parameters:
'   - arr: a 2D array
'   - columns_: an array of one or more column indexes to be used
'   - [outLowRow]: start index of the result array's 1st dimension. Default is 0
'Raises error:
'   - 5:
'       * 'arr' is not a 2D array
'       * column indexes are out of bounds
'Notes:
'   - column lower bound is preserved (same as the original array)
'*******************************************************************************
Public Function GetUniqueRows(ByRef arr As Variant _
                            , ByRef columns_() As Long _
                            , Optional ByVal outLowRow As Long = 0) As Variant()
    Const fullMethodName As String = MODULE_NAME & ".GetUniqueRows"
    '
    'Check Input Array
    If GetArrayDimsCount(arr) <> 2 Then
        Err.Raise 5, fullMethodName, "Expected 2D Array of values"
    ElseIf GetArrayDimsCount(columns_) <> 1 Then
        Err.Raise 5, fullMethodName, "Expected 1D Array of column indexes"
    End If
    '
    Dim lowerCol As Long: lowerCol = LBound(arr, 2)
    Dim upperCol As Long: upperCol = UBound(arr, 2)
    Dim v As Variant
    '
    'Check column indexes
    For Each v In columns_
        If v < lowerCol Or v > upperCol Then
            Err.Raise 5, fullMethodName, "Invalid column index"
        End If
    Next v
    '
    Dim collRows As New Collection
    Dim rowKey As String
    Dim i As Long
    '
    'Create a collection of indexes corresponding to the unique rows
    On Error Resume Next 'Ignore duplicate rows
    For i = LBound(arr, 1) To UBound(arr, 1)
        rowKey = GetRowKey(arr, i, columns_)
        If LenB(rowKey) > 0 Then collRows.Add i, rowKey
    Next i
    On Error GoTo 0
    '
    Dim res() As Variant: ReDim res(outLowRow To outLowRow + collRows.Count - 1 _
                                  , lowerCol To upperCol)
    Dim j As Long
    Dim r As Long
    Dim needsSet As Boolean
    '
    'Copy rows to the result array
    i = outLowRow
    For Each v In collRows
        r = CLng(v)
        For j = lowerCol To upperCol
            needsSet = IsObject(arr(r, j))
            If needsSet Then Set res(i, j) = arr(r, j) Else res(i, j) = arr(r, j)
        Next j
        i = i + 1
    Next v
    GetUniqueRows = res
End Function

'*******************************************************************************
'Returns a string key for an array row and indicated columns
'Utility for 'GetUniqueRows' method
'*******************************************************************************
Private Function GetRowKey(ByRef arr As Variant _
                         , ByRef rowIndex As Long _
                         , ByRef columns_() As Long) As String
    Dim colIndex As Variant
    Dim rowKey As String
    Dim keyValue As String
    '
    For Each colIndex In columns_
        keyValue = GetUniqueTextKey(arr(rowIndex, colIndex))
        If LenB(keyValue) = 0 Then Exit Function 'Ignore rows with Arrays/UDTs
        rowKey = rowKey & keyValue 'No need for a separator. See GetUniqueTextKey
    Next colIndex
    GetRowKey = rowKey
End Function

'*******************************************************************************
'Receives an iterable list of values via a variant and returns a 1D array
'   containing all the unique values
'Parameters:
'   - iterableList: an array, collection or other object that can be iterated
'                   using a For Each... Next loop
'   - [outLowBound]: the start index of the result array. Default is 0
'Raises error:
'   - 5: if 'iterableList' does not support For Each... Next loop
'*******************************************************************************
Public Function GetUniqueValues(ByRef iterableList As Variant _
                              , Optional ByVal outLowBound As Long = 0) As Variant()
    Const fullMethodName As String = MODULE_NAME & ".GetUniqueValues"
    '
    'Check Input
    If Not IsIterable(iterableList) Then
        Err.Raise 5, fullMethodName, "Variant doesn't support For Each... Next"
    End If
    '
    Dim v As Variant
    Dim collUnique As New Collection
    Dim keyValue As String
    '
    On Error Resume Next 'Ignore duplicates
    For Each v In iterableList
        keyValue = GetUniqueTextKey(v) 'Ignores Arrays and UDTs
        If LenB(keyValue) > 0 Then collUnique.Add v, keyValue
    Next v
    On Error GoTo 0
    GetUniqueValues = CollectionTo1DArray(collUnique, outLowBound)
End Function

'*******************************************************************************
'Returns a unique key for a Variant value by combining the value and data type
'   - For objects it uses the pointer returned by ObjPtr (using base interface)
'   - Arrays and User Defined Types (UDTs) are not supported (returns "")
'   - For other types it distinguishes by adding a trailing value based on type
'*******************************************************************************
Private Function GetUniqueTextKey(ByRef v As Variant) As String
    Dim obj As IUnknown 'the fundamental interface in COM
    '
    If IsObject(v) Then
        Set obj = v
        GetUniqueTextKey = ObjPtr(obj)
        Exit Function
    End If
    '
    Select Case VarType(v)
    Case vbNull:    GetUniqueTextKey = "Null_0"
    Case vbEmpty:   GetUniqueTextKey = "Empty_1"
    Case vbError:   GetUniqueTextKey = CStr(v) & "_2"
    Case vbBoolean: GetUniqueTextKey = CStr(v) & "_3"
    Case vbString:  GetUniqueTextKey = CStr(v) & "_4"
    Case vbDate:    GetUniqueTextKey = CStr(CDbl(v)) & "_5"
    Case vbByte, vbInteger, vbLong, vbLongLong 'Integer
        GetUniqueTextKey = CStr(v) & "_5"
    Case vbCurrency, vbDecimal, vbDouble, vbSingle 'Decimal-point
        GetUniqueTextKey = CStr(v) & "_5"
    Case vbDataObject
        Set obj = v
        GetUniqueTextKey = ObjPtr(obj)
    Case Else: Exit Function 'Array/UDT
    End Select
End Function

'*******************************************************************************
'Inserts rows in a 2D array before the specified row
'Parameters:
'   - arr: a 2D array to insert into
'   - rowsCount: the number of rows to insert
'   - beforeRow: the index of the row before which rows will be inserted
'Raises error:
'   - 5 if:
'       * array is not two-dimensional
'       * beforeRow index or rowsCount is invalid
'*******************************************************************************
Public Function InsertRowsAtIndex(ByRef arr As Variant _
                                , ByVal rowsCount As Long _
                                , ByVal beforeRow As Long) As Variant
    Const fullMethodName As String = MODULE_NAME & ".InsertRowsAtIndex"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 2 Then
        Err.Raise 5, fullMethodName, "Array is not two-dimensional"
    ElseIf beforeRow < LBound(arr, 1) Or beforeRow > UBound(arr, 1) + 1 Then
        Err.Raise 5, fullMethodName, "Invalid beforeRow index"
    ElseIf rowsCount < 0 Then
        Err.Raise 5, fullMethodName, "Invalid rows count"
    ElseIf rowsCount = 0 Then
        InsertRowsAtIndex = arr
        Exit Function
    End If
    '
    'Store Bounds for the input array
    Dim loRowBound As Long: loRowBound = LBound(arr, 1)
    Dim upRowBound As Long: upRowBound = UBound(arr, 1)
    Dim loColBound As Long: loColBound = LBound(arr, 2)
    Dim upColBound As Long: upColBound = UBound(arr, 2)
    '
    Dim res() As Variant
    Dim i As Long
    Dim j As Long
    Dim newRow As Long
    Dim v As Variant
    '
    'Create a new array with the required rows
    ReDim res(loRowBound To upRowBound + rowsCount, loColBound To upColBound)
    '
    'Copy values to the result array
    i = loRowBound
    j = loColBound
    For Each v In arr
        If i < beforeRow Then newRow = i Else newRow = i + rowsCount
        If IsObject(v) Then Set res(newRow, j) = v Else res(newRow, j) = v
        If i = upRowBound Then 'Switch to the next column
            j = j + 1
            i = loRowBound
        Else
            i = i + 1
        End If
    Next v
    '
    InsertRowsAtIndex = res
End Function

'*******************************************************************************
'Inserts rows in a 2D Array between rows with different values (on the specified
'   column) and optionally at the top and/or bottom of the array
'Parameters:
'   - arr: a 2D array to insert into
'   - rowsCount: the number of rows to insert
'   - columnIndex: the index of the column used for row comparison
'   - [topRowsCount]: number of rows to insert before array. Default is 0
'   - [bottomRowsCount]: number of rows to insert after array. Default is 0
'Raises error:
'   - 5 if:
'       * array is not two-dimensional
'       * columnIndex/rowsCount/topRowsCount/bottomRowsCount is invalid
'*******************************************************************************
Public Function InsertRowsAtValChange(ByRef arr As Variant _
                                    , ByVal rowsCount As Long _
                                    , ByVal columnIndex As Long _
                                    , Optional ByVal topRowsCount As Long = 0 _
                                    , Optional ByVal bottomRowsCount As Long = 0) As Variant
    Const fullMethodName As String = MODULE_NAME & ".InsertRowsAtValChange"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 2 Then
        Err.Raise 5, fullMethodName, "Expected 2D Array"
    ElseIf columnIndex < LBound(arr, 2) Or columnIndex > UBound(arr, 2) Then
        Err.Raise 5, fullMethodName, "Invalid column index for comparison"
    ElseIf rowsCount < 0 Or topRowsCount < 0 Or bottomRowsCount < 0 Then
        Err.Raise 5, fullMethodName, "Invalid rows count"
    ElseIf rowsCount = 0 And topRowsCount = 0 And bottomRowsCount = 0 Then
        InsertRowsAtValChange = arr
        Exit Function
    End If
    '
    'Store Bounds for the input array
    Dim lowRow As Long: lowRow = LBound(arr, 1)
    Dim uppRow As Long: uppRow = UBound(arr, 1)
    Dim lowCol As Long: lowCol = LBound(arr, 2)
    Dim uppCol As Long: uppCol = UBound(arr, 2)
    '
    Dim arrNewRows() As Long: ReDim arrNewRows(lowRow To uppRow)
    Dim i As Long
    Dim n As Long
    Dim currentKey As String
    Dim previousKey As String
    Dim rowsToInsert As Long
    '
    'Store new row indexes
    previousKey = GetUniqueTextKey(arr(lowRow, columnIndex))
    n = lowRow + topRowsCount
    For i = lowRow To uppRow
        currentKey = GetUniqueTextKey(arr(i, columnIndex))
        If currentKey <> previousKey Then n = n + rowsCount
        arrNewRows(i) = n
        '
        n = n + 1
        previousKey = currentKey
    Next i
    '
    rowsToInsert = arrNewRows(uppRow) - uppRow + bottomRowsCount
    If rowsToInsert = 0 Then
        InsertRowsAtValChange = arr
        Exit Function
    End If
    '
    Dim res() As Variant
    ReDim res(lowRow To uppRow + rowsToInsert, lowCol To uppCol)
    Dim j As Long
    Dim needSet As Boolean
    '
    'Copy values to the result array
    For i = lowRow To uppRow
        n = arrNewRows(i)
        For j = lowCol To uppCol
            needSet = IsObject(arr(i, j))
            If needSet Then Set res(n, j) = arr(i, j) Else res(n, j) = arr(i, j)
        Next j
    Next i
    '
    InsertRowsAtValChange = res
End Function

'*******************************************************************************
'Returns a 1D array of consecutive Long Integer values
'Parameters:
'   - startValue: the first value
'   - endValue: the last value
'   - [outLowBound]: the start index of the result array. Default is 0
'Does not raise errors
'*******************************************************************************
Public Function IntegerRange1D(ByVal startValue As Long _
                             , ByVal endValue As Long _
                             , Optional ByVal outLowBound As Long = 0) As Long()
    Dim diff As Long:  diff = endValue - startValue
    Dim arr() As Long: ReDim arr(outLowBound To outLowBound + Math.Abs(diff))
    Dim step_ As Long: step_ = Math.Sgn(diff)
    Dim i As Long
    Dim v As Long: v = startValue
    '
    For i = LBound(arr) To UBound(arr)
        arr(i) = v
        v = v + step_
    Next i
    IntegerRange1D = arr
End Function

'*******************************************************************************
'Checks if a specified row, in a 2D array, has no values
'Parameters:
'   - arr: a 2D array
'   - rowIndex: the index of the row to check if is empty
'   - [ignoreEmptyStrings]:
'       * True - Empty String values are considered Empty
'       * False - Empty String values are not considered Empty. Default
'Raises error:
'   - 5 if:
'       * the input array is not 2-dimensional
'       * the row index is out of bounds
'*******************************************************************************
Public Function Is2DArrayRowEmpty(ByRef arr As Variant _
                                , ByVal rowIndex As Long _
                                , Optional ByVal ignoreEmptyStrings As Boolean = False) As Boolean
    Const fullMethodName As String = MODULE_NAME & ".Is2DArrayRowEmpty"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 2 Then
        Err.Raise 5, fullMethodName, "Array is not two-dimensional"
    ElseIf rowIndex < LBound(arr, 1) Or rowIndex > UBound(arr, 1) Then
        Err.Raise 5, fullMethodName, "Row Index out of bounds"
    End If
    '
    Dim j As Long
    Dim v As Variant
    '
    'Exit Function if any non-Empty value is found (returns False)
    For j = LBound(arr, 2) To UBound(arr, 2)
        If IsObject(arr(rowIndex, j)) Then Exit Function
        v = arr(rowIndex, j)
        Select Case VarType(v)
        Case VbVarType.vbEmpty
            'Continue to next element
        Case VbVarType.vbString
            If Not ignoreEmptyStrings Then Exit Function
            If LenB(v) > 0 Then Exit Function
        Case Else
            Exit Function
        End Select
    Next j
    '
    Is2DArrayRowEmpty = True 'If code reached this line then row is Empty
End Function

'*******************************************************************************
'Checks if a Variant is iterable using a For Each... Next loop
'Compatible types: Arrays, Collections, Custom Collections, Dictionaries etc.
'Does not raise errors
'*******************************************************************************
Public Function IsIterable(ByRef list_ As Variant) As Boolean
    Dim v As Variant
    '
    'Custom collections classes that use Attribute NewEnum.VB_UserMemID = -4 (to
    '   get a default enumerator to be used with For Each... constructions) are
    '   causing automation errors and crashes on x64
    'Avoid bug by using a 'Set' statement or by using a call to another method
    Set v = Nothing
    '
    On Error Resume Next
    For Each v In list_
        Exit For
    Next v
    IsIterable = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Returns a boolean indicating if a value passes a filter
'Parameters:
'   - value_: the target value to check against the filter
'   - filter: a FILTER_PAIR struct (operator and compare value)
'Raises error:
'   - 5 if:
'       * filter's operator is invalid
'       * if target value or compare value are User Defined Types
'       * values are incompatible with the operator
'       * values are of incompatible data type
'       * a compare value is an array but it is not iterable
'Notes:
'   - the filter's compare value can be a list (array, collection)
'   - IN and NOT IN operators will also work if the right member is just a
'     primitive value by switching to EQUAL and respectively NOT EQUAL
'   - LIKE and NOT LIKE operators only work if the compare value is a text
'     (pattern) and the target value is one of: bool, number, text
'*******************************************************************************
Public Function IsValuePassingFilter(ByRef value_ As Variant _
                                   , ByRef filter As FILTER_PAIR) As Boolean
    Const fullMethodName As String = MODULE_NAME & ".IsValuePassingFilter"
    Dim rnk As DATA_TYPE_RANK: rnk = GetDataTypeRank(value_)
    Dim isListOperator As Boolean
    '
    isListOperator = (filter.cOperator = opin) Or (filter.cOperator = opNotIn)
    '
    'Validate input
    If filter.cOperator < [_opMin] Or filter.cOperator > [_opMax] Then
        Err.Raise 5, fullMethodName, "Invalid Filter Operator"
    ElseIf rnk = rankUDT Or filter.compValue.rank_ = rankUDT Then
        Err.Raise 5, fullMethodName, "User Defined Types not supported"
    ElseIf rnk = rankArray Then
        IsValuePassingFilter = (filter.cOperator = opNotIn) _
                            Or (filter.cOperator = opNotEqual)
        Exit Function
    ElseIf filter.compValue.rank_ = rankArray Then
        If filter.compValue.isIterable_ Then
            If Not isListOperator Then
                Err.Raise 5, fullMethodName, "Incompatible filter"
            End If
        Else
            Err.Raise 5, fullMethodName, "Compare Array is not iterable"
        End If
    End If
    '
    'Treat Empty as Null String by adjusting ranks
    If rnk = rankEmpty Then rnk = rankText
    If filter.compValue.rank_ = rankEmpty Then filter.compValue.rank_ = rankText
    '
    'Check for list inclusion
    If isListOperator And filter.compValue.isIterable_ Then
        Dim isIncluded As Boolean
        Dim textKey As String: textKey = GetUniqueTextKey(value_)
        '
        isIncluded = CollectionHasKey(filter.compValue.textKeys_, textKey)
        IsValuePassingFilter = isIncluded Xor (filter.cOperator = opNotIn)
        Exit Function
    Else
        'Adjust inclusion operators because compare value is not iterable
        If filter.cOperator = opin Then filter.cOperator = opEqual
        If filter.cOperator = opNotIn Then filter.cOperator = opNotEqual
    End If
    '
    'Check target value against filter
    Select Case filter.cOperator
    Case opSmaller, opBigger, opSmallerOrEqual, opBiggerOrEqual
        If rnk <> filter.compValue.rank_ Then Exit Function
        If rnk < rankBoolean Then Exit Function
        '
        Select Case filter.cOperator
        Case opSmaller
            IsValuePassingFilter = (value_ < filter.compValue.value_)
        Case opBigger
            IsValuePassingFilter = (value_ > filter.compValue.value_)
        Case opSmallerOrEqual
            IsValuePassingFilter = (value_ <= filter.compValue.value_)
        Case opBiggerOrEqual
            IsValuePassingFilter = (value_ >= filter.compValue.value_)
        End Select
        '
        'Force False < True. In VBA: False > True (0 > -1)
        If rnk = rankBoolean Then
            Dim areDifferent As Boolean
            areDifferent = (value_ <> filter.compValue.value_)
            IsValuePassingFilter = IsValuePassingFilter Xor areDifferent
        End If
    Case opLike, opNotLike
        If filter.compValue.rank_ <> rankText Then Exit Function 'Text Pattern
        If rnk < rankBoolean Then Exit Function
        '
        'The 'Like' operator uses Option Compare Text (see top of the module)
        Dim isLike As Boolean: isLike = (value_ Like filter.compValue.value_)
        IsValuePassingFilter = isLike Xor (filter.cOperator = opNotLike)
    Case opEqual, opNotEqual
        If rnk <> filter.compValue.rank_ Then
            IsValuePassingFilter = (filter.cOperator = opNotEqual)
            Exit Function
        End If
        '
        Dim areEqual As Boolean
        '
        If rnk = rankNull Then
            areEqual = True
        ElseIf rnk = rankObject Then
            Dim key1 As String: key1 = GetUniqueTextKey(value_)
            Dim key2 As String: key2 = GetUniqueTextKey(filter.compValue.value_)
            areEqual = (key1 = key2)
        Else
            areEqual = (value_ = filter.compValue.value_)
        End If
        '
        IsValuePassingFilter = areEqual Xor (filter.cOperator = opNotEqual)
    Case Else
        Err.Raise 5, fullMethodName, "Invalid Operator"
    End Select
End Function

'*******************************************************************************
'Merges/Combines two 1D arrays into a new 1D array
'Parameters:
'   - arr1: the first 1D array
'   - arr2: the second 1D array
'   - [outLowBound]: the start index of the result array. Default is 0
'Raises error:
'   - 5 if any of the two arrays is not 1D
'Note:
'   - if you wish to merge two 1D arrays vertically, convert them first to 2D
'     and then use the 'Merge2DArrays' method
'Examples:
'   - arr1 = [1,2] and arr2 = [3,4,5] >> results [1,2,3,4,5]
'*******************************************************************************
Public Function Merge1DArrays(ByRef arr1 As Variant _
                            , ByRef arr2 As Variant _
                            , Optional ByVal outLowBound As Long = 0) As Variant
    Const fullMethodName As String = MODULE_NAME & ".Merge1DArrays"
    '
    'Check Dimensions
    If GetArrayDimsCount(arr1) <> 1 Or GetArrayDimsCount(arr2) <> 1 Then
        Err.Raise 5, fullMethodName, "Expected 1D Arrays"
    End If
    '
    Dim elemCount1 As Long: elemCount1 = UBound(arr1, 1) - LBound(arr1, 1) + 1
    Dim elemCount2 As Long: elemCount2 = UBound(arr2, 1) - LBound(arr2, 1) + 1
    '
    'Check for zero-length arrays
    If elemCount1 = 0 Then
        If LBound(arr2) = outLowBound Then
            Merge1DArrays = arr2
            Exit Function
        End If
    End If
    If elemCount2 = 0 Then
        If LBound(arr1) = outLowBound Then
            Merge1DArrays = arr1
            Exit Function
        End If
    End If
    '
    Dim totalCount As Long: totalCount = elemCount1 + elemCount2
    If totalCount = 0 Then
        Merge1DArrays = ZeroLengthArray()
        Exit Function
    End If
    '
    Dim res() As Variant: ReDim res(outLowBound To outLowBound + totalCount - 1)
    Dim i As Long
    Dim v As Variant
    '
    'Copy first array
    i = outLowBound
    For Each v In arr1 'Column-major order
        If IsObject(v) Then Set res(i) = v Else res(i) = v
        i = i + 1
    Next v
    '
    'Copy second array
    For Each v In arr2
        If IsObject(v) Then Set res(i) = v Else res(i) = v
        i = i + 1
    Next v
    '
    Merge1DArrays = res
End Function

'*******************************************************************************
'Merges/Combines two 2D arrays into a new 2D array
'Parameters:
'   - arr1: the first 2D array
'   - arr2: the second 2D array
'   - verticalMerge:
'       * True - arrays are combined vertically i.e. rows are combined
'       * False - arrays are combined horizontally i.e. columns are combined
'   - [outLowRow]: the start index of the result's 1st dimension. Default is 0
'   - [outLowCol]: the start index of the result's 2nd dimension. Default is 0
'Raises error:
'   - 5 if:
'       * any of the two arrays is not 2D
'       * the number of rows or columns are incompatible for merging
'Examples:
'   - arr1 = [1,2] and arr2 = [5,6] and verticalMerge = True >> results [1,2]
'            [3,4]            [7,8]                                     [3,4]
'                                                                       [5,6]
'                                                                       [7,8]
'   - arr1 = [1,2] and arr2 = [5,6] and verticalMerge = False > results [1,2,5,6]
'            [3,4]            [7,8]                                     [3,4,7,8]
'*******************************************************************************
Public Function Merge2DArrays(ByRef arr1 As Variant _
                            , ByRef arr2 As Variant _
                            , ByVal verticalMerge As Boolean _
                            , Optional ByVal outLowRow As Long = 0 _
                            , Optional ByVal outLowCol As Long = 0) As Variant()
    Const fullMethodName As String = MODULE_NAME & ".Merge2DArrays"
    '
    'Check Dimensions
    If GetArrayDimsCount(arr1) <> 2 Or GetArrayDimsCount(arr2) <> 2 Then
        Err.Raise 5, fullMethodName, "Expected 2D Arrays"
    End If
    '
    Dim rowsCount1 As Long: rowsCount1 = UBound(arr1, 1) - LBound(arr1, 1) + 1
    Dim rowsCount2 As Long: rowsCount2 = UBound(arr2, 1) - LBound(arr2, 1) + 1
    Dim colsCount1 As Long: colsCount1 = UBound(arr1, 2) - LBound(arr1, 2) + 1
    Dim colsCount2 As Long: colsCount2 = UBound(arr2, 2) - LBound(arr2, 2) + 1
    Dim totalRows As Long
    Dim totalCols As Long
    '
    'Check if rows/columns are compatible
    If verticalMerge Then
        If colsCount1 <> colsCount2 Then
            Err.Raise 5, fullMethodName, "Expected same number of columns"
        End If
        totalRows = rowsCount1 + rowsCount2
        totalCols = colsCount1
    Else 'Horizontal merge
        If rowsCount1 <> rowsCount2 Then
            Err.Raise 5, fullMethodName, "Expected same number of rows"
        End If
        totalRows = rowsCount1
        totalCols = colsCount1 + colsCount2
    End If
    '
    Dim res() As Variant: ReDim res(outLowRow To outLowRow + totalRows - 1 _
                                  , outLowCol To outLowCol + totalCols - 1)
    Dim i As Long
    Dim j As Long
    Dim v As Variant
    Dim r1 As Long: r1 = outLowRow + rowsCount1
    '
    'Copy first array
    i = outLowRow
    j = outLowCol
    'For Each... loop is faster than using 2 For... Next loops
    For Each v In arr1 'Column-major order
        If IsObject(v) Then Set res(i, j) = v Else res(i, j) = v
        i = i + 1
        If i = r1 Then 'Switch to the next column
            j = j + 1
            i = outLowRow
        End If
    Next v
    '
    Dim r2 As Long: r2 = outLowRow + totalRows
    Dim s2 As Long: s2 = r2 - rowsCount2
    '
    'Copy second array
    i = s2
    j = outLowCol + totalCols - colsCount2
    For Each v In arr2
        If IsObject(v) Then Set res(i, j) = v Else res(i, j) = v
        i = i + 1
        If i = r2 Then 'Switch to the next column
            j = j + 1
            i = s2
        End If
    Next v
    '
    Merge2DArrays = res
End Function

'*******************************************************************************
'Converts a multidimensional array to a 1 dimension array
'Parameters:
'   - arr: the array to convert
'   - traverseType: read arr elements in row-wise or column-wise order
'Notes:
'   * Invalid traverseType values are defaulted to column-wise order
'Examples:
'   - arr = [1,2] and traverseType = columnWise >> results [1,3,2,4]
'           [3,4]
'   - arr = [1,2] and traverseType = rowWise >> results [1,2,3,4]
'           [3,4]
'*******************************************************************************
Public Function NDArrayTo1DArray(ByRef arr As Variant _
                               , ByVal traverseType As ARRAY_TRAVERSE_TYPE) As Variant
    Const fullMethodName As String = MODULE_NAME & ".NDArrayTo1DArray"
    '
    'Check Array Dimensions
    Select Case GetArrayDimsCount(arr)
    Case 0
        Err.Raise 5, fullMethodName, "Invalid or Uninitialized Array"
    Case 1
        NDArrayTo1DArray = arr
    Case Else
        Dim trType As ARRAY_TRAVERSE_TYPE
        '
        If traverseType = rowWise Then trType = rowWise Else trType = columnWise
        NDArrayTo1DArray = GetArrayDescriptor(arr, trType, True).elements1D
    End Select
End Function

'*******************************************************************************
'Builds and returns an ARRAY_DESCRIPTOR structure for the specified array
'Parameters:
'   - arr: the array used to generate the custom structure
'   - traverseType: read array elements in row-wise or column-wise order
'   - addElements1D: populates the ".elements1D" array for the ARRAY_DESCRIPTOR
'*******************************************************************************
Private Function GetArrayDescriptor(ByRef arr As Variant _
                                  , ByVal traverseType As ARRAY_TRAVERSE_TYPE _
                                  , ByVal addElements1D As Boolean) As ARRAY_DESCRIPTOR
    Dim descStruct As ARRAY_DESCRIPTOR
    Dim i As Long
    '
    'Prepare Struct for looping
    With descStruct
        .dimsCount = GetArrayDimsCount(arr)
        ReDim .dimensions(1 To .dimsCount)
        .elemCount = 1 'Start value. Can turn into 0 if array is zero-length
        .traverseType = traverseType
    End With
    '
    'Loop through the array dimensions and store the size of each dimension, the
    '   depth of each dimension (the product of lower dimension sizes) and the
    '   total count of elements in the entire array
    For i = descStruct.dimsCount To 1 Step -1
        With descStruct.dimensions(i)
            .index = i
            .depth = descStruct.elemCount
            .size = UBound(arr, i) - LBound(arr, i) + 1
            descStruct.elemCount = descStruct.elemCount * .size
        End With
    Next i
    '
    'Populate elements as a 1-dimensional array (vector)
    If addElements1D Then AddElementsToDescriptor descStruct, arr
    '
    GetArrayDescriptor = descStruct
End Function

'*******************************************************************************
'Populates the ".elements1D" array for an ARRAY_DESCRIPTOR Structure
'Parameters:
'   - descStruct: the structure to populate
'   - sourceArray: the multidimensional array containing the elements
'*******************************************************************************
Private Sub AddElementsToDescriptor(ByRef descStruct As ARRAY_DESCRIPTOR _
                                  , ByRef sourceArray As Variant)
    'Note that zero-length arrays (1 dimension) are covered as well
    If descStruct.dimsCount = 1 Then
        descStruct.elements1D = sourceArray
        Exit Sub
    End If
    '
    'Size the elements vector
    ReDim descStruct.elements1D(0 To descStruct.elemCount - 1)
    '
    Dim tempElement As Variant
    Dim i As Long: i = 0
    '
    'Populate target, element by element
    If descStruct.traverseType = rowWise Then
        Dim rowMajorIndex As Long
        '
        With descStruct
            ReDim .rowMajorIndexes(1 To .elemCount)
            AddRowMajorIndexes descStruct, .dimensions(.dimsCount), 1, 1
        End With
        For Each tempElement In sourceArray
            i = i + 1
            rowMajorIndex = descStruct.rowMajorIndexes(i) - 1
            If IsObject(tempElement) Then
                Set descStruct.elements1D(rowMajorIndex) = tempElement
            Else
                descStruct.elements1D(rowMajorIndex) = tempElement
            End If
        Next tempElement
    Else 'columnWise - VBA already stores arrays in columnMajorOrder
        For Each tempElement In sourceArray
            If IsObject(tempElement) Then
                Set descStruct.elements1D(i) = tempElement
            Else
                descStruct.elements1D(i) = tempElement
            End If
            i = i + 1
        Next tempElement
    End If
End Sub

'*******************************************************************************
'Populates the "rowMajorIndexes" array for an ARRAY_DESCRIPTOR Structure
'Recursive!
'Parameters:
'   - descStruct: the details structure. See ARRAY_DESCRIPTOR custom type
'   - currDimension: the current dimension that the function is working on.
'                    In the initial call it must be the last dimension (lowest)
'   - colWiseIndex: first element (current dimension) column-major index
'                   In the initial call must have the value of 1
'   - rowWiseIndex: element row-major index implemented as a counter (ByRef)
'                   In the initial call must have the value of 1
'*******************************************************************************
Private Sub AddRowMajorIndexes(ByRef descStruct As ARRAY_DESCRIPTOR _
                             , ByRef currDimension As ARRAY_DIMENSION _
                             , ByVal colWiseIndex As Long _
                             , ByRef rowWiseIndex As Long)
    Dim i As Long
    Dim tempIndex As Long
    '
    If currDimension.index = 1 Then 'First dimension (highest). Populate Indexes
        For i = 0 To currDimension.size - 1
            tempIndex = colWiseIndex + i * currDimension.depth
            descStruct.rowMajorIndexes(rowWiseIndex) = tempIndex
            rowWiseIndex = rowWiseIndex + 1
        Next i
    Else 'Pass colWise and rowWise indexes to higher dimensions
        Dim prevDim As ARRAY_DIMENSION
        '
        prevDim = descStruct.dimensions(currDimension.index - 1)
        For i = 0 To currDimension.size - 1
            tempIndex = colWiseIndex + i * currDimension.depth
            AddRowMajorIndexes descStruct, prevDim, tempIndex, rowWiseIndex
        Next i
    End If
End Sub

'*******************************************************************************
'Multidimensional array to collections
'Mimics the way arrays are stored in other languages (vectors inside vectors)
'   except that the indexes will start with 1 instead of 0 (VBA Collections)
'*******************************************************************************
Public Function NDArrayToCollections(ByRef arr As Variant) As Collection
    Const fullMethodName As String = MODULE_NAME & ".NDArrayToCollections"
    '
    If GetArrayDimsCount(arr) = 0 Then
        Err.Raise 5, fullMethodName, "Invalid or Uninitialized Array"
    End If
    '
    Dim descStruct As ARRAY_DESCRIPTOR
    descStruct = GetArrayDescriptor(arr, rowWise, True)
    '
    Set NDArrayToCollections = GetCollsFromDescriptor( _
        descStruct, descStruct.dimensions(1), LBound(descStruct.elements1D))
End Function

'*******************************************************************************
'Returns nested Collections using the descriptor of a multidimensional Array
'Recursive!
'Parameters:
'   - descStruct: the descriptor structure
'   - currDimension: the current dimension that the function is working on.
'                    In the initial call must be the first dimension (highest)
'   - elemIndex: element index implemented as a counter (ByRef).
'                In the initial call must be LBound of 'descStruct.elements1D'
'*******************************************************************************
Private Function GetCollsFromDescriptor(ByRef descStruct As ARRAY_DESCRIPTOR _
                                      , ByRef currDimension As ARRAY_DIMENSION _
                                      , ByRef elemIndex As Long) As Collection
    Dim collResult As New Collection
    Dim i As Long
    '
    If currDimension.index = descStruct.dimsCount Then
        'Last dimension (lowest). Populate Elements
        For i = 1 To currDimension.size
            collResult.Add descStruct.elements1D(elemIndex)
            elemIndex = elemIndex + 1
        Next i
    Else 'Get Collections for lower dimensions
        Dim nextDim As ARRAY_DIMENSION
        '
        nextDim = descStruct.dimensions(currDimension.index + 1)
        For i = 1 To currDimension.size
            collResult.Add GetCollsFromDescriptor(descStruct, nextDim, elemIndex)
        Next i
    End If
    '
    Set GetCollsFromDescriptor = collResult
End Function

'*******************************************************************************
'Returns a 2D array based on values contained in the specified 1D array
'Parameters:
'   - arr: the 1D array that contains the values to be used
'   - columnsCount: the number of columns that the result 2D array will have
'Raises error:
'   -  5 if:
'       * input array is not 1D
'       * input array has no elements i.e. zero-length array
'       * the number of columns is less than 1
'Notes:
'   - if the total Number of values is not divisible by columnsCount then the
'     extra values (last row) of the array are by default the value Empty
'*******************************************************************************
Public Function OneDArrayTo2DArray(ByRef arr As Variant _
                                 , ByVal columnsCount As Long) As Variant()
    Const fullMethodName As String = MODULE_NAME & ".OneDArrayTo2DArray"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 1 Then
        Err.Raise 5, fullMethodName, "Expected 1D Array"
    ElseIf LBound(arr) > UBound(arr) Then
        Err.Raise 5, fullMethodName, "Zero-length array. No elements"
    ElseIf columnsCount < 1 Then
        Err.Raise 5, fullMethodName, "Invalid Columns Count"
    End If
    '
    Dim elemCount As Long: elemCount = UBound(arr) - LBound(arr) + 1
    Dim rowsCount As Long: rowsCount = -Int(-elemCount / columnsCount)
    Dim res() As Variant
    Dim i As Long: i = 0
    Dim r As Long
    Dim c As Long
    Dim v As Variant
    '
    'Populate result array
    ReDim res(0 To rowsCount - 1, 0 To columnsCount - 1)
    For Each v In arr
        r = i \ columnsCount
        c = i Mod columnsCount
        If IsObject(v) Then Set res(r, c) = v Else res(r, c) = v
        i = i + 1
    Next v
    '
    OneDArrayTo2DArray = res
End Function

'*******************************************************************************
'Returns a collection based on values contained in the specified 1D array
'Parameters:
'   - arr: the 1D array that contains the values to be used
'Raises error:
'   -  5 if: input array is not 1D
'*******************************************************************************
Public Function OneDArrayToCollection(ByRef arr As Variant) As Collection
    Const fullMethodName As String = MODULE_NAME & ".OneDArrayToCollection"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 1 Then
        Err.Raise 5, fullMethodName, "Expected 1D Array"
    End If
    '
    Dim coll As New Collection
    Dim v As Variant
    '
    For Each v In arr
        coll.Add v
    Next v
    '
    Set OneDArrayToCollection = coll
End Function

'*******************************************************************************
'Replaces Empty values within an Array
'Curently supports 1D, 2D and 3D arrays
'*******************************************************************************
Public Sub ReplaceEmptyInArray(ByRef arr As Variant, ByVal newVal As Variant)
    Dim needsSet As Boolean: needsSet = IsObject(newVal)
    Dim v As Variant
    Dim i As Long
    '
    Select Case GetArrayDimsCount(arr)
    Case 1
        i = LBound(arr, 1)
        For Each v In arr
            If IsEmpty(v) Then
                If needsSet Then Set arr(i) = newVal Else arr(i) = newVal
            End If
            i = i + 1
        Next v
    Case 2
        Dim lowerRow As Long: lowerRow = LBound(arr, 1)
        Dim upperRow As Long: upperRow = UBound(arr, 1)
        Dim j As Long
        '
        i = lowerRow
        j = LBound(arr, 2)
        'For Each... Next loop is faster than using 2 For... Next loops
        For Each v In arr 'Column-major order
            If IsEmpty(v) Then
                If needsSet Then Set arr(i, j) = newVal Else arr(i, j) = newVal
            End If
            If i = upperRow Then 'Switch to the next column
                j = j + 1
                i = lowerRow
            Else
                i = i + 1
            End If
        Next v
    Case 3
        Dim k As Long
        '
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                For k = LBound(arr, 3) To UBound(arr, 3)
                    If IsEmpty(arr(i, j, k)) Then
                        If needsSet Then
                            Set arr(i, j, k) = newVal
                        Else
                            arr(i, j, k) = newVal
                        End If
                    End If
                Next k
            Next j
        Next i
    Case Else
        'Add logic as needed (e.g. for 4 dimensions)
    End Select
End Sub

'*******************************************************************************
'Reverses (in groups) a 1D Array, in-place
'Returns:
'   - the reversed 1D array
'Parameters:
'   - arr: a 1D array of values to be reversed
'   - [groupSize]: the number of values in each group. Default is 1
'Notes:
'   - the array is reversed in place so it is optional to use the return value
'     of the function
'Raises error:
'   - 5 if:
'       * array is not one-dimensional
'       * array has no elements (zero-length array)
'       * groupSize is smaller than 1
'       * the number of elements is not divisible by the groupSize
'Examples:
'   - arr = [1,2,3,4,5,6] and groupSize = 1 >> returns [6,5,4,3,2,1]
'   - arr = [1,2,3,4,5,6] and groupSize = 2 >> returns [5,6,3,4,1,2]
'   - arr = [1,2,3,4,5,6] and groupSize = 3 >> returns [4,5,6,1,2,3]
'   - arr = [1,2,3,4,5,6] and groupSize = 4 >> error 5 is raised
'   - arr = [1,2,3,4,5,6] and groupSize = 5 >> error 5 is raised
'   - arr = [1,2,3,4,5,6] and groupSize = 6 >> returns [1,2,3,4,5,6]
'*******************************************************************************
Public Function Reverse1DArray(ByRef arr As Variant _
                             , Optional ByVal groupSize As Long = 1) As Variant
    Const fullMethodName As String = MODULE_NAME & ".Reverse1DArray"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 1 Then
        Err.Raise 5, fullMethodName, "Expected 1D Array"
    ElseIf LBound(arr) > UBound(arr) Then
        Err.Raise 5, fullMethodName, "Zero-length array. No elements"
    ElseIf groupSize < 1 Then
        Err.Raise 5, fullMethodName, "Invalid GroupSize"
    ElseIf (UBound(arr, 1) - LBound(arr, 1) + 1) Mod groupSize <> 0 Then
        Err.Raise 5, fullMethodName, "Elements not divisible by groupSize"
    End If
    '
    Dim index1 As Long: index1 = LBound(arr, 1)
    Dim index2 As Long: index2 = UBound(arr, 1) - groupSize + 1
    Dim i As Long
    '
    'Reverse
    Do While index1 < index2
        For i = 1 To groupSize
            Swap1DArrayValues arr, index1, index2
            index1 = index1 + 1
            index2 = index2 + 1
        Next i
        index2 = index2 - 2 * groupSize
    Loop
    '
    Reverse1DArray = arr 'Useful for method chaining
End Function

'*******************************************************************************
'Reverses (in groups) a 2D Array, in-place
'Returns:
'   - the reversed 2D array
'Parameters:
'   - arr: a 2D Array of values to be reversed
'   - [groupSize]: the number of values in each group. Default is 1
'   - [verticalFlip]:
'       * True - reverse vertically
'       * False - reverse horizontally (default)
'Notes:
'   - the array is reversed in place so it is optional to use the return value
'     of the function
'Raises error:
'   - 5 if:
'       * array is not two-dimensional
'       * groupSize is smaller than 1
'       * the number of elements is not divisible by the groupSize
'Examples:
'   - arr = [1,2,3,4], groupSize = 2, verticalFlip = False > returns [3,4,1,2]
'           [5,6,7,8]                                                [7,8,5,6]
'   - arr = [1,2,3,4], groupSize = 1, verticalFlip = True >> returns [5,6,7,8]
'           [5,6,7,8]                                                [1,2,3,4]
'*******************************************************************************
Public Function Reverse2DArray(ByRef arr As Variant _
                             , Optional ByVal groupSize As Long = 1 _
                             , Optional ByVal verticalFlip As Boolean = False) As Variant
    Const fullMethodName As String = MODULE_NAME & ".Reverse2DArray"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 2 Then
        Err.Raise 5, fullMethodName, "Expected 2D Array"
    ElseIf groupSize < 1 Then
        Err.Raise 5, fullMethodName, "Invalid GroupSize"
    ElseIf verticalFlip Then
        If (UBound(arr, 1) - LBound(arr, 1) + 1) Mod groupSize <> 0 Then
            Err.Raise 5, fullMethodName, "Rows not divisible by groupSize"
        End If
    ElseIf (UBound(arr, 2) - LBound(arr, 2) + 1) Mod groupSize <> 0 Then
        Err.Raise 5, fullMethodName, "Columns not divisible by groupSize"
    End If
    '
    Dim dimension As Long: If verticalFlip Then dimension = 1 Else dimension = 2
    Dim index1 As Long: index1 = LBound(arr, dimension)
    Dim index2 As Long: index2 = UBound(arr, dimension) - groupSize + 1
    Dim i As Long
    '
    'Reverse rows or columns
    Do While index1 < index2
        For i = 1 To groupSize
            If verticalFlip Then
                Swap2DArrayRows arr, index1, index2
            Else
                Swap2DArrayColumns arr, index1, index2
            End If
            index1 = index1 + 1
            index2 = index2 + 1
        Next i
        index2 = index2 - 2 * groupSize
    Loop
    '
    Reverse2DArray = arr 'Useful for method chaining
End Function

'*******************************************************************************
'Reverses a Collection, in groups, in-place
'Returns:
'   - the reversed Collection
'Parameters:
'   - coll: a collection of values to reverse
'   - [groupSize]: the number of values in each group. Default is 1
'Notes:
'   - the collection is reversed in place so it is optional to use the return
'     value of the function
'   - a collection that has no elements is returned as-is
'Raises Error:
'   - 91: if Collection Object is not set
'   -  5: if groupSize is smaller than 1
'   -  5: if the number of elements is not divisible by the groupSize
'Examples:
'   - coll = [1,2,3,4,5,6] and groupSize = 1 >> returns [6,5,4,3,2,1]
'   - coll = [1,2,3,4,5,6] and groupSize = 2 >> returns [5,6,3,4,1,2]
'   - coll = [1,2,3,4,5,6] and groupSize = 3 >> returns [4,5,6,1,2,3]
'   - coll = [1,2,3,4,5,6] and groupSize = 4 >> error 5 is raised
'   - coll = [1,2,3,4,5,6] and groupSize = 5 >> error 5 is raised
'   - coll = [1,2,3,4,5,6] and groupSize = 6 >> returns [1,2,3,4,5,6]
'*******************************************************************************
Public Function ReverseCollection(ByVal coll As Collection _
                                , Optional ByVal groupSize As Long = 1) As Collection
    Const fullMethodName As String = MODULE_NAME & ".ReverseCollection"
    '
    'Check Input
    If coll Is Nothing Then
        Err.Raise 91, fullMethodName, "Collection not set"
    ElseIf coll.Count = 0 Then
        Err.Raise 5, fullMethodName, "No elements"
    ElseIf groupSize < 1 Then
        Err.Raise 5, fullMethodName, "Invalid GroupSize"
    ElseIf coll.Count Mod groupSize <> 0 Then
        Err.Raise 5, fullMethodName, "Elements not divisible by groupSize"
    End If
    '
    Dim index1 As Long: index1 = 1
    Dim index2 As Long: index2 = coll.Count - groupSize + 1
    Dim i As Long
    '
    'Reverse
    Do While index1 < index2
        For i = 1 To groupSize
            coll.Add Item:=coll.Item(index2), Before:=index1
            index1 = index1 + 1
            index2 = index2 + 1
            coll.Remove index2
        Next i
        index2 = index2 - groupSize
    Loop
    '
    Set ReverseCollection = coll 'Useful for method chaining
End Function

'*******************************************************************************
'Creates and returns an arithmetic progression sequence as a 1D array
'Parameters:
'   - termsCount: the number of terms
'   - [initialTerm]: the value of the first term
'   - [commonDifference]: the difference between any 2 consecutive terms
'Raises error:
'   - 5: if 'termsCount' is not at least 1
'Theory:
'   https://en.wikipedia.org/wiki/Arithmetic_progression
'*******************************************************************************
Public Function Sequence1D(ByVal termsCount As Long _
                         , Optional ByVal initialTerm As Double = 1 _
                         , Optional ByVal commonDifference As Double = 1) As Double()
    Const fullMethodName As String = MODULE_NAME & ".Sequence1D"
    If termsCount < 1 Then Err.Raise 5, fullMethodName, "Wrong number of terms"
    '
    Dim arr() As Double: ReDim arr(0 To termsCount - 1)
    Dim i As Long
    '
    For i = 0 To termsCount - 1
        arr(i) = initialTerm + i * commonDifference
    Next i
    Sequence1D = arr
End Function

'*******************************************************************************
'Creates and returns an arithmetic progression sequence as a 2D array
'Parameters:
'   - termsCount: the number of terms
'   - [initialTerm]: the value of the first term
'   - [commonDifference]: the difference between any 2 consecutive terms
'   - [columnsCount]: the number of columns in the output 2D Array
'Raises error:
'   - 5 if:
'       * 'termsCount' is not at least 1
'       * 'columnsCount' is not at least 1
'Theory:
'   https://en.wikipedia.org/wiki/Arithmetic_progression
'Notes:
'   - if the total number of terms is not divisible by columnsCount then the
'     extra values (last row) of the array are by default 0 (zero)
'*******************************************************************************
Public Function Sequence2D(ByVal termsCount As Long _
                         , Optional ByVal initialTerm As Double = 1 _
                         , Optional ByVal commonDifference As Double = 1 _
                         , Optional ByVal columnsCount As Long = 1) As Double()
    Const fullMethodName As String = MODULE_NAME & ".Sequence2D"
    If termsCount < 1 Then
        Err.Raise 5, fullMethodName, "Wrong number of terms"
    ElseIf columnsCount < 1 Then
        Err.Raise 5, fullMethodName, "Expected at least 1 output column"
    End If
    '
    Dim rowsCount As Long: rowsCount = -Int(-termsCount / columnsCount)
    Dim arr() As Double: ReDim arr(0 To rowsCount - 1, 0 To columnsCount - 1)
    Dim r As Long
    Dim c As Long
    Dim i As Long
    '
    For i = 0 To termsCount - 1
        r = i \ columnsCount
        c = i Mod columnsCount
        arr(r, c) = initialTerm + i * commonDifference
    Next i
    Sequence2D = arr
End Function

'*******************************************************************************
'Returns a shallow copy of a Source Collection
'A new collection is created and then populated with values and references to
'   to the objects found in the original collection
'Returns Nothing if the source is Nothing
'*******************************************************************************
Public Function ShallowCopyCollection(ByVal sourceColl As Collection) As Collection
    If sourceColl Is Nothing Then Exit Function
    '
    Dim v As Variant
    Dim targetColl As New Collection
    '
    For Each v In sourceColl
        targetColl.Add v
    Next v
    Set ShallowCopyCollection = targetColl
End Function

'*******************************************************************************
'Returns a slice of a 1D array as a new 1D array
'Parameters:
'   - arr: a 1D array to slice
'   - startIndex: the index of the first element to be added to result
'   - length_: the number of elements to return
'   - [outLowBound]: the start index of the result array. Default is 0
'Notes:
'   - excess length is ignored
'Raises error:
'   - 5 if:
'       * array is not 1-dimensional
'       * startIndex or length are invalid
'Examples (assumed lower bound of array is 0):
'   - arr = [1,2,3,4], startIndex = 0 and length_ = 2 >> results [1,2]
'   - arr = [1,2,3,4], startIndex = 2 and length_ = 5 >> results [3,4]
'*******************************************************************************
Public Function Slice1DArray(ByRef arr As Variant _
                           , ByVal startIndex As Long _
                           , ByVal length_ As Long _
                           , Optional ByVal outLowBound As Long = 0) As Variant
    Const fullMethodName As String = MODULE_NAME & ".Slice1DArray"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 1 Then
        Err.Raise 5, fullMethodName, "Expected 1D Array"
    ElseIf startIndex < LBound(arr, 1) Or startIndex > UBound(arr, 1) Then
        Err.Raise 5, fullMethodName, "Invalid startIndex"
    ElseIf length_ <= 0 Then
        Err.Raise 5, fullMethodName, "Invalid slice length"
    ElseIf startIndex = LBound(arr) _
       And startIndex + length_ > UBound(arr) _
       And startIndex = outLowBound _
    Then
        Slice1DArray = arr
        Exit Function
    End If
    '
    Dim endIndex As Long: endIndex = startIndex + length_ - 1
    '
    'Ignore excess length
    If endIndex > UBound(arr, 1) Then endIndex = UBound(arr, 1)
    '
    Dim res() As Variant
    Dim i As Long
    Dim adjust As Long: adjust = outLowBound - startIndex
    '
    'Add elements to result array
    ReDim res(outLowBound To endIndex + adjust)
    For i = startIndex To endIndex
        If IsObject(arr(i)) Then
            Set res(i + adjust) = arr(i)
        Else
            res(i + adjust) = arr(i)
        End If
    Next i
    '
    Slice1DArray = res
End Function

'*******************************************************************************
'Returns a slice of a 2D array as a new 2D array
'Parameters:
'   - arr: a 2D array to slice
'   - startRow: the index of the first row to be added to result
'   - startColumn: the index of the first column to be added to result
'   - height_: the number of rows to be returned
'   - width_: the number of columns to be returned
'   - [outLowRow]: the start index of the result's 1st dimension. Default is 0
'   - [outLowCol]: the start index of the result's 2nd dimension. Default is 0
'Notes:
'   - excess height or width is ignored
'Raises error:
'   - 5 if:
'       * array is not two-dimensional
'       * startRow, startColumn, height or width are invalid
'Examples assuming arr = [1,2,3,4] with both lower bounds as 0:
'                        [5,6,7,8]
'   - startRow = 0, startColumn = 1, height_ = 2, width_ = 2 >> results [2,3]
'                                                                       [6,7]
'   - startRow = 1, startColumn = 1, height_ = 2, width_ = 6 >> results [6,7,8]
'*******************************************************************************
Public Function Slice2DArray(ByRef arr As Variant _
                           , ByVal startRow As Long _
                           , ByVal startColumn As Long _
                           , ByVal height_ As Long _
                           , ByVal width_ As Long _
                           , Optional ByVal outLowRow As Long = 0 _
                           , Optional ByVal outLowCol As Long = 0) As Variant
    Const fullMethodName As String = MODULE_NAME & ".Slice2DArray"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 2 Then
        Err.Raise 5, fullMethodName, "Array is not two-dimensional"
    ElseIf startRow < LBound(arr, 1) Or startRow > UBound(arr, 1) Then
        Err.Raise 5, fullMethodName, "Invalid startRow"
    ElseIf startColumn < LBound(arr, 2) Or startColumn > UBound(arr, 2) Then
        Err.Raise 5, fullMethodName, "Invalid startColumn"
    ElseIf height_ <= 0 Then
        Err.Raise 5, fullMethodName, "Invalid height"
    ElseIf width_ <= 0 Then
        Err.Raise 5, fullMethodName, "Invalid width"
    ElseIf startRow = LBound(arr, 1) And startColumn = LBound(arr, 2) Then
        If startRow + height_ > UBound(arr, 1) _
        And startColumn + width_ > UBound(arr, 2) Then
            If startRow = outLowRow And startColumn = outLowCol Then
                Slice2DArray = arr
                Exit Function
            End If
        End If
    End If
    '
    Dim endRow As Long: endRow = startRow + height_ - 1
    Dim endColumn As Long: endColumn = startColumn + width_ - 1
    '
    'Ignore excess lengths
    If endRow > UBound(arr, 1) Then endRow = UBound(arr, 1)
    If endColumn > UBound(arr, 2) Then endColumn = UBound(arr, 2)
    '
    Dim res() As Variant
    Dim i As Long
    Dim j As Long
    Dim adjustRow As Long: adjustRow = outLowRow - startRow
    Dim adjustCol As Long: adjustCol = outLowCol - startColumn
    '
    'Add elements to result array
    ReDim res(outLowRow To endRow + adjustRow _
            , outLowCol To endColumn + adjustCol)
    For i = startRow To endRow
        For j = startColumn To endColumn
            If IsObject(arr(i, j)) Then
                Set res(i + adjustRow, j + adjustCol) = arr(i, j)
            Else
                res(i + adjustRow, j + adjustCol) = arr(i, j)
            End If
        Next j
    Next i
    '
    Slice2DArray = res
End Function

'*******************************************************************************
'Returns a slice of a collection as a new collection
'Parameters:
'   - coll: a collection to slice
'   - startIndex: the index of the first element to be added to result
'   - length_: the number of elements to return
'Notes:
'   - excess length is ignored
'Raises error:
'   - 91 if:
'       * Collection is not set
'       * startIndex or length are invalid
'Examples:
'   - arr = [1,2,3,4], startIndex = 1 and length_ = 2 >> results [1,2]
'   - arr = [1,2,3,4], startIndex = 2 and length_ = 5 >> results [2,3,4]
'*******************************************************************************
Public Function SliceCollection(ByVal coll As Collection _
                              , ByVal startIndex As Long _
                              , ByVal length_ As Long) As Collection
    Const fullMethodName As String = MODULE_NAME & ".SliceCollection"
    '
    'Check Input
    If coll Is Nothing Then
        Err.Raise 91, fullMethodName, "Collection not set"
    ElseIf startIndex < 1 Or startIndex > coll.Count Then
        Err.Raise 5, fullMethodName, "Invalid startIndex"
    ElseIf length_ <= 0 Then
        Err.Raise 5, fullMethodName, "Invalid slice length"
    End If
    '
    Dim endIndex As Long: endIndex = startIndex + length_ - 1
    '
    'Ignore excess length
    If endIndex > coll.Count Then endIndex = coll.Count
    '
    Dim collRes As New Collection
    Dim i As Long
    '
    'Add elements to result collection
    For i = startIndex To endIndex
        collRes.Add coll.Item(i)
    Next i
    '
    Set SliceCollection = collRes
End Function

'*******************************************************************************
'Sort a 1D Array, in-place
'Returns:
'   - the sorted 1D array
'Parameters:
'   - arr: a 1D array of values to sort
'   - [sortAscending]:
'       * True - Ascending (default)
'       * False - Descending
'   - [useTextNumberAsNumber]:
'       * True - numbers stored as texts are considered numbers (default)
'       * False - numbers stored as texts are considered texts
'   - [caseSensitive]:
'       * True - compare texts as case-sensitive
'       * False - ignore case when comparing texts (default)
'Notes:
'   - this function is using a Stable Quick Sort adaptation. See 'QuickSort'
'     method below
'   - the array is sorted in place so it is optional to use the return value of
'     the function
'Raises Error:
'   - 5: if array is not one-dimensional
'*******************************************************************************
Public Function Sort1DArray(ByRef arr As Variant _
                          , Optional ByVal sortAscending As Boolean = True _
                          , Optional ByVal useTextNumberAsNumber As Boolean = True _
                          , Optional ByVal caseSensitive As Boolean = False) As Variant
    Const fullMethodName As String = MODULE_NAME & ".Sort1DArray"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 1 Then
        Err.Raise 5, fullMethodName, "Expected 1D Array"
    ElseIf UBound(arr, 1) - LBound(arr, 1) <= 0 Then
        Sort1DArray = arr '1 or no value / nothing to sort
        Exit Function
    End If
    '
    Dim cOptions As COMPARE_OPTIONS
    '
    'Set Compare Options
    cOptions.compAscending = sortAscending
    cOptions.useTextNumberAsNumber = useTextNumberAsNumber
    cOptions.compareMethod = IIf(caseSensitive, vbBinaryCompare, vbTextCompare)
    '
    Dim lowerIndex As Long: lowerIndex = LBound(arr, 1)
    Dim upperIndex As Long: upperIndex = UBound(arr, 1)
    Dim arrIndex() As Long: ReDim arrIndex(lowerIndex To upperIndex)
    Dim i As Long
    '
    'QuickSort is not a 'Stable' Sort. An array of indexes is needed to ensure
    '   that equal values preserve their order. This is accomplished by sorting
    '   the array of indexes at the same time with the actual array being sorted
    '   and by comparing indexes when values are equal (or of equal rank)
    For i = lowerIndex To upperIndex
        arrIndex(i) = i
    Next i
    '
    QuickSortVector arr, lowerIndex, upperIndex, cOptions, arrIndex, vecArray
    Sort1DArray = arr 'Useful for method chaining
End Function

'*******************************************************************************
'Sorts a Vector, in-place. Could be a 1D Array or a Collection
'Notes:
'   - This method is recursive so the initial call must include the lower and
'     upper bounds of the 1D array
'   - The CompareValues function is used to compare elements
'   - To make this Sort a 'Stable' Sort, an array of indexes must be built and
'     passed in the initial call (from outside). This is done in order to ensure
'     that equal values preserve their order (by comparing initial indexes when
'     values are equal)
'Theory:
'   - https://en.wikipedia.org/wiki/Quicksort
'*******************************************************************************
Private Sub QuickSortVector(ByRef vector As Variant _
                          , ByVal lowIndex As Long _
                          , ByVal uppIndex As Long _
                          , ByRef cOptions As COMPARE_OPTIONS _
                          , ByRef arrIndex() As Long _
                          , ByVal vecType As VECTOR_TYPE)
    If lowIndex >= uppIndex Then Exit Sub
    '
    Dim p As Long:          p = (lowIndex + uppIndex) \ 2
    Dim piv As SORT_PIVOT:  SetSortPivot piv, arrIndex(p), vector(p)
    Dim newLoIndex As Long: newLoIndex = lowIndex
    Dim newUpIndex As Long: newUpIndex = uppIndex
    Dim cr As COMPARE_RESULT
    '
    Do While newLoIndex <= newUpIndex
        'Increase 'newLoIndex' until a swap is needed
        Do While newLoIndex < uppIndex
            cr = CompareValues(vector(newLoIndex), piv.value_, cOptions)
            If cr.mustSwap Then Exit Do
            If cr.areEqual Then If arrIndex(newLoIndex) >= piv.index Then Exit Do
            newLoIndex = newLoIndex + 1
        Loop
        'Decrease 'newUpIndex' until a swap is needed
        Do While newUpIndex > lowIndex
            cr = CompareValues(piv.value_, vector(newUpIndex), cOptions)
            If cr.mustSwap Then Exit Do
            If cr.areEqual Then If piv.index >= arrIndex(newUpIndex) Then Exit Do
            newUpIndex = newUpIndex - 1
        Loop
        'Swap values, if needed
        If newLoIndex <= newUpIndex Then
            Select Case vecType
            Case vecArray
                Swap1DArrayValues vector, newLoIndex, newUpIndex
            Case vecCollection
                SwapCollectionValues vector, newLoIndex, newUpIndex
            End Select
            Swap1DArrayValues arrIndex, newLoIndex, newUpIndex 'Sync Indexes
            newLoIndex = newLoIndex + 1
            newUpIndex = newUpIndex - 1
        End If
    Loop
    'Sort both remaining sub-vectors
    QuickSortVector vector, lowIndex, newUpIndex, cOptions, arrIndex, vecType
    QuickSortVector vector, newLoIndex, uppIndex, cOptions, arrIndex, vecType
End Sub

'*******************************************************************************
'Set a SORT_PIVOT struct from values
'*******************************************************************************
Private Sub SetSortPivot(ByRef sPivot As SORT_PIVOT _
                       , ByVal index As Long _
                       , ByVal v As Variant)
    sPivot.index = index
    If IsObject(v) Then Set sPivot.value_ = v Else sPivot.value_ = v
End Sub

'*******************************************************************************
'Compare 2 values of unknown type (Variant) based on a ranking convention and
'   sorting compare options
'Notes:
'   - GetDataTypeRank function (returns DATA_TYPE_RANK enumeration) is used to
'     rank the compared values based on a predefined enumeration. If the values
'     being compared are of different ranks then the comparison is made based
'     on rank alone
'   - Empty values are always moved at the end (ignoring sort order), mimicking
'     how Excel sorts ranges of values
'   - Note that we consider False < True (in VBA False > True), mimicking how
'     Excel sorts ranges of values
'Utility for the QuickSort methods
'*******************************************************************************
Private Function CompareValues(ByRef val1 As Variant _
                             , ByRef val2 As Variant _
                             , ByRef cOptions As COMPARE_OPTIONS) As COMPARE_RESULT
    Dim rnk1 As DATA_TYPE_RANK: rnk1 = GetDataTypeRank(val1)
    Dim rnk2 As DATA_TYPE_RANK: rnk2 = GetDataTypeRank(val2)
    '
    'Adjust Rank for numbers stored as text, if needed
    If cOptions.useTextNumberAsNumber Then
        If rnk1 >= rankText And rnk2 >= rankText Then
            If rnk1 = rankText Then If IsNumeric(val1) Then rnk1 = rankNumber
            If rnk2 = rankText Then If IsNumeric(val2) Then rnk2 = rankNumber
        End If
    End If
    '
    'Compare ranks/values as appropriate
    If rnk1 < rnk2 Then
        CompareValues.mustSwap = (rnk1 = rankEmpty Or cOptions.compAscending)
    ElseIf rnk1 > rnk2 Then
        CompareValues.mustSwap = Not (rnk2 = rankEmpty Or cOptions.compAscending)
    Else 'Ranks are equal
        Select Case rnk1
        Case rankEmpty, rankUDT, rankObject, rankArray, rankNull, rankError
            CompareValues.areEqual = True 'For sorting purposes
        Case rankBoolean
            If val1 = val2 Then
                CompareValues.areEqual = True
            Else
                'Note that we consider False < True (in VBA False > True)
                CompareValues.mustSwap = (cOptions.compAscending Xor val2)
            End If
        Case rankText
            Select Case StrComp(val1, val2, cOptions.compareMethod)
                Case -1: CompareValues.mustSwap = Not cOptions.compAscending
                Case 0:  CompareValues.areEqual = True
                Case 1:  CompareValues.mustSwap = cOptions.compAscending
            End Select
        Case rankNumber
            Dim no1 As Double: no1 = CDbl(val1) 'Maybe test for decimal on Win
            Dim no2 As Double: no2 = CDbl(val2)
            '
            If no1 = no2 Then
                CompareValues.areEqual = True
            Else
                CompareValues.mustSwap = (cOptions.compAscending Xor no1 < no2)
            End If
        End Select
    End If
End Function

'*******************************************************************************
'Returns a rank (Enum) for a given value's data type which simplifies the number
'   of existing data types. This simplification speeds up the comparison of 2
'   values by removing the need to compare values of incompatible data types
'*******************************************************************************
Private Function GetDataTypeRank(ByRef varValue As Variant) As DATA_TYPE_RANK
    If IsObject(varValue) Then
        GetDataTypeRank = rankObject
        Exit Function
    End If
    Select Case VarType(varValue)
    Case vbNull
        GetDataTypeRank = rankNull
    Case vbEmpty
        GetDataTypeRank = rankEmpty
    Case vbError
        GetDataTypeRank = rankError
    Case vbBoolean
        GetDataTypeRank = rankBoolean
    Case vbString
        GetDataTypeRank = rankText
    Case vbByte, vbInteger, vbLong, vbLongLong 'Integers
        GetDataTypeRank = rankNumber
    Case vbCurrency, vbDecimal, vbDouble, vbSingle, vbDate 'Decimal-point
        GetDataTypeRank = rankNumber
    Case vbArray To vbArray + vbUserDefinedType
        GetDataTypeRank = rankArray
    Case vbUserDefinedType
        GetDataTypeRank = rankUDT
    Case vbDataObject
        GetDataTypeRank = rankObject
    End Select
End Function

'*******************************************************************************
'Sort a 2D Array by a particular column, in-place
'Returns:
'   - the sorted 2D array using the specified column for comparison
'Parameters:
'   - arr: a 2D array of values to sort
'   - sortColumn: the index of the column used for sorting
'   - [sortAscending]:
'       * True - Ascending (default)
'       * False - Descending
'   - [useTextNumberAsNumber]:
'       * True - numbers stored as texts are considered numbers (default)
'       * False - numbers stored as texts are considered texts
'   - [caseSensitive]:
'       * True - compare texts as case-sensitive
'       * False - ignore case when comparing texts (default)
'Notes:
'   - this function is using a Stable Quick Sort adaptation. See 'QuickSort'
'     method below
'   - the array is sorted in place so it is optional to use the return value of
'     the function
'Raises Error:
'   - 5 if:
'       * array is not two-dimensional
'       * sort column index is out of bounds
'*******************************************************************************
Public Function Sort2DArray(ByRef arr As Variant, ByVal sortColumn As Long _
                          , Optional ByVal sortAscending As Boolean = True _
                          , Optional ByVal useTextNumberAsNumber As Boolean = True _
                          , Optional ByVal caseSensitive As Boolean = False) As Variant
    Const fullMethodName As String = MODULE_NAME & ".Sort2DArray"
    '
    'Check Input
    If GetArrayDimsCount(arr) <> 2 Then
        Err.Raise 5, fullMethodName, "Array is not two-dimensional"
    ElseIf sortColumn < LBound(arr, 2) Or sortColumn > UBound(arr, 2) Then
        Err.Raise 5, fullMethodName, "Sort Column out of bounds"
    ElseIf UBound(arr, 1) - LBound(arr, 1) = 0 Then
        Sort2DArray = arr 'Only 1 row / nothing to sort
        Exit Function
    End If
    '
    Dim cOptions As COMPARE_OPTIONS
    '
    'Set Compare Options
    cOptions.compAscending = sortAscending
    cOptions.useTextNumberAsNumber = useTextNumberAsNumber
    cOptions.compareMethod = IIf(caseSensitive, vbBinaryCompare, vbTextCompare)
    '
    Dim lowerRow As Long: lowerRow = LBound(arr, 1)
    Dim upperRow As Long: upperRow = UBound(arr, 1)
    Dim arrIndex() As Long: ReDim arrIndex(lowerRow To upperRow)
    Dim i As Long
    '
    'QuickSort is not a 'Stable' Sort. An array of indexes is needed to ensure
    '   that equal values preserve their order. This is accomplished by sorting
    '   the array of indexes at the same time with the actual array being sorted
    '   and by comparing indexes when values are equal (or of equal rank)
    For i = lowerRow To upperRow
        arrIndex(i) = i
    Next i
    '
    QuickSort2DArray arr, lowerRow, upperRow, sortColumn, cOptions, arrIndex
    Sort2DArray = arr 'Useful for method chaining
End Function

'*******************************************************************************
'Sort a 2D Array (in place) by using a Stable Quick Sort adaptation. Stable
'   means that the order for equal elements is preserved
'Notes:
'   - This method is recursive so the initial call must include the lower and
'     upper bounds of the first dimension
'   - The CompareValues function is used to compare elements on the sort column
'   - To make this Sort a 'Stable' Sort, an array of indexes must be built and
'     passed in the initial call (from outside). This is done in order to ensure
'     that equal values preserve their order (by comparing initial indexes when
'     values are equal)
'Theory:
'   - https://en.wikipedia.org/wiki/Quicksort
'*******************************************************************************
Private Sub QuickSort2DArray(ByRef arr As Variant _
                           , ByVal lowerRow As Long _
                           , ByVal upperRow As Long _
                           , ByVal sortColumn As Long _
                           , ByRef cOptions As COMPARE_OPTIONS _
                           , ByRef arrIndex() As Long)
    If lowerRow >= upperRow Then Exit Sub
    '
    Dim p As Long:         p = (lowerRow + upperRow) \ 2
    Dim piv As SORT_PIVOT: SetSortPivot piv, arrIndex(p), arr(p, sortColumn)
    Dim newLowRow As Long: newLowRow = lowerRow
    Dim newUppRow As Long: newUppRow = upperRow
    Dim cr As COMPARE_RESULT
    '
    Do While newLowRow <= newUppRow
        'Increase 'newLowRow' until a swap is needed
        Do While newLowRow < upperRow
            cr = CompareValues(arr(newLowRow, sortColumn), piv.value_, cOptions)
            If cr.mustSwap Then Exit Do
            If cr.areEqual Then If arrIndex(newLowRow) >= piv.index Then Exit Do
            newLowRow = newLowRow + 1
        Loop
        'Decrease 'newUppRow' until a swap is needed
        Do While newUppRow > lowerRow
            cr = CompareValues(piv.value_, arr(newUppRow, sortColumn), cOptions)
            If cr.mustSwap Then Exit Do
            If cr.areEqual Then If piv.index >= arrIndex(newUppRow) Then Exit Do
            newUppRow = newUppRow - 1
        Loop
        'Swap rows, if needed
        If newLowRow <= newUppRow Then
            Swap2DArrayRows arr, newLowRow, newUppRow
            Swap1DArrayValues arrIndex, newLowRow, newUppRow 'Sync Indexes
            newLowRow = newLowRow + 1
            newUppRow = newUppRow - 1
        End If
    Loop
    'Sort both remaining sub-arrays
    QuickSort2DArray arr, lowerRow, newUppRow, sortColumn, cOptions, arrIndex
    QuickSort2DArray arr, newLowRow, upperRow, sortColumn, cOptions, arrIndex
End Sub

'*******************************************************************************
'Sorts a Collection, in-place
'Returns:
'   - the sorted collection
'Parameters:
'   - coll: a Collection to sort
'   - [sortAscending]:
'       * True - Ascending (default)
'       * False - Descending
'   - [useTextNumberAsNumber]:
'       * True - numbers stored as texts are considered numbers (default)
'       * False - numbers stored as texts are considered texts
'   - [caseSensitive]:
'       * True - compare texts as case-sensitive
'       * False - ignore case when comparing texts (default)
'Notes:
'   - this function is using a Stable Quick Sort adaptation
'   - the collection is sorted in place so it is optional to use the return
'     value of the function
'Raises Error:
'   - 91: if collection is not set
'*******************************************************************************
Public Function SortCollection(ByRef coll As Collection _
                             , Optional ByVal sortAscending As Boolean = True _
                             , Optional ByVal useTextNumberAsNumber As Boolean = True _
                             , Optional ByVal caseSensitive As Boolean = False) As Collection
    Const fullMethodName As String = MODULE_NAME & ".SortCollection"
    '
    'Check Input
    If coll Is Nothing Then
        Err.Raise 91, fullMethodName, "Collection not set"
    ElseIf coll.Count <= 1 Then
        Set SortCollection = coll '1 or no value / nothing to sort
        Exit Function
    End If
    '
    Dim cOptions As COMPARE_OPTIONS
    '
    'Set Compare Options
    cOptions.compAscending = sortAscending
    cOptions.useTextNumberAsNumber = useTextNumberAsNumber
    cOptions.compareMethod = IIf(caseSensitive, vbBinaryCompare, vbTextCompare)
    '
    Dim arrIndex() As Long: ReDim arrIndex(1 To coll.Count)
    Dim i As Long
    '
    'QuickSort is not a 'Stable' Sort. An array of indexes is needed to ensure
    '   that equal values preserve their order. This is accomplished by sorting
    '   the array of indexes at the same time with the actual array being sorted
    '   and by comparing indexes when values are equal (or of equal rank)
    For i = 1 To coll.Count
        arrIndex(i) = i
    Next i
    '
    QuickSortVector coll, 1, coll.Count, cOptions, arrIndex, vecCollection
    Set SortCollection = coll  'Useful for method chaining
End Function

'*******************************************************************************
'Swaps 2 values in a 1D Array, in-place
'*******************************************************************************
Private Sub Swap1DArrayValues(ByRef arr As Variant _
                            , ByVal index1 As Long _
                            , ByVal index2 As Long)
    If index1 <> index2 Then SwapValues arr(index1), arr(index2)
End Sub

'*******************************************************************************
'Swaps 2 values in a Collection, in-place
'*******************************************************************************
Private Sub SwapCollectionValues(ByVal coll As Collection _
                               , ByVal index1 As Long _
                               , ByVal index2 As Long)
    If index1 = index2 Then Exit Sub
    '
    Dim i1 As Long
    Dim i2 As Long
    '
    If index1 < index2 Then
        i1 = index1
        i2 = index2
    Else
        i1 = index2
        i2 = index1
    End If
    '
    coll.Add Item:=coll.Item(i1), Before:=i2
    coll.Add Item:=coll.Item(i2 + 1), Before:=i1
    coll.Remove i1 + 1
    coll.Remove i2 + 1
End Sub

'*******************************************************************************
'Swaps 2 columns in a 2D Array, in-place
'*******************************************************************************
Private Sub Swap2DArrayColumns(ByRef arr As Variant _
                             , ByVal column1 As Long _
                             , ByVal column2 As Long)
    If column1 <> column2 Then
        Dim i As Long
        '
        For i = LBound(arr, 1) To UBound(arr, 1)
            SwapValues arr(i, column1), arr(i, column2)
        Next i
    End If
End Sub

'*******************************************************************************
'Swaps 2 rows in a 2D Array, in-place
'*******************************************************************************
Private Sub Swap2DArrayRows(ByRef arr As Variant _
                          , ByVal row1 As Long _
                          , ByVal row2 As Long)
    If row1 <> row2 Then
        Dim j As Long
        '
        For j = LBound(arr, 2) To UBound(arr, 2)
            SwapValues arr(row1, j), arr(row2, j)
        Next j
    End If
End Sub

'*******************************************************************************
'Swaps 2 values of any data type
'*******************************************************************************
Public Sub SwapValues(ByRef val1 As Variant, ByRef val2 As Variant)
    Dim temp As Variant
    Dim needsSet1 As Boolean: needsSet1 = IsObject(val1)
    Dim needsSet2 As Boolean: needsSet2 = IsObject(val2)
    '
    If needsSet1 Then Set temp = val1 Else temp = val1
    If needsSet2 Then Set val1 = val2 Else val1 = val2
    If needsSet1 Then Set val2 = temp Else val2 = temp
End Sub

'*******************************************************************************
'Returns a Collection where keys are the received list of texts and items
'   are their corresponding position/index
'Parameters:
'   - arrText: a 1D array or a 2D array with 1 column or 1 row of values that
'              are or can be casted to String
'   - [ignoreDuplicates]:
'       * True - any duplicated text is ignored (first found position returned)
'       * False - error 457 will get raised if duplicate is found
'Raises error:
'   - 457: if duplicate is found and 'ignoreDuplicates' is set to 'False'
'   -  13: if any of the values received cannot be casted to String
'   -   5: if input is not a 1D array or a single row/column 2D array
'Example usage:
'    Dim arrHeaders() As Variant: arrHeaders = headersRange.Value2
'    Dim headerIndex As Collection: Set headerIndex = TextArrayToIndex(arrHeaders)
'    Dim v As Variant
'    Dim h As String
'    For Each v In requiredHeadersList
'        If Not CollectionHasKey(headerIndex, v) Then
'            MsgBox "Missing header: " & v
'            Exit Sub
'        End If
'    Next v
'    h = "An existing header"
'    Debug.Print "Position of " & h & " is " & headerIndex(h)
'*******************************************************************************
Public Function TextArrayToIndex(ByRef arrText() As Variant _
                               , Optional ByVal ignoreDuplicates As Boolean = True) As Collection
    Const fullMethodName As String = MODULE_NAME & ".TextArrayToIndex"
    '
    Dim i As Long
    Dim v As Variant
    Dim collIndex As New Collection
    Dim dimsCount As Long: dimsCount = GetArrayDimsCount(arrText)
    Const errDuplicate As Long = 457
    '
    If dimsCount = 2 Then
        Dim r As Long: r = UBound(arrText, 1) - LBound(arrText, 1) + 1
        Dim c As Long: c = UBound(arrText, 2) - LBound(arrText, 2) + 1
        '
        If r > 1 And c > 1 Then
            Err.Raise 5, fullMethodName, "Expected row or column of texts"
        ElseIf r = 1 Then
            i = LBound(arrText, 2)
        Else
            i = LBound(arrText, 1)
        End If
    ElseIf dimsCount = 1 Then
        i = LBound(arrText, 1)
    Else
        Err.Raise 5, fullMethodName, "Expected 1D or 2D array of text values"
    End If
    '
    On Error Resume Next
    For Each v In arrText
        collIndex.Add i, CStr(v)
        If Err.Number <> 0 Then
            If Err.Number = errDuplicate Then
                If Not ignoreDuplicates Then
                    On Error GoTo 0
                    Err.Raise errDuplicate, fullMethodName, "Duplicated text"
                End If
            Else
                On Error GoTo 0
                Err.Raise 13, fullMethodName, "Type mismatch. Expected text"
            End If
        End If
        i = i + 1
    Next v
    On Error GoTo 0
    '
    Set TextArrayToIndex = collIndex
End Function

'*******************************************************************************
'Transposes a 1D or 2D Array
'Raises error:
'   - 5: if input array is not 1D or 2D
'Notes:
'   - 1D Arrays are transposed to a 1 column 2D Array
'   - resulting bounds are reflecting the input bounds
'*******************************************************************************
Public Function TransposeArray(ByRef arr As Variant) As Variant()
    Const fullMethodName As String = MODULE_NAME & ".TransposeArray"
    Dim res() As Variant
    '
    Select Case GetArrayDimsCount(arr)
    Case 1
        If LBound(arr, 1) > UBound(arr, 1) Then
            TransposeArray = ZeroLengthArray()
            Exit Function
        End If
        Dim lowBound As Long: lowBound = LBound(arr, 1)
        ReDim res(lowBound To UBound(arr, 1), lowBound To lowBound)
    Case 2
        ReDim res(LBound(arr, 2) To UBound(arr, 2) _
                , LBound(arr, 1) To UBound(arr, 1))
    Case Else
        Err.Raise 5, fullMethodName, "Expected 1D or 2D Array"
    End Select
    '
    Dim v As Variant
    Dim lowerCol As Long: lowerCol = LBound(res, 2)
    Dim upperCol As Long: upperCol = UBound(res, 2)
    Dim i As Long: i = LBound(res, 1)
    Dim j As Long: j = lowerCol
    '
    'For Each... loop is faster than using 2 For... Next loops
    For Each v In arr 'Column-major order
        If IsObject(v) Then Set res(i, j) = v Else res(i, j) = v
        If j = upperCol Then 'Switch to next row
            i = i + 1
            j = lowerCol
        Else
            j = j + 1
        End If
    Next v
    TransposeArray = res
End Function

'*******************************************************************************
'Receives a value or multiple values via a variant and returns either a single
'   collection containing all the values (from all nest levels) or nested
'   collections
'Parameters:
'   - values: the Value(s) that will be returned in a new collection
'   - nestType (applicable to array/collection/range but can be extended):
'       * nestNone - return a single collection of values that are not nested.
'         No returned element can be an array, Excel Range or a collection
'       * nestMultiItemsOnly - maintain original nesting but only if the list
'         (array/collection/range) has more than 1 element and return
'         collection(s) inside collection(s). Arrays and Ranges are turned into
'         Collections as well
'       * nestAll - maintain original nesting and return collection(s) inside
'         collection(s). Arrays and Ranges are turned into Collections as well
'   - traverseArrType (applicable to multi-dimensional arrays only):
'       * rowWise
'       * columnWise
'Does not raise errors
'Notes:
'   - invalid array traverseType values are defaulted to column-wise order
'   - invalid nestType values are defaulted to nestNone (no nesting)
'   - uninitialized arrays are ignored
'*******************************************************************************
Public Function ValuesToCollection(ByRef values As Variant _
                                 , ByVal nestType As NESTING_TYPE _
                                 , ByVal traverseArrType As ARRAY_TRAVERSE_TYPE) As Collection
    If traverseArrType <> rowWise Then traverseArrType = columnWise
    If nestType < [_nMin] And nestType > [_nMax] Then nestType = nestNone
    '
    Dim coll As New Collection
    '
    AddToCollection values, coll, nestType, traverseArrType, False, True
    Set ValuesToCollection = coll
End Function

'*******************************************************************************
'Adds all values to the specified target collection, recursively
'Called from ValuesToCollection
'*******************************************************************************
Private Sub AddToCollection(ByRef values As Variant _
                          , ByVal coll As Collection _
                          , ByVal nestType As NESTING_TYPE _
                          , ByVal traverseType As ARRAY_TRAVERSE_TYPE _
                          , ByVal hasSiblings As Boolean _
                          , Optional ByVal isRoot As Boolean = False)
    Dim v As Variant
    Dim hasMultiItems As Boolean
    '
    If IsObject(values) Then
        If values Is Nothing Then
            coll.Add Nothing
        ElseIf TypeOf values Is Collection Then
            hasMultiItems = (values.Count > 1)
            If NeedsNesting(nestType, hasMultiItems, hasSiblings, isRoot) Then
                Set coll = AddNewCollectionTo(coll)
            End If
            For Each v In values
                AddToCollection v, coll, nestType, traverseType, hasMultiItems
            Next v
        ElseIf IsExcelRange(values) Then
            hasMultiItems = (values.Count > 1)
            If NeedsNesting(nestType, hasMultiItems, hasSiblings, isRoot) Then
                Set coll = AddNewCollectionTo(coll)
            End If
            For Each v In values.Areas
                AddToCollection v.Value2, coll, nestNone, traverseType, hasMultiItems
            Next v
        Else
            'Logic can be added here, for any other object type(s) needed
            coll.Add values
        End If
    ElseIf IsArray(values) Then
        Dim dimsCount As Long: dimsCount = GetArrayDimsCount(values)
        '
        If dimsCount > 0 Then
            hasMultiItems = (GetArrayElemCount(values) > 1)
            If NeedsNesting(nestType, hasMultiItems, hasSiblings, isRoot) Then
                Set coll = AddNewCollectionTo(coll)
            End If
            If traverseType = rowWise And dimsCount > 1 Then
                values = NDArrayTo1DArray(values, rowWise)
            End If
            For Each v In values
                AddToCollection v, coll, nestType, traverseType, hasMultiItems
            Next v
        End If
    Else
        coll.Add values
    End If
End Sub

'*******************************************************************************
'Utility for 'AddToCollection'
'*******************************************************************************
Private Function NeedsNesting(ByVal nestType As NESTING_TYPE _
                            , ByVal hasMultiItems As Boolean _
                            , ByVal hasSiblings As Boolean _
                            , ByVal isRoot As Boolean) As Boolean
    Select Case nestType
        Case nestAll:            NeedsNesting = Not isRoot
        Case nestMultiItemsOnly: NeedsNesting = (hasMultiItems And hasSiblings)
        Case Else:               NeedsNesting = False
    End Select
End Function

'*******************************************************************************
'Utility for 'AddToCollection'
'*******************************************************************************
Private Function AddNewCollectionTo(ByVal collTarget As Collection) As Collection
    Set AddNewCollectionTo = New Collection
    collTarget.Add AddNewCollectionTo
End Function

'*******************************************************************************
'Checks if a Variant is of Excel.Range type
'It compiles for other Applications in addition to Excel (like Word, PowerPoint)
'*******************************************************************************
Private Function IsExcelRange(ByRef v As Variant) As Boolean
    If TypeName(v) = "Range" Then
        On Error Resume Next
        IsExcelRange = (v.Areas.Count > 0)
        On Error GoTo 0
    End If
End Function

'*******************************************************************************
'Returns a Zero-Length array of Variant type
'*******************************************************************************
Public Function ZeroLengthArray() As Variant()
    #If Mac Then
        ZeroLengthArray = Array()
    #Else
        #If Win64 Then
            ZeroLengthArray = Array() 'Could be done using APIs as below
        #Else
            'There's a bug in x32 when using Array(). It cannot be assigned to
            '   another Variant, cannot be added to Collections/Arrays
            'Solution is to build an array using Windows APIs
            '
            'Update Jan-2021
            'The bug seems to have been fixed in newer version of Excel so, a
            '   static array will be used to mimimize Win API calls
            Static zArr As Variant
            Static isArrSet As Boolean
            '
            If Not isArrSet Then
                zArr = Array()
                '
                'Try assigning to another variant
                Dim v As Variant
                On Error Resume Next
                v = zArr
                isArrSet = (Err.Number = 0)
                On Error GoTo 0
            End If
            If Not isArrSet Then
                Const vType As Integer = vbVariant
                Dim bounds(0 To 0) As SAFEARRAYBOUND
                Dim ptrArray As Long 'No need for LongPtr (x32 branch)
                Dim tVariant As TagVariant
                '
                'Create empty array and store pointer
                ptrArray = SafeArrayCreate(vType, 1, bounds(0))
                '
                'Create a Variant pointing to the array
                tVariant.vt = vbArray + vType
                tVariant.ptr = ptrArray
                '
                'Copy result
                VariantCopy zArr, tVariant
                '
                'Clean-up
                SafeArrayDestroy ptrArray
                isArrSet = True
            End If
            ZeroLengthArray = zArr
        #End If
    #End If
End Function
