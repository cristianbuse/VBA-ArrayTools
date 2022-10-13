Attribute VB_Name = "TestLibArrayTools"
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

'*******************************************************************************
''This module tests methods from the LibArrayTools module
''Call:
''  - TestLibArrayTools.RunAllTests
'*******************************************************************************

Private Const MODULE_NAME As String = "TestLibArrayTools"

'Struct returned by test methods
Private Type TEST_RESULT
    passed As Boolean
    methodName As String
    failDetails As String
End Type

Private Type EXPECTED_ERROR
    code_ As Long
    wasRaised As Boolean
End Type

'The error code raised by the Assert... methods
Private Const ERR_ASSERT_FAILED As Long = &HBAD

'*******************************************************************************
'Creates an EXPECTED_ERROR struct
'*******************************************************************************
Private Function NewExpectedError(ByVal code_ As Long) As EXPECTED_ERROR
    NewExpectedError.code_ = code_
    NewExpectedError.wasRaised = False
End Function

'*******************************************************************************
'Main procedure
'*******************************************************************************
Public Sub RunAllTests()
    Dim testResults() As TEST_RESULT
    Dim t As Double: t = Timer
    '
    AddTestResult testResults, TestGetArrayDimsCount
    AddTestResult testResults, TestGetArrayElemCount
    AddTestResult testResults, TestNDArrayToCollections
    AddTestResult testResults, TestCollectionToCSV 'Internal Utility for Testing
    AddTestResult testResults, TestArrayToCSV      'Internal Utility for Testing
    AddTestResult testResults, TestCollection
    AddTestResult testResults, TestCollectionHasKey
    AddTestResult testResults, TestCollectionTo1DArray
    AddTestResult testResults, TestCollectionTo2DArray
    AddTestResult testResults, TestNDArrayTo1DArray
    AddTestResult testResults, TestValuesToCollection
    AddTestResult testResults, TestIsIterable
    AddTestResult testResults, TestCreateFilter
    AddTestResult testResults, TestIsValuePassingFilter
    AddTestResult testResults, TestCreateFiltersArray
    AddTestResult testResults, TestFilter1DArray
    AddTestResult testResults, TestOneDArrayTo2DArray
    AddTestResult testResults, TestFilter2DArray
    AddTestResult testResults, TestFilterCollection
    AddTestResult testResults, TestFindTextsRow
    AddTestResult testResults, TestIntegerRange1D
    AddTestResult testResults, TestGetUniqueIntegers
    AddTestResult testResults, TestGetUniqueRows
    AddTestResult testResults, TestGetUniqueValues
    AddTestResult testResults, TestInsertRowsAtIndex
    AddTestResult testResults, TestInsertRowsAtValChange
    AddTestResult testResults, TestIs2DArrayRowEmpty
    AddTestResult testResults, TestMerge1DArrays
    AddTestResult testResults, TestMerge2DArrays
    AddTestResult testResults, TestOneDArrayToCollection
    AddTestResult testResults, TestRemoveEmptyRows
    AddTestResult testResults, TestReplaceEmptyInArray
    AddTestResult testResults, TestReplaceNullInArray
    AddTestResult testResults, TestReverse1DArray
    AddTestResult testResults, TestReverse2DArray
    AddTestResult testResults, TestReverseCollection
    AddTestResult testResults, TestSequence1D
    AddTestResult testResults, TestSequence2D
    AddTestResult testResults, TestShallowCopyCollection
    AddTestResult testResults, TestSlice1DArray
    AddTestResult testResults, TestSlice2DArray
    AddTestResult testResults, TestSliceCollection
    AddTestResult testResults, TestSort1DArray
    AddTestResult testResults, TestSort2DArray
    AddTestResult testResults, TestSortCollection
    AddTestResult testResults, TestSwapValues
    AddTestResult testResults, TestTextArrayToIndex
    AddTestResult testResults, TestTransposeArray
    '
    ShowTestResults testResults, Timer - t
End Sub

'*******************************************************************************
'Adds a single TEST_RESULT struct to the end of a target array
'*******************************************************************************
Private Sub AddTestResult(ByRef arrTarget() As TEST_RESULT, ByRef testResult As TEST_RESULT)
    Dim lowerBound As Long
    Dim upperBound As Long
    '
    On Error Resume Next
    lowerBound = LBound(arrTarget)
    upperBound = UBound(arrTarget)
    If Err.Number <> 0 Then upperBound = upperBound - 1
    On Error GoTo 0
    '
    ReDim Preserve arrTarget(lowerBound To upperBound + 1)
    arrTarget(upperBound + 1) = testResult
End Sub

'*******************************************************************************
'Displays the results for an array of TEST_RESULT structs
'*******************************************************************************
Private Sub ShowTestResults(arr() As TEST_RESULT, ByVal secondsDuration As Double)
    Dim testResult As TEST_RESULT
    Dim i As Long
    Dim failedCount As Long
    Dim totalCount As Long: totalCount = UBound(arr) - LBound(arr) + 1
    Dim arrOut() As String
    '
    ReDim arrOut(LBound(arr) To UBound(arr), 0 To 2)
    For i = LBound(arr) To UBound(arr)
        testResult = arr(i)
        If Not testResult.passed Then failedCount = failedCount + 1
        arrOut(i, 0) = testResult.methodName
        If testResult.passed Then arrOut(i, 1) = "Passed" Else arrOut(i, 1) = "Failed"
        arrOut(i, 2) = testResult.failDetails
    Next i
    '
    With New frmTestResults
        .SetSummary failedCount, totalCount, secondsDuration
        .TestList = arrOut
        .CodeModuleName = MODULE_NAME
        .Show
    End With
End Sub

'>>>>>>>>>>>>>>
'Assert methods
'>>>>>>>>>>>>>>

'*******************************************************************************
'Raises error ERR_ASSERT_FAILED if 'boolExpression' is False with the
'   Err.Description set to 'detailsIfFalse' value
'*******************************************************************************
Private Sub AssertIsTrue(ByVal boolExpression As Boolean _
    , Optional ByVal detailsIfFalse As String _
)
    If Not boolExpression Then Err.Raise ERR_ASSERT_FAILED, , detailsIfFalse
End Sub

'*******************************************************************************
'Raises error ERR_ASSERT_FAILED if 'boolExpression' is True with the
'   Err.Description set to 'detailsIfFalse' value
'*******************************************************************************
Private Sub AssertIsFalse(ByVal boolExpression As Boolean _
    , Optional ByVal detailsIfFalse As String _
)
    If boolExpression Then Err.Raise ERR_ASSERT_FAILED, , detailsIfFalse
End Sub

'*******************************************************************************
'Raises error ERR_ASSERT_FAILED if vActual <> vExpected with the
'   Err.Description set to 'detailsIfFalse' value
'*******************************************************************************
Private Sub AssertAreEqual(ByVal vExpected As Variant, ByVal vActual As Variant _
    , Optional ByVal detailsIfFalse As String _
)
    If Not vActual = vExpected Then
        Err.Raise ERR_ASSERT_FAILED, , "Expected " & vExpected & " but got " _
            & vActual & ". " & detailsIfFalse
    End If
End Sub

'*******************************************************************************
'Raises error ERR_ASSERT_FAILED with the Err.Description set to 'failDetails'
'*******************************************************************************
Private Sub AssertFail(ByVal failDetails As String)
    Err.Raise ERR_ASSERT_FAILED, , failDetails
End Sub

'>>>>>>>>>>>>>>>>>>
'Converting to text
'>>>>>>>>>>>>>>>>>>

'*******************************************************************************
'Converts a multidimensional array to comma separated values and adds square
'   brackets [] around each dimension
'*******************************************************************************
Private Function ArrayToCSV(arr As Variant _
    , Optional ByVal delimiter As String = "," _
) As String
    Const fullMethodName As String = MODULE_NAME & ".ArrayToCSV"
    '
    If LibArrayTools.GetArrayDimsCount(arr) = 0 Then
        Err.Raise 5, fullMethodName, "Invalid or Uninitialized Array"
    End If
    ArrayToCSV = CollectionToCSV(LibArrayTools.NDArrayToCollections(arr), delimiter)
End Function

'*******************************************************************************
'Converts a collection to comma separated values and adds square brackets []
'   around main collection and all nested collections
'*******************************************************************************
Private Function CollectionToCSV(ByVal coll As Collection _
    , Optional ByVal delimiter As String = "," _
) As String
    Const fullMethodName As String = MODULE_NAME & ".CollectionToCSV"
    Dim s As String
    '
    If coll Is Nothing Then
        CollectionToCSV = "Nothing"
        Exit Function
    End If
    '
    Dim v As Variant
    Dim tColl As Collection
    '
    s = "["
    For Each v In coll
        If VBA.IsObject(v) Then
            If v Is Nothing Then
                s = s & "Nothing"
            ElseIf TypeOf v Is VBA.Collection Then
                Set tColl = v
                s = s & CollectionToCSV(tColl)
            Else
                Err.Raise 5, fullMethodName, "Object type not supported"
            End If
        Else
            Select Case VBA.VarType(v)
            Case vbNull
                s = s & "Null"
            Case vbEmpty
                s = s & "Empty"
            Case vbString
                s = s & """" & v & """"
            Case vbError, vbBoolean
                s = s & v
            Case vbError, vbBoolean, vbByte, vbInteger, vbLong, 20 _
               , vbCurrency, vbDecimal, vbDouble, vbSingle
                s = s & v
            Case vbDate
                s = s & CDbl(v)
            Case vbArray To vbArray + vbUserDefinedType
                s = s & ArrayToCSV(v, delimiter)
            Case vbUserDefinedType
                Err.Raise 5, fullMethodName, "User defined types not supported"
            Case vbDataObject
                Err.Raise 5, fullMethodName, "Object type not supported"
            End Select
        End If
        s = s & delimiter
    Next v
    If coll.Count > 0 Then s = Left$(s, Len(s) - Len(delimiter))
    CollectionToCSV = s & "]"
End Function

'>>>>>>>>>>>>>
'Test template
'>>>>>>>>>>>>>

'###############################################################################
'Testing LibArrayTools.
'###############################################################################
Private Function Test() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "Test"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    expectedError = NewExpectedError(5)
    'Test
    If Not expectedError.wasRaised Then AssertFail "Err not raised"
    '
    testResult.passed = True
ExitTest:
    Test = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'>>>>>>>>>>>>>>>
'All tests below
'>>>>>>>>>>>>>>>

'###############################################################################
'Testing LibArrayTools.GetArrayDimsCount
'###############################################################################
Private Function TestGetArrayDimsCount() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestGetArrayDimsCount"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    AssertAreEqual 1, LibArrayTools.GetArrayDimsCount(ZeroLengthArray())
    '
    Dim arr1(0 To 1, 0 To 2) As Variant
    AssertAreEqual 2, LibArrayTools.GetArrayDimsCount(arr1)
    '
    Dim arr2(0 To 0, 0 To 1, 0 To 2, 0 To 5) As Double
    AssertAreEqual 4, LibArrayTools.GetArrayDimsCount(arr2)
    '
    Dim arr As Variant: ReDim arr(0)
    AssertAreEqual 1, LibArrayTools.GetArrayDimsCount(arr)
    '
    ReDim arr(0, 0, 0 To 1, 0, 0)
    AssertAreEqual 5, LibArrayTools.GetArrayDimsCount(arr)
    '
    arr = 1
    AssertAreEqual 0, LibArrayTools.GetArrayDimsCount(arr)
    '
    AssertAreEqual 0, LibArrayTools.GetArrayDimsCount(Nothing)
    '
    testResult.passed = True
ExitTest:
    TestGetArrayDimsCount = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.GetArrayElemCount
'###############################################################################
Private Function TestGetArrayElemCount() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestGetArrayElemCount"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    AssertAreEqual 0, LibArrayTools.GetArrayElemCount(ZeroLengthArray())
    '
    Dim arr1(0 To 1, 0 To 2) As Variant
    AssertAreEqual 6, LibArrayTools.GetArrayElemCount(arr1)
    '
    Dim arr2(0 To 0, 0 To 1, 0 To 2, 0 To 5) As Double
    AssertAreEqual 36, LibArrayTools.GetArrayElemCount(arr2)
    '
    Dim arr As Variant: ReDim arr(0)
    AssertAreEqual 1, LibArrayTools.GetArrayElemCount(arr)
    '
    ReDim arr(0, 0, 0, 0, 0)
    AssertAreEqual 1, LibArrayTools.GetArrayElemCount(arr)
    '
    AssertAreEqual 0, LibArrayTools.GetArrayElemCount(1)
    AssertAreEqual 0, LibArrayTools.GetArrayElemCount(Nothing)
    AssertAreEqual 0, LibArrayTools.GetArrayElemCount("icb")
    '
    testResult.passed = True
ExitTest:
    TestGetArrayElemCount = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.NDArrayToCollections
'###############################################################################
Private Function TestNDArrayToCollections() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestNDArrayToCollections"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.NDArrayToCollections arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised"
    '
    Dim arr1D(1 To 4) As Long
    Dim arr2D(1 To 3, 1 To 2) As Long
    Dim arr3D(1 To 2, 1 To 3, 1 To 1) As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long
    Dim coll As Collection
    Dim coll2 As Collection
    Dim coll3 As Collection
    '
    'Populate a 1 Dimensional Array with consecutive numbers row-wise
    n = 0
    For i = LBound(arr1D, 1) To UBound(arr1D, 1)
        n = n + 1
        arr1D(i) = n
    Next i
    '
    'Check 1D Array conversion
    Set coll = LibArrayTools.NDArrayToCollections(arr1D)
    For i = 1 To coll.Count
        AssertAreEqual i, coll(i)
    Next i
    '
    'Populate a 2 Dimensional Array with consecutive numbers row-wise
    n = 0
    For i = LBound(arr2D, 1) To UBound(arr2D, 1)
        For j = LBound(arr2D, 2) To UBound(arr2D, 2)
            n = n + 1
            arr2D(i, j) = n
        Next j
    Next i
    '
    'Check 2D Array conversion
    Set coll = LibArrayTools.NDArrayToCollections(arr2D)
    n = 0
    For i = 1 To coll.Count
        Set coll2 = coll(i)
        For j = 1 To coll2.Count
            n = n + 1
            AssertAreEqual n, coll2(j)
        Next j
    Next i
    '
    'Populate a 3 Dimensional Array with consecutive numbers row-wise
    n = 0
    For i = LBound(arr3D, 1) To UBound(arr3D, 1)
        For j = LBound(arr3D, 2) To UBound(arr3D, 2)
            For k = LBound(arr3D, 3) To UBound(arr3D, 3)
                n = n + 1
                arr3D(i, j, k) = n
            Next k
        Next j
    Next i
    '
    'Check 3D Array conversion
    Set coll = LibArrayTools.NDArrayToCollections(arr3D)
    n = 0
    For i = 1 To coll.Count
        Set coll2 = coll(i)
        For j = 1 To coll2.Count
            Set coll3 = coll2(j)
            For k = 1 To coll3.Count
                n = n + 1
                AssertAreEqual n, coll3(k)
            Next k
        Next j
    Next i
    '
    testResult.passed = True
ExitTest:
    TestNDArrayToCollections = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing CollectionToCSV. This is a testing utility and is part of this module!
'###############################################################################
Private Function TestCollectionToCSV() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestCollectionToCSV"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim coll As Collection
    Dim coll2 As Collection
    '
    AssertAreEqual "Nothing", CollectionToCSV(Nothing)
    '
    Set coll = New Collection
    coll.Add Nothing
    coll.Add Empty
    coll.Add 2
    coll.Add "test"
    coll.Add "2"
    coll.Add True
    coll.Add Null
    coll.Add Array(1, 2, 3)
    coll.Add 76900
    Set coll2 = New Collection
    coll2.Add 4
    coll2.Add False
    coll.Add coll2
    '
    AssertAreEqual vExpected:="[Nothing,Empty,2,""test"",""2"",True,Null,[1,2,3],76900,[4,False]]" _
                 , vActual:=CollectionToCSV(coll)
    '
    coll2.Add Application
    '
    expectedError = NewExpectedError(5)
    CollectionToCSV coll2
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Object not supported"
    End If
    '
    Set coll = New Collection
    coll.Add New Collection
    coll.Add ZeroLengthArray()
    coll.Add 1
    coll.Add Array(ZeroLengthArray(), ZeroLengthArray())
    '
    AssertAreEqual vExpected:="[[] [] 1 [[] []]]", vActual:=CollectionToCSV(coll, " ")
    AssertAreEqual vExpected:="[[][]1[[][]]]", vActual:=CollectionToCSV(coll, vbNullString)
    '
    testResult.passed = True
ExitTest:
    TestCollectionToCSV = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing ArrayToCSV. This is a testing utility and is part of this module!
'###############################################################################
Private Function TestArrayToCSV() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestArrayToCSV"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    ArrayToCSV arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Invalid Array"
    '
    AssertAreEqual "[]", ArrayToCSV(ZeroLengthArray())
    AssertAreEqual "[[]]", ArrayToCSV(Array(ZeroLengthArray()))
    AssertAreEqual "[[0],1]", ArrayToCSV(Array(Array(0), 1))
    '
    Dim coll As New Collection
    '
    coll.Add 4
    coll.Add False
    arr = Array(Nothing, Empty, 2, "test", "2", True, Null, Array(1, 2, 3), 76900, coll)
    '
    AssertAreEqual vExpected:="[Nothing,Empty,2,""test"",""2"",True,Null,[1,2,3],76900,[4,False]]" _
                 , vActual:=ArrayToCSV(arr)
    '
    coll.Add Application
    '
    expectedError = NewExpectedError(5)
    ArrayToCSV arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Object not supported"
    '
    arr = Array(ZeroLengthArray(), ZeroLengthArray(), 1, Array(ZeroLengthArray(), ZeroLengthArray()))
    '
    AssertAreEqual vExpected:="[[] [] 1 [[] []]]", vActual:=ArrayToCSV(arr, " ")
    AssertAreEqual vExpected:="[[][]1[[][]]]", vActual:=ArrayToCSV(arr, vbNullString)
    '
    testResult.passed = True
ExitTest:
    TestArrayToCSV = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Collection
'###############################################################################
Private Function TestCollection() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestCollection"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    AssertAreEqual vExpected:="[]" _
                 , vActual:=CollectionToCSV(LibArrayTools.Collection(), ",")
    '
    AssertAreEqual vExpected:="[1,2,3]" _
                 , vActual:=CollectionToCSV(LibArrayTools.Collection(1, 2, 3), ",")
    '
    AssertAreEqual vExpected:="[1,2,[3,4]]" _
                 , vActual:=CollectionToCSV(LibArrayTools.Collection(1, 2, Array(3, 4)), ",")
    '
    AssertAreEqual vExpected:="[1,2,[3,4]]" _
                 , vActual:=CollectionToCSV(LibArrayTools.Collection(1, 2 _
                                            , LibArrayTools.Collection(3, 4)), ",")
    '
    testResult.passed = True
ExitTest:
    TestCollection = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.CollectionHasKey
'###############################################################################
Private Function TestCollectionHasKey() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestCollectionHasKey"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim coll As Collection
    '
    AssertIsFalse LibArrayTools.CollectionHasKey(coll, "Key1"), "Key shouldn't exist"
    AssertIsFalse LibArrayTools.CollectionHasKey(Nothing, "Key2"), "Key shouldn't exist"
    '
    Set coll = New Collection
    coll.Add 1, "Key1"
    coll.Add 2, "Key2"
    '
    AssertIsTrue LibArrayTools.CollectionHasKey(coll, "Key1"), "Key should exist"
    AssertIsTrue LibArrayTools.CollectionHasKey(coll, "Key2"), "Key should exist"
    AssertIsFalse LibArrayTools.CollectionHasKey(coll, "Key3"), "Key shouldn't exist"
    '
    testResult.passed = True
ExitTest:
    TestCollectionHasKey = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.CollectionTo1DArray
'###############################################################################
Private Function TestCollectionTo1DArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestCollectionTo1DArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    expectedError = NewExpectedError(91)
    LibArrayTools.CollectionTo1DArray Nothing
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Collection not set"
    End If
    '
    Dim arr() As Variant: arr = LibArrayTools.CollectionTo1DArray(New Collection)
    AssertIsTrue LBound(arr) > UBound(arr), "Expected (0, -1) bounds"
    '
    Dim coll As New Collection
    Dim i As Long
    For i = 1 To 5
        coll.Add i
    Next i
    AssertAreEqual vExpected:="[1,2,3,4,5]" _
                 , vActual:=ArrayToCSV(LibArrayTools.CollectionTo1DArray(coll)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    coll.Add Empty
    coll.Add "test"
    AssertAreEqual vExpected:="[1,2,3,4,5,Empty,""test""]" _
                 , vActual:=ArrayToCSV(LibArrayTools.CollectionTo1DArray(coll)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:=3 _
                 , vActual:=LibArrayTools.CollectionTo1DArray(coll)(2) _
                 , detailsIfFalse:="Wrong array index"
    '
    AssertAreEqual vExpected:=3 _
                 , vActual:=LibArrayTools.CollectionTo1DArray(coll, -5)(-3) _
                 , detailsIfFalse:="Wrong array index"
    '
    testResult.passed = True
ExitTest:
    TestCollectionTo1DArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.CollectionTo2DArray
'###############################################################################
Private Function TestCollectionTo2DArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestCollectionTo2DArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    expectedError = NewExpectedError(91)
    LibArrayTools.CollectionTo2DArray Nothing, 1
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Collection not set"
    End If
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.CollectionTo2DArray New Collection, -1
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Number of columns must be positive"
    End If
    '
    Dim arr() As Variant: arr = LibArrayTools.CollectionTo2DArray(New Collection, 1)
    AssertIsTrue LBound(arr) > UBound(arr), "Expected (0, -1) bounds"
    '
    Dim coll As New Collection
    Dim i As Long
    For i = 1 To 6
        coll.Add i
    Next i
    AssertAreEqual vExpected:="[[1,2],[3,4],[5,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.CollectionTo2DArray(coll, 2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    
    '
    AssertAreEqual vExpected:="[[1,2,3],[4,5,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.CollectionTo2DArray(coll, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    coll.Add 7
    AssertAreEqual vExpected:="[[1,2,3],[4,5,6],[7,Empty,Empty]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.CollectionTo2DArray(coll, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    coll.Add Null
    coll.Add vbNullString
    AssertAreEqual vExpected:="[[1,2,3],[4,5,6],[7,Null,""""]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.CollectionTo2DArray(coll, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:=6 _
                 , vActual:=LibArrayTools.CollectionTo2DArray(coll, 3)(1, 2) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:=6 _
                 , vActual:=LibArrayTools.CollectionTo2DArray(coll, 3, -2, 2)(-1, 4) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestCollectionTo2DArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.NDArrayTo1DArray
'###############################################################################
Private Function TestNDArrayTo1DArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestNDArrayTo1DArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.NDArrayTo1DArray Nothing, columnWise
    If Not expectedError.wasRaised Then AssertFail "Err not raised. No Array"
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.NDArrayTo1DArray arr, columnWise
    If Not expectedError.wasRaised Then AssertFail "Err not raised. No Array"
    '
    Dim arr1D(1 To 10) As Long
    Dim arr2D(1 To 5, 1 To 3) As Long
    Dim arr3D(1 To 2, 1 To 3, 1 To 4) As Long
    Dim arr4D(1 To 4, 1 To 3, 1 To 2, 1 To 2) As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim n As Long
    '
    'Populate a 1 Dimensional Array with consecutive numbers row-wise
    n = 0
    For i = LBound(arr1D, 1) To UBound(arr1D, 1)
        n = n + 1
        arr1D(i) = n
    Next i
    '
    'Populate a 2 Dimensional Array with consecutive numbers row-wise
    n = 0
    For i = LBound(arr2D, 1) To UBound(arr2D, 1)
        For j = LBound(arr2D, 2) To UBound(arr2D, 2)
            n = n + 1
            arr2D(i, j) = n
        Next j
    Next i
    '
    'Populate a 3 Dimensional Array with consecutive numbers row-wise
    n = 0
    For i = LBound(arr3D, 1) To UBound(arr3D, 1)
        For j = LBound(arr3D, 2) To UBound(arr3D, 2)
            For k = LBound(arr3D, 3) To UBound(arr3D, 3)
                n = n + 1
                arr3D(i, j, k) = n
            Next k
        Next j
    Next i
    '
    'Populate a 4 Dimensional Array with consecutive numbers row-wise
    n = 0
    For i = LBound(arr4D, 1) To UBound(arr4D, 1)
        For j = LBound(arr4D, 2) To UBound(arr4D, 2)
            For k = LBound(arr4D, 3) To UBound(arr4D, 3)
                For l = LBound(arr4D, 4) To UBound(arr4D, 4)
                    n = n + 1
                    arr4D(i, j, k, l) = n
                Next l
            Next k
        Next j
    Next i
    '
    AssertAreEqual vExpected:="[1,2,3,4,5,6,7,8,9,10]" _
                 , vActual:=ArrayToCSV(LibArrayTools.NDArrayTo1DArray(arr1D, columnWise)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[1,2,3,4,5,6,7,8,9,10]" _
                 , vActual:=ArrayToCSV(LibArrayTools.NDArrayTo1DArray(arr1D, rowWise)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[1,4,7,10,13,2,5,8,11,14,3,6,9,12,15]" _
                 , vActual:=ArrayToCSV(LibArrayTools.NDArrayTo1DArray(arr2D, columnWise)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]" _
                 , vActual:=ArrayToCSV(LibArrayTools.NDArrayTo1DArray(arr2D, rowWise)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[1,13,5,17,9,21,2,14,6,18,10,22,3,15,7,19,11,23,4,16,8,20,12,24]" _
                 , vActual:=ArrayToCSV(LibArrayTools.NDArrayTo1DArray(arr3D, columnWise)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24]" _
                 , vActual:=ArrayToCSV(LibArrayTools.NDArrayTo1DArray(arr3D, rowWise)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[1,13,25,37,5,17,29,41,9,21,33,45,3,15,27,39,7,19,31,43,11,23,35,47,2,14,26,38,6,18,30,42,10,22,34,46,4,16,28,40,8,20,32,44,12,24,36,48]" _
                 , vActual:=ArrayToCSV(LibArrayTools.NDArrayTo1DArray(arr4D, columnWise)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48]" _
                 , vActual:=ArrayToCSV(LibArrayTools.NDArrayTo1DArray(arr4D, rowWise)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestNDArrayTo1DArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.ValuesToCollection
'###############################################################################
Private Function TestValuesToCollection() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestValuesToCollection"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim c As Collection
    Set c = LibArrayTools.ValuesToCollection(ZeroLengthArray(), nestNone, rowWise)
    AssertAreEqual 0, c.Count, "Collection should have no element"
    '
    Dim coll As New Collection
    Dim arr() As Variant
    Dim arr2D() As Variant
    Dim tempColl As New Collection
    '
    ReDim arr2D(1 To 2, 1 To 2)
    arr2D(1, 1) = 1
    arr2D(1, 2) = 2
    arr2D(2, 1) = 3
    arr2D(2, 2) = 4
    coll.Add arr2D
    '
    coll.Add Array(5, 6)
    '
    tempColl.Add 7
    tempColl.Add 8
    tempColl.Add Array(8.2, 8.5)
    tempColl.Add 9
    tempColl.Add New Collection
    tempColl.Add ZeroLengthArray()
    tempColl.Add 10
    tempColl.Add arr
    coll.Add tempColl
    '
    coll.Add Array(11, 12, 13, 14, 15)
    
    ReDim arr2D(1 To 2, 1 To 4)
    arr2D(1, 1) = 16
    arr2D(1, 2) = 17
    arr2D(1, 3) = 18
    arr2D(1, 4) = 19
    arr2D(2, 1) = 20
    arr2D(2, 2) = 21
    arr2D(2, 3) = 22
    arr2D(2, 4) = 23
    '
    coll.Add arr2D
    coll.Add 24
    coll.Add Array(25) 'Note this should not keep the nesting (single element array)
    '
    AssertAreEqual vExpected:="[1,2,3,4,5,6,7,8,8.2,8.5,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25]" _
                 , vActual:=CollectionToCSV(LibArrayTools.ValuesToCollection(coll, nestNone, rowWise)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[1,3,2,4,5,6,7,8,8.2,8.5,9,10,11,12,13,14,15,16,20,17,21,18,22,19,23,24,25]" _
                 , vActual:=CollectionToCSV(LibArrayTools.ValuesToCollection(coll, nestNone, columnWise)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,2,3,4],[5,6],[7,8,[8.2,8.5],9,10],[11,12,13,14,15],[16,17,18,19,20,21,22,23],24,25]" _
                 , vActual:=CollectionToCSV(LibArrayTools.ValuesToCollection(coll, nestMultiItemsOnly, rowWise)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,3,2,4],[5,6],[7,8,[8.2,8.5],9,10],[11,12,13,14,15],[16,20,17,21,18,22,19,23],24,25]" _
                 , vActual:=CollectionToCSV(LibArrayTools.ValuesToCollection(coll, nestMultiItemsOnly, columnWise)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,2,3,4],[5,6],[7,8,[8.2,8.5],9,[],[],10],[11,12,13,14,15],[16,17,18,19,20,21,22,23],24,[25]]" _
                 , vActual:=CollectionToCSV(LibArrayTools.ValuesToCollection(coll, nestAll, rowWise)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,3,2,4],[5,6],[7,8,[8.2,8.5],9,[],[],10],[11,12,13,14,15],[16,20,17,21,18,22,19,23],24,[25]]" _
                 , vActual:=CollectionToCSV(LibArrayTools.ValuesToCollection(coll, nestAll, columnWise)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    arr = Array(Array(Array(Array(Null), ZeroLengthArray(), arr)))
    '
    AssertAreEqual vExpected:="[Null]" _
                 , vActual:=CollectionToCSV(LibArrayTools.ValuesToCollection(arr, nestMultiItemsOnly, columnWise)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    AssertAreEqual vExpected:="[[[[Null],[]]]]" _
                 , vActual:=CollectionToCSV(LibArrayTools.ValuesToCollection(arr, nestAll, columnWise)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestValuesToCollection = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.IsIterable
'###############################################################################
Private Function TestIsIterable() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestIsIterable"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr()  As Variant
    Dim coll As Collection
    '
    AssertIsFalse LibArrayTools.IsIterable(arr), "Uninitialized array is not iterable"
    AssertIsFalse LibArrayTools.IsIterable(coll), "Uninstanced collection is not iterable"
    AssertIsFalse LibArrayTools.IsIterable(5), "5 is not iterable"
    AssertIsTrue LibArrayTools.IsIterable(ZeroLengthArray()), "[] is iterable"
    AssertIsTrue LibArrayTools.IsIterable(New Collection), "[] is iterable"
    AssertIsTrue LibArrayTools.IsIterable(Array(1, 2, 3)), "[1,2,3] is iterable"
    '
    ReDim arr(0)
    Set coll = New Collection
    coll.Add 1
    '
    AssertIsTrue LibArrayTools.IsIterable(arr), "[Empty] is iterable"
    AssertIsTrue LibArrayTools.IsIterable(coll), "[1] is iterable"
    '
    testResult.passed = True
ExitTest:
    TestIsIterable = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.CreateFilter
'###############################################################################
Private Function TestCreateFilter() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestCreateFilter"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim filter As FILTER_PAIR
    
    filter = LibArrayTools.CreateFilter(-55, Null)
    AssertAreEqual 0, filter.cOperator
    AssertIsTrue IsNull(filter.compValue.value_)
    AssertIsFalse filter.compValue.isIterable_
    '
    filter = LibArrayTools.CreateFilter(opEqual, ZeroLengthArray())
    AssertIsTrue 0 <> filter.cOperator
    AssertIsTrue IsArray(filter.compValue.value_)
    AssertIsTrue filter.compValue.isIterable_
    '
    filter = LibArrayTools.CreateFilter(opin, Array("A", "B", "C", "A"))
    AssertIsTrue 0 <> filter.cOperator
    AssertIsTrue filter.compValue.textKeys_.Count = 3
    '
    testResult.passed = True
ExitTest:
    TestCreateFilter = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.IsValuePassingFilter
'###############################################################################
Private Function TestIsValuePassingFilter() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestIsValuePassingFilter"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim filter As FILTER_PAIR
    '
    filter = LibArrayTools.CreateFilter(65, 1)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.IsValuePassingFilter 2, filter
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Operator is invalid"
    End If
    '
    filter = LibArrayTools.CreateFilter(opBigger, ZeroLengthArray())
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.IsValuePassingFilter 2, filter
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Filter Arrays require IN/NOT IN operator"
    End If
    '
    Dim arr() As Variant
    filter = LibArrayTools.CreateFilter(opin, arr)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.IsValuePassingFilter 2, filter
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Filter Arrays must be iterable"
    End If
    '
    filter = LibArrayTools.CreateFilter(opEqual, 2)
    AssertIsFalse LibArrayTools.IsValuePassingFilter("2", filter)
    '
    filter = LibArrayTools.CreateFilter(opEqual, 2)
    AssertIsFalse LibArrayTools.IsValuePassingFilter("bbH", filter)
    '
    filter = LibArrayTools.CreateFilter(opEqual, CDate("14/02/2017"))
    AssertIsFalse LibArrayTools.IsValuePassingFilter("14/02/2017", filter)
    '
    filter = LibArrayTools.CreateFilter(opEqual, CDate(42385.2))
    AssertIsTrue LibArrayTools.IsValuePassingFilter(42385.2, filter)
    '
    filter = LibArrayTools.CreateFilter(opEqual, Null)
    AssertIsFalse LibArrayTools.IsValuePassingFilter(Empty, filter)
    '
    filter = LibArrayTools.CreateFilter(opNotEqual, Null)
    AssertIsTrue LibArrayTools.IsValuePassingFilter(Empty, filter)
    '
    filter = LibArrayTools.CreateFilter(opEqual, vbNullString)
    AssertIsTrue LibArrayTools.IsValuePassingFilter(Empty, filter)
    '
    filter = LibArrayTools.CreateFilter(opNotEqual, "256")
    AssertIsTrue LibArrayTools.IsValuePassingFilter("255", filter)
    '
    filter = LibArrayTools.CreateFilter(opSmallerOrEqual, 256)
    AssertIsTrue LibArrayTools.IsValuePassingFilter(255, filter)
    '
    filter = LibArrayTools.CreateFilter(opBiggerOrEqual, 255)
    AssertIsTrue LibArrayTools.IsValuePassingFilter(255, filter)
    '
    filter = LibArrayTools.CreateFilter(opBiggerOrEqual, 25)
    AssertIsTrue LibArrayTools.IsValuePassingFilter(255, filter)
    '
    filter = LibArrayTools.CreateFilter(opBigger, 2550)
    AssertIsFalse LibArrayTools.IsValuePassingFilter(255, filter)
    '
    Dim coll As New Collection
    coll.Add 1
    coll.Add 2
    coll.Add 5
    coll.Add "test"
    '
    filter = LibArrayTools.CreateFilter(opin, coll)
    AssertIsTrue LibArrayTools.IsValuePassingFilter(5, filter)
    '
    filter = LibArrayTools.CreateFilter(opNotIn, coll)
    AssertIsFalse LibArrayTools.IsValuePassingFilter("test", filter)
    '
    filter = LibArrayTools.CreateFilter(opNotIn, coll)
    AssertIsTrue LibArrayTools.IsValuePassingFilter("test2", filter)
    '
    filter = LibArrayTools.CreateFilter(opin, "test")
    AssertIsTrue LibArrayTools.IsValuePassingFilter("test", filter)
    '
    filter = LibArrayTools.CreateFilter(opNotIn, "test")
    AssertIsFalse LibArrayTools.IsValuePassingFilter("test", filter)
    '
    filter = LibArrayTools.CreateFilter(opLike, "?es*")
    AssertIsTrue LibArrayTools.IsValuePassingFilter("testing", filter)
    '
    filter = LibArrayTools.CreateFilter(opLike, "?es?")
    AssertIsFalse LibArrayTools.IsValuePassingFilter("testing", filter)
    '
    filter = LibArrayTools.CreateFilter(opNotLike, "*es?")
    AssertIsFalse LibArrayTools.IsValuePassingFilter("test", filter)
    '
    filter = LibArrayTools.CreateFilter(opSmaller, True)
    AssertIsTrue LibArrayTools.IsValuePassingFilter(False, filter), "False < True (convention)"
    '
    Dim coll2 As New Collection
    '
    filter = LibArrayTools.CreateFilter(opin, coll2)
    AssertIsFalse LibArrayTools.IsValuePassingFilter("test", filter)
    '
    filter = LibArrayTools.CreateFilter(opin, Array(1, 2, 3))
    AssertIsFalse LibArrayTools.IsValuePassingFilter("test", filter)
    '
    filter = LibArrayTools.CreateFilter(opin, Array(1, 2, Null, 4))
    AssertIsTrue LibArrayTools.IsValuePassingFilter(Null, filter)
    '
    testResult.passed = True
ExitTest:
    TestIsValuePassingFilter = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.CreateFiltersArray
'###############################################################################
Private Function TestCreateFiltersArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestCreateFiltersArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.CreateFiltersArray
    If Not expectedError.wasRaised Then AssertFail "Err not raised. No values"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.CreateFiltersArray ">"
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not pairs"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.CreateFiltersArray ">>", 1
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Wrong operator"
    End If
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.CreateFiltersArray -55, 1
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Wrong operator"
    End If
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.CreateFiltersArray Empty, 1
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Wrong operator"
    End If
    '
    Dim arr() As FILTER_PAIR
    arr = LibArrayTools.CreateFiltersArray(opBigger, 3, "<=", 9, "NOT IN", Array(5, 7))
    AssertAreEqual 3, UBound(arr) - LBound(arr) + 1, "Different elements count"
    AssertAreEqual 2, arr(UBound(arr)).compValue.textKeys_.Count, "Unique keys"
    AssertIsTrue arr(UBound(arr)).compValue.isIterable_, "Array should be iterable"
    '
    testResult.passed = True
ExitTest:
    TestCreateFiltersArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Filter1DArray
'###############################################################################
Private Function TestFilter1DArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestFilter1DArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim filters() As FILTER_PAIR
    Dim arr() As Variant
    Dim arr2() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Filter1DArray Array(1, 2, 3), filters
    If Not expectedError.wasRaised Then AssertFail "Err not raised. No Filters"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Filter1DArray 5, LibArrayTools.CreateFiltersArray(Array(">", 1))
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D Array"
    '
    arr = Array(1, 2, 3, 4, 5, 6, 7, 8)
    filters = LibArrayTools.CreateFiltersArray(Array(">", 3, "<=", 7, "NOT IN", Array(4, 6)))
    AssertAreEqual vExpected:="[5,7]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Filter1DArray(arr, filters)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[5,7]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Filter1DArray(arr, filters, -3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:=-3 _
                 , vActual:=LBound(LibArrayTools.Filter1DArray(arr, filters, -3)) _
                 , detailsIfFalse:="Array doesn't have the expected lower bound"
    '
    arr = Array(1, Application, 3)
    filters = LibArrayTools.CreateFiltersArray(Array("<>", Application))
    AssertAreEqual vExpected:="[1,3]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Filter1DArray(arr, filters)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array("gg", "cat", "that's", "at", "ate")
    filters = LibArrayTools.CreateFiltersArray(Array("LIKE", "*at"))
    AssertAreEqual vExpected:="[""cat"",""at""]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Filter1DArray(arr, filters)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    
    arr = Array(1, 2, 15, 7, Empty, 25, 3, Null, "TEST")
    arr2 = Array(1, 15, 7, 25, 3, Null, 2)
    filters = LibArrayTools.CreateFiltersArray(Array("IN", arr2, ">=", 2, "LIKE", "??"))
    AssertAreEqual vExpected:="[15,25]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Filter1DArray(arr, filters)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array("A", "B", "C", "D", "E", "F", "G")
    filters = LibArrayTools.CreateFiltersArray(Array("LIKE", "[B-E]", "NOT LIKE", "[C-D]"))
    AssertAreEqual vExpected:="[""B"",""E""]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Filter1DArray(arr, filters)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestFilter1DArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.OneDArrayTo2DArray
'###############################################################################
Private Function TestOneDArrayTo2DArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestOneDArrayTo2DArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.OneDArrayTo2DArray arr, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.OneDArrayTo2DArray ZeroLengthArray(), 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. No elements"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.OneDArrayTo2DArray Array(1, 2, 3), 0
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Number of columns must be positive"
    End If
    '
    arr = Array(1, 2, 3, 4, 5, 6)
    AssertAreEqual vExpected:="[[1,2],[3,4],[5,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.OneDArrayTo2DArray(arr, 2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 3, 4, 5, 6)
    AssertAreEqual vExpected:="[[1,2,3],[4,5,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.OneDArrayTo2DArray(arr, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 3, 4, 5, 6, 7)
    AssertAreEqual vExpected:="[[1,2,3],[4,5,6],[7,Empty,Empty]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.OneDArrayTo2DArray(arr, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1)
    AssertAreEqual vExpected:="[[1]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.OneDArrayTo2DArray(arr, 1)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestOneDArrayTo2DArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Filter2DArray
'###############################################################################
Private Function TestFilter2DArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestFilter2DArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim filters() As FILTER_PAIR
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Filter2DArray arr, 1, LibArrayTools.CreateFiltersArray(">", 1)
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Expected 2D Array"
    End If
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4), 2)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Filter2DArray arr, 1, filters
    If Not expectedError.wasRaised Then AssertFail "Err not raised. No Filters"
    '
    filters = LibArrayTools.CreateFiltersArray("<", 3)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Filter2DArray arr, -1, filters
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong Column"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10), 2)
    filters = LibArrayTools.CreateFiltersArray(">=", 3, "<", 9)
    AssertAreEqual vExpected:="[[3,4],[5,6],[7,8]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Filter2DArray(arr, 0, filters)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[3,4],[5,6],[7,8]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Filter2DArray(arr, 0, filters, 9)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:=9 _
                 , vActual:=LBound(LibArrayTools.Filter2DArray(arr, 0, filters, 9), 1) _
                 , detailsIfFalse:="Array doesn't have the expected row lower bound"
    '
    arr = LibArrayTools.Filter2DArray(arr, 0, filters)
    filters = LibArrayTools.CreateFiltersArray("IN", Array(4, 6, 7))
    AssertAreEqual vExpected:="[[3,4],[5,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Filter2DArray(arr, 1, filters)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    Dim arr2() As Variant
    arr = Array(1, 2, 9, 2, Empty, 10, 3, Null, 11, 4, True, 12, 5, "Test", 13 _
        , 6, "True", 14, 7, Application, 15, 8, 4, 16)
    arr = LibArrayTools.OneDArrayTo2DArray(arr, 3)
    arr2 = Array(Null, "True", "Test", Application)
    filters = LibArrayTools.CreateFiltersArray("NOT IN", arr2)
    AssertAreEqual vExpected:="[[1,2,9],[2,Empty,10],[4,True,12],[8,4,16]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Filter2DArray(arr, 1, filters)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1), 1)
    filters = LibArrayTools.CreateFiltersArray("=", 1)
    AssertAreEqual vExpected:="[[1]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Filter2DArray(arr, 0, filters)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestFilter2DArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.FilterCollection
'###############################################################################
Private Function TestFilterCollection() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestFilterCollection"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim filters() As FILTER_PAIR
    '
    expectedError = NewExpectedError(91)
    LibArrayTools.FilterCollection Nothing, LibArrayTools.CreateFiltersArray(">", 1)
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Collection not set"
    End If
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.FilterCollection New Collection, filters
    If Not expectedError.wasRaised Then AssertFail "Err not raised. No Filters"
    '
    Dim coll As New Collection
    '
    filters = LibArrayTools.CreateFiltersArray("<", 2)
    LibArrayTools.FilterCollection coll, filters
    AssertIsTrue coll.Count = 0
    '
    Dim arr() As Variant
    '
    Set coll = LibArrayTools.Collection(1, vbNullString, 2, vbNullString, Null, vbNullString, Empty, 3, Application)
    filters = LibArrayTools.CreateFiltersArray(">=", 1)
    AssertAreEqual vExpected:="[1,2,3]" _
                 , vActual:=CollectionToCSV(LibArrayTools.FilterCollection(coll, filters)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, 2, 3)
    filters = LibArrayTools.CreateFiltersArray(">", 5)
    AssertAreEqual vExpected:="[]" _
                 , vActual:=CollectionToCSV(LibArrayTools.FilterCollection(coll, filters)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Dim coll2 As New Collection
    coll2.Add 5
    '
    Set coll = LibArrayTools.Collection(Application, ZeroLengthArray(), coll2, "test", 2)
    filters = LibArrayTools.CreateFiltersArray("NOT IN", Array(Application, "test"))
    AssertAreEqual vExpected:="[[],[5],2]" _
                 , vActual:=CollectionToCSV(LibArrayTools.FilterCollection(coll, filters)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection("A", "B", "C", "D", "E")
    filters = LibArrayTools.CreateFiltersArray("LIKE", "[B-E]", "NOT LIKE", "[C-D]")
    AssertAreEqual vExpected:="[""B"",""E""]" _
                 , vActual:=CollectionToCSV(LibArrayTools.FilterCollection(coll, filters)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestFilterCollection = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.FindTextsRow
'###############################################################################
Private Function TestFindTextsRow() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestFindTextsRow"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.FindTextsRow arr, Array()
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Array not 2D"
    End If
    '
    arr = OneDArrayTo2DArray(Array("AB", "AC", "AD", "AF", "AB", "AC", "AE", "AF", "AB", "AC", "AD", "AE"), 4)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.FindTextsRow arr, "AB"
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not iterable"
    '
    AssertAreEqual vExpected:=2 _
                 , vActual:=LibArrayTools.FindTextsRow(arr, Array("AB", "AB", "AD", "AE")) _
                 , detailsIfFalse:="Invalid row found"
    '
    AssertAreEqual vExpected:=-1 _
                 , vActual:=LibArrayTools.FindTextsRow(arr, Array("AB", "AB", "AD", "AE"), maxRowsToSearch:=2) _
                 , detailsIfFalse:="Invalid row found"
    '
    AssertAreEqual vExpected:=-1 _
                 , vActual:=LibArrayTools.FindTextsRow(arr, Array()) _
                 , detailsIfFalse:="Invalid row found"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.FindTextsRow arr, Array(Null, 5)
    If Not expectedError.wasRaised Then AssertFail "Expected text to search for"
    '
    AssertAreEqual vExpected:=1 _
                 , vActual:=LibArrayTools.FindTextsRow(arr, Array("ab", "ae")) _
                 , detailsIfFalse:="Invalid row found"
    '
    AssertAreEqual vExpected:=-1 _
                 , vActual:=LibArrayTools.FindTextsRow(arr, Array("ab", "ae"), caseSensitive:=True) _
                 , detailsIfFalse:="Invalid row found"
    '
    arr = OneDArrayTo2DArray(Array("ABCD", "EFGH", "IJKL"), 3)
    '
    AssertAreEqual vExpected:=-1 _
                 , vActual:=LibArrayTools.FindTextsRow(arr, Array("abc", "efg", "ijk")) _
                 , detailsIfFalse:="Invalid row found"
    '
    AssertAreEqual vExpected:=0 _
                 , vActual:=LibArrayTools.FindTextsRow(arr, Array("abc", "efg", "ijk"), maxCharsToMatch:=3) _
                 , detailsIfFalse:="Invalid row found"
    '
    AssertAreEqual vExpected:=0 _
                 , vActual:=LibArrayTools.FindTextsRow(arr, Array("abm", "efm", "ijm"), maxCharsToMatch:=2) _
                 , detailsIfFalse:="Invalid row found"
    '
    arr = OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), 3)
    '
    AssertAreEqual vExpected:=-1 _
                 , vActual:=LibArrayTools.FindTextsRow(arr, Array(1, 2, 3)) _
                 , detailsIfFalse:="Invalid row found"
    '
    arr = OneDArrayTo2DArray(Array("1", "2", "3", "4", "5", "6", "7", "8", "9"), 3)
    '
    AssertAreEqual vExpected:=1 _
                 , vActual:=LibArrayTools.FindTextsRow(arr, Array(4, 5, 6)) _
                 , detailsIfFalse:="Invalid row found"
    '
    testResult.passed = True
ExitTest:
    TestFindTextsRow = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.IntegerRange1D
'###############################################################################
Private Function TestIntegerRange1D() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestIntegerRange1D"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    AssertAreEqual vExpected:="[1,2,3]" _
                 , vActual:=ArrayToCSV(LibArrayTools.IntegerRange1D(1, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[1,0,-1,-2,-3]" _
                 , vActual:=ArrayToCSV(LibArrayTools.IntegerRange1D(1, -3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[-3,-2,-1,0]" _
                 , vActual:=ArrayToCSV(LibArrayTools.IntegerRange1D(-3, -0)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:=0 _
                 , vActual:=LBound(LibArrayTools.IntegerRange1D(-3, -0)) _
                 , detailsIfFalse:="Array doesn't have the expected bounds"
    '
    AssertAreEqual vExpected:=-50 _
                 , vActual:=LBound(LibArrayTools.IntegerRange1D(-3, -0, -50)) _
                 , detailsIfFalse:="Array doesn't have the expected bounds"
    '
    AssertAreEqual vExpected:=734 _
                 , vActual:=UBound(LibArrayTools.IntegerRange1D(1, 768, -33)) _
                 , detailsIfFalse:="Array doesn't have the expected bounds"
    '
    testResult.passed = True
ExitTest:
    TestIntegerRange1D = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.GetUniqueIntegers
'###############################################################################
Private Function TestGetUniqueIntegers() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestGetUniqueIntegers"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    Dim lowBound As Long
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.GetUniqueIntegers 5
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not iterable"
    '
    expectedError = NewExpectedError(9)
    lowBound = LBound(LibArrayTools.GetUniqueIntegers(Array()))
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not initialized"
    '
    arr = Array(1, Empty, 2, 3)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.GetUniqueIntegers arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong data type"
    '
    arr = Array(1, 2, , Null, 3)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.GetUniqueIntegers arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong data type"
    '
    arr = Array(1, "1", 2, 3)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.GetUniqueIntegers arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong data type"
    '
    arr = Array(1, 2, 3)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.GetUniqueIntegers arr, 5, -1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong limits"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.GetUniqueIntegers arr, 2, 4
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Outside of limits"
    '
    arr = Array(1, 2, 3)
    AssertAreEqual vExpected:="[1,2,3]" _
                 , vActual:=ArrayToCSV(LibArrayTools.GetUniqueIntegers(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 1, 2.2, 2.7, 1.5)
    AssertAreEqual vExpected:="[1,2]" _
                 , vActual:=ArrayToCSV(LibArrayTools.GetUniqueIntegers(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 1, 2.2, 2.7, 1.5, 1, 2, 3, 4, 1, 2.2, 3.5)
    AssertAreEqual vExpected:="[1,2,3,4]" _
                 , vActual:=ArrayToCSV(LibArrayTools.GetUniqueIntegers(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestGetUniqueIntegers = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.GetUniqueRows
'###############################################################################
Private Function TestGetUniqueRows() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestGetUniqueRows"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    Dim columns_() As Long
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.GetUniqueRows arr, LibArrayTools.IntegerRange1D(1, 3)
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Array is not 2D"
    End If
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3), 3)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.GetUniqueRows arr, columns_
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Indexes array is not 1D"
    End If
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.GetUniqueRows arr, LibArrayTools.IntegerRange1D(-1, 0)
    If Not expectedError.wasRaised Then
        AssertFail "Err not raised. Invalid column indexes"
    End If
    '
    columns_ = LibArrayTools.IntegerRange1D(0, 1)
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 1, 5, 6, 1, 2, 9), 3)
    '
    AssertAreEqual vExpected:="[[1,2,3],[1,5,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.GetUniqueRows(arr, columns_)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    columns_ = LibArrayTools.IntegerRange1D(0, 0)
    '
    AssertAreEqual vExpected:="[[1,2,3]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.GetUniqueRows(arr, columns_)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    columns_ = LibArrayTools.IntegerRange1D(1, 2)
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 1, 5, 6, 1, 2, 9, 2, 3, 3), 3)
    '
    AssertAreEqual vExpected:="[[1,2,3],[1,5,6],[1,2,9],[2,3,3]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.GetUniqueRows(arr, columns_)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,2,3],[1,5,6],[1,2,9],[2,3,3]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.GetUniqueRows(arr, columns_, -2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:=-2 _
                 , vActual:=LBound(LibArrayTools.GetUniqueRows(arr, columns_, -2), 1) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestGetUniqueRows = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.GetUniqueValues
'###############################################################################
Private Function TestGetUniqueValues() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestGetUniqueValues"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    Dim coll As Collection
    Dim coll2 As Collection
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.GetUniqueValues 5
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not iterable"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.GetUniqueValues arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not iterable"
    '
    AssertAreEqual vExpected:="[]" _
                 , vActual:=ArrayToCSV(LibArrayTools.GetUniqueValues(ZeroLengthArray())) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[]" _
                 , vActual:=ArrayToCSV(LibArrayTools.GetUniqueValues(New Collection)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    Set coll = LibArrayTools.NDArrayToCollections(Array(1, 2, 3))
    arr = Array(coll, 1, 2, coll, 1, Null, 3, Empty, vbNullString, 2, coll, True, "2" _
        , "True", Null, False, coll)
    '
    AssertAreEqual vExpected:="[[1,2,3],1,2,Null,3,Empty,"""",True,""2"",""True"",False]" _
                 , vActual:=ArrayToCSV(LibArrayTools.GetUniqueValues(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,2,3],1,2,Null,3,Empty,"""",True,""2"",""True"",False]" _
                 , vActual:=ArrayToCSV(LibArrayTools.GetUniqueValues(arr, 4)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:=4 _
                 , vActual:=LBound(LibArrayTools.GetUniqueValues(arr, 4)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestGetUniqueValues = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.InsertRowsAtIndex
'###############################################################################
Private Function TestInsertRowsAtIndex() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestInsertRowsAtIndex"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.InsertRowsAtIndex arr, 2, 2
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 2D"
    '
    arr = Array(1, 2, 3)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.InsertRowsAtIndex arr, 2, 2
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 2D"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6), 2)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.InsertRowsAtIndex arr, 2, -2
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong row"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.InsertRowsAtIndex arr, 2, 4
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong row"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.InsertRowsAtIndex arr, -2, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong count"
    '
    AssertAreEqual vExpected:="[[1,2],[3,4],[5,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.InsertRowsAtIndex(arr, 0, 1)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,2],[Empty,Empty],[Empty,Empty],[3,4],[5,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.InsertRowsAtIndex(arr, 2, 1)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[Empty,Empty],[1,2],[3,4],[5,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.InsertRowsAtIndex(arr, 1, 0)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,2],[3,4],[5,6],[Empty,Empty]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.InsertRowsAtIndex(arr, 1, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,2],[3,4],[Empty,Empty],[Empty,Empty],[5,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.InsertRowsAtIndex(arr, 2, 2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestInsertRowsAtIndex = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.InsertRowsAtValChange
'###############################################################################
Private Function TestInsertRowsAtValChange() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestInsertRowsAtValChange"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.InsertRowsAtValChange arr, 2, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 2D"
    '
    arr = Array(1, 2, 3)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.InsertRowsAtValChange arr, 2, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 2D"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6), 2)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.InsertRowsAtValChange arr, 2, 2
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong column"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.InsertRowsAtValChange arr, 2, -1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong column"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.InsertRowsAtValChange arr, -2, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong count"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.InsertRowsAtValChange arr, 2, 1, -1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong count"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.InsertRowsAtValChange arr, 2, 1, 1, -2
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong count"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 1, 3, 2, 3, 2, 4), 2)
    '
    AssertAreEqual vExpected:="[[1,2],[1,3],[2,3],[2,4]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.InsertRowsAtValChange(arr, 0, 1)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,2],[1,3],[2,3],[2,4]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.InsertRowsAtValChange(arr, 0, 1)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,2],[1,3],[Empty,Empty],[Empty,Empty],[2,3],[2,4]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.InsertRowsAtValChange(arr, 2, 0)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,2],[Empty,Empty],[Empty,Empty],[1,3],[2,3],[Empty,Empty],[Empty,Empty],[2,4]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.InsertRowsAtValChange(arr, 2, 1)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[Empty,Empty],[1,2],[1,3],[Empty,Empty],[Empty,Empty],[2,3],[2,4],[Empty,Empty]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.InsertRowsAtValChange(arr, 2, 0, 1, 1)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5), 1)
    '
    AssertAreEqual vExpected:="[[1],[Empty],[2],[Empty],[3],[Empty],[4],[Empty],[5]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.InsertRowsAtValChange(arr, 1, 0)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3), 1)
    '
    AssertAreEqual vExpected:="[[Empty],[Empty],[1],[Empty],[2],[Empty],[3],[Empty],[Empty],[Empty]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.InsertRowsAtValChange(arr, 1, 0, 2, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestInsertRowsAtValChange = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Is2DArrayRowEmpty
'###############################################################################
Private Function TestIs2DArrayRowEmpty() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestIs2DArrayRowEmpty"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Is2DArrayRowEmpty arr, 0, False
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 2D"
    '
    arr = Array(1, 2, 3, 4)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Is2DArrayRowEmpty arr, 0, False
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 2D"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, vbNullString, vbNullString), 2)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Is2DArrayRowEmpty arr, -1, False
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong row"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Is2DArrayRowEmpty arr, 4, False
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong row"
    '
    AssertIsFalse LibArrayTools.Is2DArrayRowEmpty(arr, 1), "Not empty"
    AssertIsFalse LibArrayTools.Is2DArrayRowEmpty(arr, 3, False), "Not empty"
    AssertIsTrue LibArrayTools.Is2DArrayRowEmpty(arr, 3, True), "Empty"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, Empty, Empty, Empty, Empty, vbNullString), 2)
    '
    AssertIsFalse LibArrayTools.Is2DArrayRowEmpty(arr, 0, False), "Not empty"
    AssertIsFalse LibArrayTools.Is2DArrayRowEmpty(arr, 0, True), "Not empty"
    AssertIsTrue LibArrayTools.Is2DArrayRowEmpty(arr, 1, False), "Empty"
    AssertIsTrue LibArrayTools.Is2DArrayRowEmpty(arr, 1, True), "Empty"
    AssertIsFalse LibArrayTools.Is2DArrayRowEmpty(arr, 2, False), "Not empty"
    AssertIsTrue LibArrayTools.Is2DArrayRowEmpty(arr, 2, True), "Empty"
    '
    testResult.passed = True
ExitTest:
    TestIs2DArrayRowEmpty = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Merge1DArrays
'###############################################################################
Private Function TestMerge1DArrays() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestMerge1DArrays"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr1() As Variant
    Dim arr2() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Merge1DArrays arr1, arr2
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Merge1DArrays arr1, ZeroLengthArray()
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Merge1DArrays ZeroLengthArray(), arr2
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D"
    '
    arr1 = ZeroLengthArray()
    arr2 = ZeroLengthArray()
    '
    AssertAreEqual vExpected:="[]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Merge1DArrays(arr1, arr2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr1 = ZeroLengthArray()
    arr2 = Array(4, 5, 6)
    '
    AssertAreEqual vExpected:="[4,5,6]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Merge1DArrays(arr1, arr2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr1 = Array(1, 2, 3)
    arr2 = ZeroLengthArray()
    '
    AssertAreEqual vExpected:="[1,2,3]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Merge1DArrays(arr1, arr2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr1 = Array(1, 2, 3)
    arr2 = Array(4, 5, 6)
    '
    AssertAreEqual vExpected:="[1,2,3,4,5,6]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Merge1DArrays(arr1, arr2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[1,2,3,4,5,6]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Merge1DArrays(arr1, arr2, 5)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:=5 _
                 , vActual:=LBound(LibArrayTools.Merge1DArrays(arr1, arr2, 5)) _
                 , detailsIfFalse:="Array doesn't have the expected lower bound"
    '
    arr1 = Array(1)
    arr2 = Array(2, 3)
    '
    AssertAreEqual vExpected:="[1,2,3]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Merge1DArrays(arr1, arr2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestMerge1DArrays = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Merge2DArrays
'###############################################################################
Private Function TestMerge2DArrays() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestMerge2DArrays"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr1() As Variant
    Dim arr2() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Merge2DArrays arr1, arr2, False
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 2D"
    '
    arr1 = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4), 2)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Merge2DArrays arr1, arr2, False
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 2D"
    '
    Erase arr1
    arr2 = LibArrayTools.OneDArrayTo2DArray(Array(5, 6), 1)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Merge2DArrays arr1, arr2, False
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 2D"
    '
    arr1 = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4), 2)
    arr2 = LibArrayTools.OneDArrayTo2DArray(Array(5, 6, 7, 8, 9, 10), 2)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Merge2DArrays arr1, arr2, False
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong rows count"
    '
    arr1 = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4), 2)
    arr2 = LibArrayTools.OneDArrayTo2DArray(Array(5, 6, 7, 8, 9, 10), 3)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Merge2DArrays arr1, arr2, True
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong columns count"
    '
    arr1 = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4), 2)
    arr2 = LibArrayTools.OneDArrayTo2DArray(Array(5, 6, 7, 8, 9, 10), 3)
    '
    AssertAreEqual vExpected:="[[1,2,5,6,7],[3,4,8,9,10]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Merge2DArrays(arr1, arr2, False)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr1 = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4), 2)
    arr2 = LibArrayTools.OneDArrayTo2DArray(Array(5, 6, 7, 8, 9, 10), 2)
    '
    AssertAreEqual vExpected:="[[1,2],[3,4],[5,6],[7,8],[9,10]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Merge2DArrays(arr1, arr2, True)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr1 = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3), 3)
    arr2 = LibArrayTools.OneDArrayTo2DArray(Array(4, 5, 6), 3)
    '
    AssertAreEqual vExpected:="[[1,2,3,4,5,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Merge2DArrays(arr1, arr2, False)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[1,2,3],[4,5,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Merge2DArrays(arr1, arr2, True)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    Dim temp() As Variant
    temp = LibArrayTools.Merge2DArrays(arr1, arr2, False, -2, 3)
    '
    AssertAreEqual vExpected:="[[1,2,3,4,5,6]]" _
                 , vActual:=ArrayToCSV(temp) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:=-2 _
                 , vActual:=LBound(temp, 1) _
                 , detailsIfFalse:="Array doesn't have the expected row bound"
    '
    AssertAreEqual vExpected:=3 _
                 , vActual:=LBound(temp, 2) _
                 , detailsIfFalse:="Array doesn't have the expected column bound"
    '
    temp = LibArrayTools.Merge2DArrays(arr1, arr2, True, -2, 3)
    '
    AssertAreEqual vExpected:="[[1,2,3],[4,5,6]]" _
                 , vActual:=ArrayToCSV(temp) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:=-2 _
                 , vActual:=LBound(temp, 1) _
                 , detailsIfFalse:="Array doesn't have the expected row bound"
    '
    AssertAreEqual vExpected:=3 _
                 , vActual:=LBound(temp, 2) _
                 , detailsIfFalse:="Array doesn't have the expected column bound"
    '
    testResult.passed = True
ExitTest:
    TestMerge2DArrays = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.OneDArrayToCollection
'###############################################################################
Private Function TestOneDArrayToCollection() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestOneDArrayToCollection"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.OneDArrayToCollection arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.OneDArrayToCollection Nothing
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D"
    '
    arr = ZeroLengthArray()
    AssertAreEqual vExpected:="[]" _
                 , vActual:=CollectionToCSV(LibArrayTools.OneDArrayToCollection(arr))
    '
    arr = Array(1, 2, 3)
    AssertAreEqual vExpected:="[1,2,3]" _
                 , vActual:=CollectionToCSV(LibArrayTools.OneDArrayToCollection(arr))
    '
    arr = Array(1, 2, Array(3, 4))
    AssertAreEqual vExpected:="[1,2,[3,4]]" _
                 , vActual:=CollectionToCSV(LibArrayTools.OneDArrayToCollection(arr))
    '
    testResult.passed = True
ExitTest:
    TestOneDArrayToCollection = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.RemoveEmptyRows
'###############################################################################
Private Function TestRemoveEmptyRows() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestReplaceEmptyInArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    Dim arr2() As Variant
    '
    arr = OneDArrayTo2DArray(Array(1, 2, Empty, 4, Empty, Empty, Empty, Empty, vbNullString, Empty, Empty, Empty), 4)
    '
    expectedError = NewExpectedError(5)
    RemoveEmptyRows arr2
    If Not expectedError.wasRaised Then AssertFail "Expected 2D Array"
    '
    arr2 = arr
    '
    LibArrayTools.RemoveEmptyRows arr2
    AssertAreEqual vExpected:="[[1,2,Empty,4],["""",Empty,Empty,Empty]]" _
                 , vActual:=ArrayToCSV(arr2) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr2 = arr
    '
    LibArrayTools.RemoveEmptyRows arr2, True
    AssertAreEqual vExpected:="[[1,2,Empty,4]]" _
                 , vActual:=ArrayToCSV(arr2) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestRemoveEmptyRows = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.ReplaceEmptyInArray
'###############################################################################
Private Function TestReplaceEmptyInArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestReplaceEmptyInArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    ReplaceEmptyInArray arr, 0
    '
    arr = Array(1, Empty, Empty, 2, 3, Empty)
    '
    ReplaceEmptyInArray arr, 7
    AssertAreEqual vExpected:="[1,7,7,2,3,7]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, Empty, Empty, 2, 3, Empty)
    '
    ReplaceEmptyInArray arr, vbNullString
    AssertAreEqual vExpected:="[1,"""","""",2,3,""""]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, Empty, Empty, 2, 3, Empty), 2)
    '
    ReplaceEmptyInArray arr, " "
    AssertAreEqual vExpected:="[[1,"" ""],["" "",2],[3,"" ""]]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, Empty, Empty, 2, 3, Empty), 3)
    '
    ReplaceEmptyInArray arr, 0
    AssertAreEqual vExpected:="[[1,0,0],[2,3,0]]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = ZeroLengthArray()
    '
    ReplaceEmptyInArray arr, 0
    AssertAreEqual vExpected:="[]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    Dim arr3D(1 To 2, 1 To 3, 1 To 1) As Variant
    arr3D(1, 1, 1) = 0
    arr3D(1, 2, 1) = 0
    arr3D(2, 2, 1) = 0
    '
    ReplaceEmptyInArray arr3D, 5
    AssertAreEqual vExpected:="[[[0],[0],[5]],[[5],[0],[5]]]" _
                 , vActual:=ArrayToCSV(arr3D) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestReplaceEmptyInArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.ReplaceNullInArray
'###############################################################################
Private Function TestReplaceNullInArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestReplaceNullInArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    ReplaceNullInArray arr, 0
    '
    arr = Array(1, Null, Null, 2, 3, Null)
    '
    ReplaceNullInArray arr, 7
    AssertAreEqual vExpected:="[1,7,7,2,3,7]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, Null, Null, 2, 3, Null)
    '
    ReplaceNullInArray arr, vbNullString
    AssertAreEqual vExpected:="[1,"""","""",2,3,""""]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, Null, Null, 2, 3, Null), 2)
    '
    ReplaceNullInArray arr, " "
    AssertAreEqual vExpected:="[[1,"" ""],["" "",2],[3,"" ""]]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, Null, Null, 2, 3, Null), 3)
    '
    ReplaceNullInArray arr, 0
    AssertAreEqual vExpected:="[[1,0,0],[2,3,0]]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = ZeroLengthArray()
    '
    ReplaceNullInArray arr, 0
    AssertAreEqual vExpected:="[]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    Dim arr3D(1 To 2, 1 To 3, 1 To 1) As Variant
    arr3D(1, 1, 1) = 0
    arr3D(1, 2, 1) = 0
    arr3D(1, 3, 1) = Null
    arr3D(2, 1, 1) = Null
    arr3D(2, 2, 1) = 0
    arr3D(2, 3, 1) = Null
    '
    ReplaceNullInArray arr3D, 5
    AssertAreEqual vExpected:="[[[0],[0],[5]],[[5],[0],[5]]]" _
                 , vActual:=ArrayToCSV(arr3D) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestReplaceNullInArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Reverse1DArray
'###############################################################################
Private Function TestReverse1DArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestReverse1DArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Reverse1DArray arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Reverse1DArray ZeroLengthArray()
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Zero-length"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Reverse1DArray Array(1, 2, 3), 0
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong group size"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Reverse1DArray Array(1, 2, 3), 2
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Size not divisible"
    '
    arr = Array(1, 2, 3, 4, 5, 6)
    LibArrayTools.Reverse1DArray arr
    AssertAreEqual vExpected:="[6,5,4,3,2,1]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 3, 4, 5, 6)
    LibArrayTools.Reverse1DArray arr, 2
    AssertAreEqual vExpected:="[5,6,3,4,1,2]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 3, 4, 5, 6)
    LibArrayTools.Reverse1DArray arr, 3
    AssertAreEqual vExpected:="[4,5,6,1,2,3]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 3, 4, 5, 6)
    LibArrayTools.Reverse1DArray arr, 6
    AssertAreEqual vExpected:="[1,2,3,4,5,6]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestReverse1DArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Reverse2DArray
'###############################################################################
Private Function TestReverse2DArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestReverse2DArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Reverse2DArray arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 2D"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4), 2)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Reverse2DArray arr, 0
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong group size"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6), 3)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Reverse2DArray arr, 2
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Size not divisible"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Reverse2DArray arr, 3, True
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Size not divisible"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, 7, 8), 4)
    LibArrayTools.Reverse2DArray arr, 2, False
    AssertAreEqual vExpected:="[[3,4,1,2],[7,8,5,6]]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, 7, 8), 4)
    LibArrayTools.Reverse2DArray arr, 1, False
    AssertAreEqual vExpected:="[[4,3,2,1],[8,7,6,5]]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, 7, 8), 2)
    LibArrayTools.Reverse2DArray arr, 1, True
    AssertAreEqual vExpected:="[[7,8],[5,6],[3,4],[1,2]]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, 7, 8), 2)
    LibArrayTools.Reverse2DArray arr, 2, True
    AssertAreEqual vExpected:="[[5,6],[7,8],[1,2],[3,4]]" _
                 , vActual:=ArrayToCSV(arr) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestReverse2DArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.ReverseCollection
'###############################################################################
Private Function TestReverseCollection() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestReverseCollection"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim coll As Collection
    '
    expectedError = NewExpectedError(91)
    LibArrayTools.ReverseCollection coll
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not set"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.ReverseCollection New Collection
    If Not expectedError.wasRaised Then AssertFail "Err not raised. No elements"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.ReverseCollection LibArrayTools.Collection(1, 2, 3), 0
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong group size"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.ReverseCollection Collection(1, 2, 3), 2
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Size not divisible"
    '
    Set coll = LibArrayTools.Collection(1, 2, 3, 4, 5, 6)
    LibArrayTools.ReverseCollection coll
    AssertAreEqual vExpected:="[6,5,4,3,2,1]" _
                 , vActual:=CollectionToCSV(coll) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, 2, 3, 4, 5, 6)
    LibArrayTools.ReverseCollection coll, 2
    AssertAreEqual vExpected:="[5,6,3,4,1,2]" _
                 , vActual:=CollectionToCSV(coll) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, 2, 3, 4, 5, 6)
    LibArrayTools.ReverseCollection coll, 3
    AssertAreEqual vExpected:="[4,5,6,1,2,3]" _
                 , vActual:=CollectionToCSV(coll) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, 2, 3, 4, 5, 6)
    LibArrayTools.ReverseCollection coll, 6
    AssertAreEqual vExpected:="[1,2,3,4,5,6]" _
                 , vActual:=CollectionToCSV(coll) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestReverseCollection = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Sequence1D
'###############################################################################
Private Function TestSequence1D() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestSequence1D"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Sequence1D 0, 1, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong term count"
    '
    AssertAreEqual vExpected:="[1,2,3,4,5]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence1D(5)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[3,4,5,6,7]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence1D(5, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[3,5,7,9,11]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence1D(5, 3, 2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[3,1,-1,-3,-5]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence1D(5, 3, -2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[3]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence1D(1, 3, -2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[-3,-10,-17,-24,-31,-38]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence1D(6, -3, -7)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[-30,-26,-22,-18]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence1D(4, -30, 4)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestSequence1D = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Sequence2D
'###############################################################################
Private Function TestSequence2D() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestSequence2D"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Sequence2D 0, 1, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong term count"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Sequence2D 1, 1, 1, 0
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong column count"
    '
    AssertAreEqual vExpected:="[[1,2],[3,4],[5,0]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence2D(5, , , 2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[3],[4],[5],[6],[7]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence2D(5, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[3,4,5,6,7]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence2D(5, 3, , 5)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[3,5,7],[9,11,13]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence2D(6, 3, 2, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[3,1],[-1,-3],[-5,-7]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence2D(6, 3, -2, 2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[3]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence2D(1, 3, -2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[-3,-10,-17],[-24,-31,-38]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence2D(6, -3, -7, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:="[[-30,-26],[-22,-18]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sequence2D(4, -30, 4, 2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestSequence2D = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.ShallowCopyCollection
'###############################################################################
Private Function TestShallowCopyCollection() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestShallowCopyCollection"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim coll As Collection
    Dim coll2 As Collection
    '
    AssertAreEqual vExpected:="Nothing" _
                 , vActual:=CollectionToCSV(LibArrayTools.ShallowCopyCollection(Nothing)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, 2, 3, New Collection, Empty, Null, ZeroLengthArray())
    Set coll2 = LibArrayTools.ShallowCopyCollection(coll)
    '
    AssertIsFalse VBA.ObjPtr(coll) = VBA.ObjPtr(coll2)
    AssertIsTrue VBA.ObjPtr(coll(4)) = VBA.ObjPtr(coll2(4))
    AssertAreEqual CollectionToCSV(coll), CollectionToCSV(coll2)
    '
    Set coll = New Collection
    Set coll2 = LibArrayTools.ShallowCopyCollection(coll)
    '
    AssertIsFalse VBA.ObjPtr(coll) = VBA.ObjPtr(coll2)
    AssertAreEqual CollectionToCSV(coll), CollectionToCSV(coll2)
    '
    testResult.passed = True
ExitTest:
    TestShallowCopyCollection = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Slice1DArray
'###############################################################################
Private Function TestSlice1DArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestSlice1DArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Slice1DArray arr, 1, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D"
    '
    arr = Array(1, 2, 3)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Slice1DArray arr, -2, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong start"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Slice1DArray arr, 3, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong start"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Slice1DArray arr, 1, 0
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong length"
    '
    arr = Array(1, 2, 3, 4)
    AssertAreEqual vExpected:="[1,2,3,4]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Slice1DArray(arr, 0, 9)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 3, 4)
    AssertAreEqual vExpected:="[2,3,4]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Slice1DArray(arr, 1, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 3, 4)
    AssertAreEqual vExpected:="[2,3]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Slice1DArray(arr, 1, 2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 3, 4)
    AssertAreEqual vExpected:="[2,3]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Slice1DArray(arr, 1, 2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 3, 4)
    AssertAreEqual vExpected:="[2,3]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Slice1DArray(arr, 1, 2, -5)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(2, 3, 4)
    AssertAreEqual vExpected:=-5 _
                 , vActual:=LBound(LibArrayTools.Slice1DArray(arr, 0, 1, -5)) _
                 , detailsIfFalse:="Array doesn't have the expected bound"
    '
    testResult.passed = True
ExitTest:
    TestSlice1DArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Slice2DArray
'###############################################################################
Private Function TestSlice2DArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestSlice2DArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Slice2DArray arr, 1, 1, 1, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4), 2)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Slice2DArray arr, -2, 1, 1, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong start"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Slice2DArray arr, 2, 1, 1, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong start"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Slice2DArray arr, 1, -2, 1, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong start"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Slice2DArray arr, 1, 2, 1, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong start"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Slice2DArray arr, 1, 1, 0, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong height"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Slice2DArray arr, 1, 1, 1, 0
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong width"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, 7, 8), 4)
    AssertAreEqual vExpected:="[[2,3],[6,7]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Slice2DArray(arr, 0, 1, 2, 2)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, 7, 8), 2)
    AssertAreEqual vExpected:="[[3],[5]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Slice2DArray(arr, 1, 0, 2, 1)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, 7, 8), 8)
    AssertAreEqual vExpected:="[[1,2,3]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Slice2DArray(arr, 0, 0, 1, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, 7, 8), 1)
    AssertAreEqual vExpected:="[[5],[6],[7],[8]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Slice2DArray(arr, 4, 0, 9, 3)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    Dim temp As Variant
    temp = LibArrayTools.Slice2DArray(arr, 4, 0, 9, 3, 99, -99)
    '
    AssertAreEqual vExpected:="[[5],[6],[7],[8]]" _
                 , vActual:=ArrayToCSV(temp) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    AssertAreEqual vExpected:=99 _
                 , vActual:=LBound(temp, 1) _
                 , detailsIfFalse:="Array doesn't have the expected row bound"
    '
    AssertAreEqual vExpected:=-99 _
                 , vActual:=LBound(temp, 2) _
                 , detailsIfFalse:="Array doesn't have the expected column bound"
    '
    testResult.passed = True
ExitTest:
    TestSlice2DArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.SliceCollection
'###############################################################################
Private Function TestSliceCollection() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestSliceCollection"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim coll As Collection
    '
    expectedError = NewExpectedError(91)
    LibArrayTools.SliceCollection Nothing, 1, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not set"
    '
    Set coll = LibArrayTools.Collection(1, 2, 3)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.SliceCollection coll, 0, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong start"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.SliceCollection coll, 4, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong start"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.SliceCollection New Collection, 1, 1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong start"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.SliceCollection coll, 1, 0
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong length"
    '
    Set coll = LibArrayTools.Collection(1, 2, 3, 4)
    AssertAreEqual vExpected:="[1,2,3,4]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SliceCollection(coll, 1, 9)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, 2, 3, 4)
    AssertAreEqual vExpected:="[2,3,4]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SliceCollection(coll, 2, 3)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, 2, 3, 4)
    AssertAreEqual vExpected:="[2,3]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SliceCollection(coll, 2, 2)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(2, 3, 4)
    AssertAreEqual vExpected:="[2]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SliceCollection(coll, 1, 1)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestSliceCollection = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Sort1DArray
'###############################################################################
Private Function TestSort1DArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestSort1DArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Sort1DArray arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D"
    '
    arr = ZeroLengthArray()
    AssertAreEqual vExpected:="[]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(5)
    AssertAreEqual vExpected:="[5]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 5, 2, 3, 4, 3, 6)
    AssertAreEqual vExpected:="[1,2,2,3,3,4,5,6]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(5, 4, 3, 2, 1)
    AssertAreEqual vExpected:="[1,2,3,4,5]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, Null, Empty, vbNullString, "test", 5, 1, Empty, Array(1, 2, 3))
    AssertAreEqual vExpected:="[1,1,5,"""",""test"",Null,[1,2,3],Empty,Empty]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, Null, Empty, vbNullString, "test", 5, 1, Empty, Array(1, 2, 3))
    AssertAreEqual vExpected:="[[1,2,3],Null,""test"","""",5,1,1,Empty,Empty]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr, False)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, "2", 2, Null, Empty, vbNullString, "test", 5, "4", 1)
    AssertAreEqual vExpected:="[1,1,""2"",2,""4"",5,"""",""test"",Null,Empty]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr, True, True)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, "2", 2, Null, Empty, vbNullString, "test", 5, "4", 1)
    AssertAreEqual vExpected:="[1,1,2,5,"""",""2"",""4"",""test"",Null,Empty]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr, True, False)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, "2", 2, Null, Empty, vbNullString, "test", 5, "4", 1)
    AssertAreEqual vExpected:="[Null,""test"","""",5,""4"",""2"",2,1,1,Empty]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr, False, True)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, "2", 2, Null, Empty, vbNullString, "test", 5, "4", 1)
    AssertAreEqual vExpected:="[Null,""test"",""4"",""2"","""",5,2,1,1,Empty]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr, False, False)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array("bB", "aa", "Ab", "Aa", "ba", "cc", "CC")
    AssertAreEqual vExpected:="[""aa"",""Aa"",""Ab"",""ba"",""bB"",""cc"",""CC""]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr, True, , False)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array("bB", "aa", "Ab", "Aa", "ba", "cc", "CC")
    AssertAreEqual vExpected:="[""Aa"",""Ab"",""CC"",""aa"",""bB"",""ba"",""cc""]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr, True, , True)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array("bB", "aa", "Ab", "Aa", "ba", "cc", "CC")
    AssertAreEqual vExpected:="[""cc"",""CC"",""bB"",""ba"",""Ab"",""aa"",""Aa""]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr, False, , False)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array("bB", "aa", "Ab", "Aa", "ba", "cc", "CC")
    AssertAreEqual vExpected:="[""cc"",""ba"",""bB"",""aa"",""CC"",""Ab"",""Aa""]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort1DArray(arr, False, , True)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestSort1DArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.Sort2DArray
'###############################################################################
Private Function TestSort2DArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestSort2DArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Sort2DArray arr, 0
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 2D"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6), 2)
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Sort2DArray arr, -1
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong column"
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.Sort2DArray arr, 2
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Wrong column"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1), 1)
    AssertAreEqual vExpected:="[[1]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort2DArray(arr, 0)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array("4", 1, 1, 2, Null, 3, Empty, 4, 2, 5, "2", 6, 2, 7, 4, 8), 2)
    AssertAreEqual vExpected:="[[1,2],[2,5],[2,7],[4,8],[""2"",6],[""4"",1],[Null,3],[Empty,4]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort2DArray(arr, 0, True, False)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array("4", 1, 1, 2, Null, 3, Empty, 4, 2, 5, "2", 6, 2, 7, 4, 8), 2)
    AssertAreEqual vExpected:="[[1,2],[2,5],[""2"",6],[2,7],[""4"",1],[4,8],[Null,3],[Empty,4]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort2DArray(arr, 0, True, True)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array("4", 1, 1, 2, Null, 3, Empty, 4, 2, 5, "2", 6, 2, 7, 4, 8), 2)
    AssertAreEqual vExpected:="[[Null,3],[""4"",1],[""2"",6],[4,8],[2,5],[2,7],[1,2],[Empty,4]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort2DArray(arr, 0, False, False)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array("4", 1, 1, 2, Null, 3, Empty, 4, 2, 5, "2", 6, 2, 7, 4, 8), 2)
    AssertAreEqual vExpected:="[[Null,3],[""4"",1],[4,8],[2,5],[""2"",6],[2,7],[1,2],[Empty,4]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort2DArray(arr, 0, False, True)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array("bB", 1, "aa", 2, "Ab", 3, "Aa", 4, "ba", 5, "cc", 6, "CC", 7), 2)
    AssertAreEqual vExpected:="[[""aa"",2],[""Aa"",4],[""Ab"",3],[""ba"",5],[""bB"",1],[""cc"",6],[""CC"",7]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort2DArray(arr, 0, True, , False)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array("bB", 1, "aa", 2, "Ab", 3, "Aa", 4, "ba", 5, "cc", 6, "CC", 7), 2)
    AssertAreEqual vExpected:="[[""Aa"",4],[""Ab"",3],[""CC"",7],[""aa"",2],[""bB"",1],[""ba"",5],[""cc"",6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort2DArray(arr, 0, True, , True)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array("bB", 1, "aa", 2, "Ab", 3, "Aa", 4, "ba", 5, "cc", 6, "CC", 7), 2)
    AssertAreEqual vExpected:="[[""cc"",6],[""CC"",7],[""bB"",1],[""ba"",5],[""Ab"",3],[""aa"",2],[""Aa"",4]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort2DArray(arr, 0, False, , False)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array("bB", 1, "aa", 2, "Ab", 3, "Aa", 4, "ba", 5, "cc", 6, "CC", 7), 2)
    AssertAreEqual vExpected:="[[""cc"",6],[""ba"",5],[""bB"",1],[""aa"",2],[""CC"",7],[""Ab"",3],[""Aa"",4]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort2DArray(arr, 0, False, , True)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(4, 2, 3, 1, 4, 1, 2, 2, 4, 3, 1, 2, 1, 1, 2, 1, 1, 3), 2)
    AssertAreEqual vExpected:="[[1,2],[1,1],[1,3],[2,2],[2,1],[3,1],[4,2],[4,1],[4,3]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort2DArray(arr, 0)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(4, 2, 3, 1, 4, 1, 2, 2, 4, 3, 1, 2, 1, 1, 2, 1, 1, 3), 2)
    AssertAreEqual vExpected:="[[3,1],[4,1],[1,1],[2,1],[4,2],[2,2],[1,2],[4,3],[1,3]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort2DArray(arr, 1)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    'Double-sort
    arr = LibArrayTools.OneDArrayTo2DArray(Array(4, 2, 3, 1, 4, 1, 2, 2, 4, 3, 1, 2, 1, 1, 2, 1, 1, 3), 2)
    AssertAreEqual vExpected:="[[1,1],[1,2],[1,3],[2,1],[2,2],[3,1],[4,1],[4,2],[4,3]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.Sort2DArray(LibArrayTools.Sort2DArray(arr, 1), 0)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestSort2DArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.SortCollection
'###############################################################################
Private Function TestSortCollection() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestSortCollection"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim coll As Collection
    '
    expectedError = NewExpectedError(91)
    LibArrayTools.SortCollection coll
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not set"
    '
    Set coll = LibArrayTools.Collection()
    AssertAreEqual vExpected:="[]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(5)
    AssertAreEqual vExpected:="[5]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, 2, 5, 2, 3, 4, 3, 6)
    AssertAreEqual vExpected:="[1,2,2,3,3,4,5,6]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(5, 4, 3, 2, 1)
    AssertAreEqual vExpected:="[1,2,3,4,5]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, Null, Empty, vbNullString, "test", 5, 1, Empty, Array(1, 2, 3))
    AssertAreEqual vExpected:="[1,1,5,"""",""test"",Null,[1,2,3],Empty,Empty]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, Null, Empty, vbNullString, "test", 5, 1, Empty, Array(1, 2, 3))
    AssertAreEqual vExpected:="[[1,2,3],Null,""test"","""",5,1,1,Empty,Empty]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll, False)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, "2", 2, Null, Empty, vbNullString, "test", 5, "4", 1)
    AssertAreEqual vExpected:="[1,1,""2"",2,""4"",5,"""",""test"",Null,Empty]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll, True, True)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, "2", 2, Null, Empty, vbNullString, "test", 5, "4", 1)
    AssertAreEqual vExpected:="[1,1,2,5,"""",""2"",""4"",""test"",Null,Empty]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll, True, False)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, "2", 2, Null, Empty, vbNullString, "test", 5, "4", 1)
    AssertAreEqual vExpected:="[Null,""test"","""",5,""4"",""2"",2,1,1,Empty]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll, False, True)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection(1, "2", 2, Null, Empty, vbNullString, "test", 5, "4", 1)
    AssertAreEqual vExpected:="[Null,""test"",""4"",""2"","""",5,2,1,1,Empty]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll, False, False)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection("bB", "aa", "Ab", "Aa", "ba", "cc", "CC")
    AssertAreEqual vExpected:="[""aa"",""Aa"",""Ab"",""ba"",""bB"",""cc"",""CC""]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll, True, , False)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection("bB", "aa", "Ab", "Aa", "ba", "cc", "CC")
    AssertAreEqual vExpected:="[""Aa"",""Ab"",""CC"",""aa"",""bB"",""ba"",""cc""]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll, True, , True)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection("bB", "aa", "Ab", "Aa", "ba", "cc", "CC")
    AssertAreEqual vExpected:="[""cc"",""CC"",""bB"",""ba"",""Ab"",""aa"",""Aa""]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll, False, , False)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    Set coll = LibArrayTools.Collection("bB", "aa", "Ab", "Aa", "ba", "cc", "CC")
    AssertAreEqual vExpected:="[""cc"",""ba"",""bB"",""aa"",""CC"",""Ab"",""Aa""]" _
                 , vActual:=CollectionToCSV(LibArrayTools.SortCollection(coll, False, , True)) _
                 , detailsIfFalse:="Collection doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestSortCollection = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.SwapValues
'###############################################################################
Private Function TestSwapValues() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestSwapValues"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim v1 As Variant
    Dim v2 As Variant
    '
    v1 = 1
    v2 = 7
    LibArrayTools.SwapValues v1, v2
    AssertAreEqual 7, v1
    AssertAreEqual 1, v2
    '
    v1 = Empty
    v2 = Null
    LibArrayTools.SwapValues v1, v2
    AssertIsTrue IsEmpty(v2)
    AssertIsTrue IsNull(v1)
    '
    Set v1 = Application
    v2 = 5
    LibArrayTools.SwapValues v1, v2
    AssertIsTrue v2 Is Application
    AssertAreEqual 5, v1
    '
    Set v1 = Application
    v2 = ZeroLengthArray()
    LibArrayTools.SwapValues v1, v2
    AssertIsTrue v2 Is Application
    AssertIsTrue IsArray(v1)
    '
    Set v1 = Application
    Set v2 = New Collection
    LibArrayTools.SwapValues v1, v2
    AssertIsTrue v2 Is Application
    AssertIsTrue v1.Count = 0
    AssertIsTrue TypeName(v1) = "Collection"
    '
    v1 = True
    Set v2 = Application
    LibArrayTools.SwapValues v1, v2
    AssertIsTrue v2 = True
    AssertIsTrue v1 Is Application
    '
    v1 = Array(1, 2, 3)
    Set v2 = LibArrayTools.Collection(4, 5)
    LibArrayTools.SwapValues v1, v2
    AssertIsTrue TypeName(v1) = "Collection"
    AssertIsTrue v1.Count = 2
    AssertAreEqual 5, v1(2)
    AssertIsTrue IsArray(v2)
    AssertAreEqual 3, v2(2)
    '
    testResult.passed = True
ExitTest:
    TestSwapValues = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.TextArrayToIndex
'###############################################################################
Private Function TestTextArrayToIndex() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestTextArrayToIndex"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    Dim coll As Collection
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.TextArrayToIndex arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D/2D"
    '
    expectedError = NewExpectedError(5)
    arr = LibArrayTools.OneDArrayTo2DArray(Array("a", "b", "c", "d"), 2)
    LibArrayTools.TextArrayToIndex arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not row/column"
    '
    arr = ZeroLengthArray()
    AssertAreEqual vExpected:="[]" _
                 , vActual:=CollectionToCSV(LibArrayTools.TextArrayToIndex(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array("test")
    AssertAreEqual vExpected:="[0]" _
                 , vActual:=CollectionToCSV(LibArrayTools.TextArrayToIndex(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array("test", Array())
    expectedError = NewExpectedError(13)
    LibArrayTools.TextArrayToIndex arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not text"
    '
    arr = Array("test", "test")
    expectedError = NewExpectedError(457)
    LibArrayTools.TextArrayToIndex arr, False
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Duplicate"
    '
    arr = Array("test", "test", "test2")
    Set coll = LibArrayTools.TextArrayToIndex(arr, True)
    AssertAreEqual vExpected:="[0,2]" _
                 , vActual:=CollectionToCSV(coll) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    AssertAreEqual coll("test2"), 2
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array("a", "b", "c", "d"), 1)
    Set coll = LibArrayTools.TextArrayToIndex(arr)
    AssertAreEqual coll("a"), 0
    AssertAreEqual coll("d"), 3
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array("a", "b", "c", "d"), 4)
    Set coll = LibArrayTools.TextArrayToIndex(arr)
    AssertAreEqual coll("a"), 0
    AssertAreEqual coll("d"), 3
    '
    testResult.passed = True
ExitTest:
    TestTextArrayToIndex = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function

'###############################################################################
'Testing LibArrayTools.TransposeArray
'###############################################################################
Private Function TestTransposeArray() As TEST_RESULT
    Dim testResult As TEST_RESULT: testResult.methodName = "TestTransposeArray"
    Dim expectedError As EXPECTED_ERROR
    On Error GoTo ErrorHandler
    '
    Dim arr() As Variant
    '
    expectedError = NewExpectedError(5)
    LibArrayTools.TransposeArray arr
    If Not expectedError.wasRaised Then AssertFail "Err not raised. Not 1D/2D"
    '
    arr = ZeroLengthArray()
    AssertAreEqual vExpected:="[]" _
                 , vActual:=ArrayToCSV(LibArrayTools.TransposeArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1)
    AssertAreEqual vExpected:="[[1]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.TransposeArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = Array(1, 2, 3)
    AssertAreEqual vExpected:="[[1],[2],[3]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.TransposeArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1), 1)
    AssertAreEqual vExpected:="[[1]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.TransposeArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3), 3)
    AssertAreEqual vExpected:="[[1],[2],[3]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.TransposeArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3), 1)
    AssertAreEqual vExpected:="[[1,2,3]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.TransposeArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4), 2)
    AssertAreEqual vExpected:="[[1,3],[2,4]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.TransposeArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6), 2)
    AssertAreEqual vExpected:="[[1,3,5],[2,4,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.TransposeArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6), 3)
    AssertAreEqual vExpected:="[[1,4],[2,5],[3,6]]" _
                 , vActual:=ArrayToCSV(LibArrayTools.TransposeArray(arr)) _
                 , detailsIfFalse:="Array doesn't have the expected elements"
    '
    testResult.passed = True
ExitTest:
    TestTransposeArray = testResult
Exit Function
ErrorHandler:
    Select Case Err.Number
    Case ERR_ASSERT_FAILED
        testResult.failDetails = Err.Description
    Case expectedError.code_
        expectedError.wasRaised = True
        expectedError.code_ = 0
        Resume Next
    Case Else
        testResult.failDetails = "Err: #" & Err.Number & " - " & Err.Description
    End Select
    '
    testResult.passed = False
    Resume ExitTest
End Function
