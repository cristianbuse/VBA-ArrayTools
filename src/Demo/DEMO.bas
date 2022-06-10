Attribute VB_Name = "DEMO"
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

Public Sub DemoConversions()
    Dim coll As Collection
    '
    'Create a Collection from values
    Set coll = Collection(1, 2, 3, 4, 5)
    'Result:
    '   [1,2,3,4,5]
    '
    Dim arr() As Variant
    '
    'Convert a Collection to a 1D Array
    arr = CollectionTo1DArray(coll)
    'Result:
    '   [1,2,3,4,5,6]
    '
    'Convert a Collection to a 2D Array
    arr = CollectionTo2DArray(coll, 3)
    'Result:
    '   [1,2,3]
    '   [4,5,6]
    '
    'Convert 1D Array to a 2D Array
    arr = OneDArrayTo2DArray(Array(5, 2, 1, 3, 6, 1, 9, 5), 2)
    'Result:
    '   [5,2]
    '   [1,3]
    '   [6,1]
    '   [9,5]
    '
    arr = OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), 4)
    Dim arr2() As Variant
    '
    'Convert 2D Array to 1D Array
    arr2 = NDArrayTo1DArray(arr, rowWise)
    'Result:
    '   [1,2,3,4,5,6,7,8,9,10,11,12]
    arr2 = NDArrayTo1DArray(arr, columnWise)
    'Result:
    '   [1,5,9,2,6,10,3,7,11,4,8,12]
    '
    'Convert 2D Array to nested Collections
    Set coll = NDArrayToCollections(arr)
    'Result:
    '   [[1,2,3,4],[5,6,7,8],[9,10,11,12]]
    '
    'Merge two 1D arrays
    arr = Merge1DArrays(Array(1, 2, 3), Array(4, 5))
    'Result:
    '   [1,2,3,4,5]
    '
    Dim arr1() As Variant
    arr1 = OneDArrayTo2DArray(Array(1, 2, 3, 4), 2)
    arr2 = OneDArrayTo2DArray(Array(5, 6, 7, 8), 2)
    
    'Merge two 2D arrays
    arr = Merge2DArrays(arr1, arr2, False)
    'Result:
    '   [1,2,5,6]
    '   [3,4,7,8]
    arr = Merge2DArrays(arr1, arr2, True)
    'Result:
    '   [1,2]
    '   [3,4]
    '   [5,6]
    '   [7,8]
    '
    'Transpose a 2D Array
    arr = TransposeArray(arr1)
    'Result:
    '   [1,3]
    '   [2,4]
End Sub

Public Sub DemoFiltering()
    Dim arr() As Variant
    Dim coll As Collection
    Dim filters() As FILTER_PAIR
    Dim boolExpression As Boolean
    '
    'Check if a value is passing a filter
    boolExpression = IsValuePassingFilter(5, CreateFilter(">", 3))               'True
    boolExpression = IsValuePassingFilter(5, CreateFilter(">", 7))               'False
    boolExpression = IsValuePassingFilter(5, CreateFilter("IN", Array(1, 3, 5))) 'True
    boolExpression = IsValuePassingFilter("test", CreateFilter("LIKE", "?es?"))  'True
    boolExpression = IsValuePassingFilter("c", CreateFilter("LIKE", "[a-d]"))    'True
    '
    'Create array of filters
    filters = CreateFiltersArray(">", 1, "<=", 5, "NOT IN", Array(3, 4))
    '
    'Filter a 1D Array
    arr = Filter1DArray(Array(1, 2, 3, 4, 5), filters)
    'Result:
    '   [2,5]
    '
    arr = OneDArrayTo2DArray(Array(5, 2, 1, 3, 6, 1, 9, 5), 2)
    filters = CreateFiltersArray("IN", Array(1, 3, 5, 7, 9))
    '
    'Filter a 2D Array
    arr = Filter2DArray(arr, 0, filters)
    'Result:
    '   [5,2]
    '   [1,3]
    '   [9,5]
    arr = Filter2DArray(arr, 1, filters)
    'Result:
    '   [1,3]
    '   [9,5]
    '
    'Filter a Collection
    Set coll = FilterCollection(Collection("A", "B", "C", "D", "E") _
        , CreateFiltersArray("LIKE", "[B-E]", "NOT LIKE", "[C-D]"))
    'Result:
    '   ["B","E"]
End Sub

Public Sub DemoGetInformation()
    Dim coll As New Collection
    Dim boolExpression As Boolean
    '
    coll.Add 6, "Key1"
    '
    'Check if a Collection has a key
    boolExpression = CollectionHasKey(coll, "Key1") 'True
    boolExpression = CollectionHasKey(coll, "Key2") 'False
    '
    Dim arr() As Variant
    Dim arr2D(0 To 2, 0 To 3) As Variant
    Dim arr3D(1 To 3, 1 To 2, 1 To 5) As Variant
    Dim arr4D(1 To 2, 1 To 3, 1 To 4, 1 To 5) As Variant
    Dim dimensionsCount As Long
    Dim elementsCount As Long
    '
    'Get the number of dimensions for an array
    dimensionsCount = GetArrayDimsCount(7)       '0
    dimensionsCount = GetArrayDimsCount(arr)     '0
    dimensionsCount = GetArrayDimsCount(Array()) '1
    dimensionsCount = GetArrayDimsCount(arr2D)   '2
    dimensionsCount = GetArrayDimsCount(arr3D)   '3
    dimensionsCount = GetArrayDimsCount(arr4D)   '4
    '
    'Get the number of elements for an array
    elementsCount = GetArrayElemCount(5)              '0
    elementsCount = GetArrayElemCount(arr)            '0
    elementsCount = GetArrayElemCount(Array(1, 5, 6)) '3
    elementsCount = GetArrayElemCount(arr2D)          '12
    elementsCount = GetArrayElemCount(arr3D)          '30
    elementsCount = GetArrayElemCount(arr4D)          '120
    '
    'Check if a variant support For...Each loop
    boolExpression = IsIterable(arr)     'False
    boolExpression = IsIterable(Array()) 'True
    boolExpression = IsIterable(coll)    'True
    boolExpression = IsIterable(Nothing) 'False
End Sub

Public Sub DemoSort()
    Dim arr() As Variant
    '
    'Sort 1D Array with default options
    arr = Sort1DArray(Array(1, "2", 2, Null, Empty, vbNullString, "test", 5, "4", 1))
    'Result:
    '   [1,1,"2",2,"4",5,"","test",Null,Empty]
    '
    'Sort 1D Array while considering numbers stored as text to be just text
    arr = Sort1DArray(Array(1, "2", 2, Null, Empty, vbNullString, "test", 5, "4", 1), True, False)
    'Result:
    '   [1,1,2,5,","2","4","","test",Null,Empty]
    '
    'Sort 1D Array with default options
    arr = Sort1DArray(Array("bB", "aa", "Ab", "Aa", "ba", "cc", "CC"))
    'Result:
    '   ["aa","Aa","Ab","ba","bB","cc","CC"]
    '
    'Sort 1D Array with case sensitive texts
    arr = Sort1DArray(Array("bB", "aa", "Ab", "Aa", "ba", "cc", "CC"), , , True)
    'Result:
    '   ["Aa","Ab","CC","aa","bB","ba","cc"]
    '
    arr = LibArrayTools.OneDArrayTo2DArray(Array(4, 2, 3, 1, 4, 1, 2, 2, 4, 3, 1, 2, 1, 1, 2, 1, 1, 3), 2)
    '[4,2]
    '[3,1]
    '[4,1]
    '[2,2]
    '[4,3]
    '[1,2]
    '[1,1]
    '[2,1]
    '[1,3]
    '
    'Sort a 2D Array twice by 2 different columns
    arr = Sort2DArray(Sort2DArray(arr, 1), 0)
    'Result:
    '   [1,1]
    '   [1,2]
    '   [1,3]
    '   [2,1]
    '   [2,2]
    '   [3,1]
    '   [4,1]
    '   [4,2]
    '   [4,3]
    '
    Dim coll As Collection
    '
    'Sort a Collection
    Set coll = SortCollection(Collection(1, 2, 5, 2, 3, 4, 3, 6))
    'Result:
    '   [1,2,2,3,3,4,5,6]
End Sub

Public Sub DemoSlicing()
    Dim arr() As Variant
    Dim coll As Collection
    '
    'Slicing a 1D Array
    arr = Slice1DArray(Array(1, 2, 3, 4, 5, 6), 3, 2)
    'Result:
    '   [4,5]
    arr = Slice1DArray(Array(1, 2, 3, 4, 5, 6), 2, 1)
    'Result:
    '   [3]
    arr = Slice1DArray(Array(1, 2, 3, 4, 5, 6), 4, 5)
    'Result (note that excess length is ignored):
    '   [5,6]
    '
    arr = OneDArrayTo2DArray(Array(1, 2, 3, 4, 5, 6, 7, 8), 4)
    '[1,2,3,4]
    '[5,6,7,8]
    '
    'Slicing a 2D Array
    arr = Slice2DArray(arr, 0, 1, 2, 2)
    'Result:
    '   [2,3]
    '   [6,7]
    '
    'Slicing a Collection
    Set coll = SliceCollection(Collection(1, 2, 3, 4, 5, 6), 3, 2)
    'Result:
    '   [3,4]
End Sub
