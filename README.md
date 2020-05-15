# VBA-ArrayTools

ArrayTools is a Project that allows easy data manipulation when working with Arrays and Collections in VBA (regardless of host Application). Operations like sorting, filtering, converting, reversing, slicing are trivial using the LibArrayTools module. Additionaly, a User-Defined-Function (UDF) module is available for Microsoft Excel.

## Installation

Just import the following code modules in your VBA Project:

* **LibArrayTools.bas**
* **UDF_DataManipulation.bas** (optional - works in MS Excel interface only, with exposed User Defined Functions)

## Testing

Import the folowing code modules:
* **TestLibArrayTools.bas**
* **frmTestResults.frm**

and execute method:
```vba
TestLibArrayTools.RunAllTests
```

## Usage
```vba
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
```

## Notes
* Argument Descriptions in the Function Help for the Excel Function Arguments Dialog (fx on the formula bar or Shift + F3) are available by running:
```vba
UDF_DataManipulation.RegisterDMFunctions
```
* Download the available Demo Workbook. Each UDF is presented with examples in a separate worksheet.

## License
Copyright (C) 2012 Cristian Buse

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program. If not, see [http://www.gnu.org/licenses/](http://www.gnu.org/licenses/) or
[GPLv3](https://choosealicense.com/licenses/gpl-3.0/).