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
To be written

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