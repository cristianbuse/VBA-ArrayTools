VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestResults 
   Caption         =   "Test"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12390
   OleObjectBlob   =   "frmTestResults.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTestResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''=============================================================================
''' VBA ArrayTools
'''-----------------------------------------------
''' Copyright (C) 2012 Cristian Buse
''' https://github.com/cristianbuse/VBA-ArrayTools
'''-----------------------------------------------
''' This program is free software: you can redistribute it and/or modify
''' it under the terms of the GNU General Public License as published by
''' the Free Software Foundation, either version 3 of the License, or
''' (at your option) any later version.
'''
''' This program is distributed in the hope that it will be useful,
''' but WITHOUT ANY WARRANTY; without even the implied warranty of
''' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
''' GNU General Public License for more details.
'''
''' You should have received a copy of the GNU General Public License
''' along with this program. If not, see <https://www.gnu.org/licenses/>.
'''=============================================================================

Option Explicit

Private m_codeModuleName As String

Private Sub UserForm_Initialize()
    If Application.Left + Application.Width > 0 Then
        Me.StartUpPosition = 0
        Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
        Me.Top = Application.Top + Application.Height / 2 - Me.Height / 2
    End If
End Sub

Public Sub SetSummary(failedCount As Long, totalCount As Long, secondsDuration As Double)
    lblPassed.Visible = (failedCount = 0)
    lblFailed.Visible = Not lblPassed.Visible
    '
    lblSummary.Caption = failedCount & " failed out of " & totalCount _
            & " (" & Format$(secondsDuration, "0.000") & " seconds)"
End Sub

Public Property Let TestList(arr() As String)
    On Error Resume Next
    lboxTests.List = arr
    On Error GoTo 0
End Property

Public Property Let CodeModuleName(newVal As String)
    m_codeModuleName = newVal
End Property

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub lboxTests_Change()
    Dim hasSelection As Boolean
    '
    hasSelection = (lboxTests.ListIndex > -1)
    btnJump.Enabled = hasSelection
    If Not hasSelection Then Exit Sub
    '
    tboxSelected.Text = lboxTests.List(lboxTests.ListIndex, 2)
End Sub

'*******************************************************************************
'Jumpes to the code of the selected method
'*******************************************************************************
Private Sub lboxTests_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lboxTests.ListIndex = -1 Then Exit Sub
    JumpToMethod lboxTests.List(lboxTests.ListIndex, 0)
End Sub

'*******************************************************************************
'Jumpes to the code of the selected method
'*******************************************************************************
Private Sub btnJump_Click()
    If lboxTests.ListIndex = -1 Then Exit Sub
    JumpToMethod lboxTests.List(lboxTests.ListIndex, 0)
End Sub

'*******************************************************************************
'Jumps to the selected method in the code pane
'*******************************************************************************
Private Sub JumpToMethod(methodName As String)
    If Not IsAccessToVBProjectsOn() Then
        MsgBox "You do not have access to VBProject" & vbNewLine & vbNewLine _
            & "To turn access on, go to:" & vbNewLine & "File/Options/Trust " _
            & "Center/Trust Center Settings/Macro Settings/Developer Macro " _
            & "Settings and check ""Trust access to the VBA project object " _
            & "model"" checkbox!", vbExclamation, "Missing VB Projects Access"
        Exit Sub
    End If
    '
    Dim codeModule_ As Object
    Set codeModule_ = ThisWorkbook.VBProject.VBComponents(m_codeModuleName).CodeModule
    Dim endRow As Long
    Dim endCol As Long
    '
    If codeModule_.Find(methodName & "(", 1, 1, endRow, endCol) Then
        Me.Hide
        codeModule_.CodePane.Show
        codeModule_.CodePane.SetSelection endRow, 1, endRow, endCol
    Else
        MsgBox "Method not found", vbExclamation, "Not found"
    End If
End Sub

'*******************************************************************************
'Checks if "Trust access to the VBA project object model" is on
'*******************************************************************************
Private Function IsAccessToVBProjectsOn() As Boolean
    Dim dummyProject As Object
    '
    On Error Resume Next
    Set dummyProject = ThisWorkbook.VBProject
    IsAccessToVBProjectsOn = (Err.Number = 0)
    On Error GoTo 0
End Function
