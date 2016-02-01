VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PathFinder 
   Caption         =   "UserForm1"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7170
   OleObjectBlob   =   "PathFinder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PathFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject

Private Sub pathBox_Change()

End Sub

Private Sub UserForm_Initialize()
Static pathStr As String
Static keyStr As String
Static extStr As String
pathBox.Value = pathStr
keyBox.Value = keyStr
ExtBox.Value = extStr

End Sub
Private Sub Browse_Click()
    pathBox.Text = GetFolder("\\pwdhqr\oows\")
    End Sub

Private Sub RunButton_Click()
Dim Ctr As Integer
Dim fl As File
Dim Fldr As Folder

pathStr = pathBox.Text
keyStr = keyBox.Text
extStr = ExtBox.Text

Set Fldr = fso.GetFolder(pathBox.Text)
'start filling entries from row 2
Ctr = Cells(Rows.count, "A").End(xlUp).Row + 1
ActiveSheet.Cells(Ctr, 1) = "File Path"
'looping through Fldr
For Each fl In Fldr.Files
DoEvents
'find each xlsx file in the Fldr object, and write its path on sheet1
'object.GetExtensionName(path), returns the extension of path(file) w/o the dot, e.g., exe instead of .exe
 '   If fso.GetExtensionName(Fl.path) = ExtBox.Text And WorksheetFunction.Search(keyBox.Text, Fl.Name) Then
If InStr(1, fso.GetExtensionName(fl.path), ExtBox.Text, vbTextCompare) <> 0 And InStr(1, fl.Name, keyBox.Text, vbTextCompare) <> 0 Then
        Ctr = Ctr + 1
        ActiveSheet.Cells(Ctr, 1) = fl.path
    End If
Next fl
'when current folder has subfolders,then call the recursion function
    If Fldr.SubFolders.count > 0 Then
        Recursive_path Fldr, Ctr
    End If
 Unload Me
End Sub
Function Recursive_path(SFolder As Folder, Ctr As Integer)
'On Error GoTo/resume are powerful yet dangerous operations to handle runtime errors
'On Error GoTo ErrorHandler
Dim Sub_Fldr As Folder
'loop through each subfolder and each file inside
For Each Sub_Fldr In SFolder.SubFolders
DoEvents
For Each fl In Sub_Fldr.Files
DoEvents
'find each xlsx file, and write its path on sheet1(append new)
If InStr(1, fso.GetExtensionName(fl.path), ExtBox.Text, vbTextCompare) <> 0 And InStr(1, fl.Name, keyBox.Text, vbTextCompare) <> 0 Then
        Ctr = Ctr + 1
        ActiveSheet.Cells(Ctr, 1) = fl.path
    End If
Next fl
'when current subfolder has subfolders, then call the recurison function (itself) till the bottom level
If Sub_Fldr.SubFolders.count > 0 Then
    Recursive_path Sub_Fldr, Ctr
End If
Next Sub_Fldr
Exit Function

'ErrorHandler is the line name for GOTO statment
ErrorHandler:
MsgBox "Some Error Occurred at " & Sub_Fldr.path & vbNewLine & "Press OK To continue!"

End Function

Function GetFolder(strPath As String) As String
Dim Fldr As FileDialog
Dim sItem As String
Set Fldr = Application.FileDialog(msoFileDialogFolderPicker)
With Fldr
    .title = "Select a Folder"
    .AllowMultiSelect = True
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem
Set Fldr = Nothing
End Function



