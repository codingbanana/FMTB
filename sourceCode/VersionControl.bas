Attribute VB_Name = "VersionControl"
Public Sub ExportSourceFiles()
'https://christopherjmcclellan.wordpress.com/2014/10/10/vba-and-git/
'extract source files
'requires reference to 'Microsoft Visual Basic for Applications Extensibility 5.3 library'
Dim component As VBComponent
Dim destPath As String
destPath = ThisWorkbook.path & "\sourceCode\"
If Len(Dir(destPath, vbDirectory)) = 0 Then
    MkDir destPath
End If

For Each component In Application.VBE.ActiveVBProject.VBComponents
    If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Or component.Type = vbext_ct_MSForm Then
        component.Export destPath & component.Name & ToFileExtension(component.Type)
        Debug.Print destPath & component.Name & ToFileExtension(component.Type) & " exported."
    End If
Next

End Sub
 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String

Select Case vbeComponentType
    Case vbext_ComponentType.vbext_ct_ClassModule
    ToFileExtension = ".cls"
    Case vbext_ComponentType.vbext_ct_StdModule
    ToFileExtension = ".bas"
    Case vbext_ComponentType.vbext_ct_MSForm
    ToFileExtension = ".frm"
    Case vbext_ComponentType.vbext_ct_ActiveXDesigner
    Case vbext_ComponentType.vbext_ct_Document
    Case Else
    ToFileExtension = vbNullString
End Select

End Function
'
'Sub SaveVBACode()
''This code Exports all VBA modules
'Dim i As Integer
'Dim mName As String
'Dim Fname As String
'
'    With ThisWorkbook
'        For i = 1 To .VBProject.VBComponents.Count
'            If .VBProject.VBComponents(i).CodeModule.CountOfLines > 0 Then
'                mName = .VBProject.VBComponents(i%).CodeModule.Name
'
'                 Fname = .Path & "\" & mName & ".txt"
'                .VBProject.VBComponents(mName).Export Fname
'            End If
'        Next i
'    End With
'
'End Sub
'
'Sub ImportCodeModules()
''This code imports all VBA modules
'Dim i As Integer
'Dim ModuleName As String
'
'With ThisWorkbook.VBProject
'    For i = .VBComponents.Count To 1 Step -1
'        mName = .VBComponents(i%).CodeModule.Name
'        Fname = .Path & "\" & mName & ".txt"
'
'        If mName <> "VersionControl" Then
'            If .VBComponents(i).Type <> vbext_ct_Document Then
'                .VBComponents.Remove .VBComponents(ModuleName)
'                .VBComponents.Import Fname
'            End If
'        End If
'    Next i
'End With
'
'End Sub

