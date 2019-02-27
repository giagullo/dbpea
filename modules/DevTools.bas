Option Compare Database

Public Sub ExportSourceFiles(destPath As String)
 
Dim component As VBComponent
Dim filePath As String
For Each component In Application.VBE.ActiveVBProject.VBComponents
    Debug.Print component.name
    If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
        filePath = destPath & component.name & ToFileExtension(component.Type)
        Debug.Print filePath
        component.Export destPath & component.name & ToFileExtension(component.Type)
    End If
Next
 
End Sub
 
 Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
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

Public Sub ImportSourceFiles(sourcePath As String)
    Dim project As VBProject
    Set project = Application.VBE.ActiveVBProject
    Dim file As String
     
    Dim comp As VBComponent
    For Each comp In project.VBComponents
        If Not comp.name = "DevTools" And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
            project.VBComponents.Remove comp
        End If
    Next
    
    file = Dir(sourcePath)
    While (file <> vbNullString)
        Application.VBE.ActiveVBProject.VBComponents.Import sourcePath & file
        file = Dir
    Wend

End Sub