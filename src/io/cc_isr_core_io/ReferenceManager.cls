VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReferenceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Cleans up, removes and adds references. </summary>
''' <remarks>
''' <see href="https://superuser.com/questions/980823/how-to-make-excel-use-relative-paths-in-external-workbook-links"/>
''' <see href="https://stackoverflow.com/questions/9879825/how-to-add-a-reference-programmatically-using-vba"/>
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Removes any broken library references fromthe active workbook. </summary>
Public Sub RemoveWorkbookBrokenReferences()

    Me.RemoveBrokenReferences ThisWorkbook.VBProject

End Sub

''' <summary>   Checks if the project has broken references. </summary>
''' <param name="a_project">    [VBIDE.VBProject] The project which references are to be removed.  </param>
''' <returns>   [Boolean]. </returns>
Public Function HasBrokenReferences(ByVal a_project As VBIDE.VBProject) As Boolean

    HasBrokenReferences = False
    Dim p_reference As VBIDE.Reference, i As Long
     
    ' Remove any missing references
    For i = a_project.References.Count To 1 Step -1
        Set p_reference = a_project.References.Item(i)
        If p_reference.IsBroken = True Then
            HasBrokenReferences = True
            Exit For
        End If
    Next i
     
End Function

''' <summary>   Removes any broken library references. </summary>
''' <param name="a_project">    [VBIDE.VBProject] The project which references are to be removed.  </param>
Public Sub RemoveBrokenReferences(ByVal a_project As VBIDE.VBProject)

    Dim p_reference As VBIDE.Reference, i As Long
     
    ' Remove any missing references
    For i = a_project.References.Count To 1 Step -1
        Set p_reference = a_project.References.Item(i)
        If p_reference.IsBroken = True Then
            a_project.References.Remove p_reference
        End If
    Next i
     
End Sub

''' <summary>   Removes any broken library references. </summary>
''' <param name="a_referenceName">    [String] The name of the reference. </param>
''' <param name="a_project">    [VBIDE.VBProject] The project which references are to be removed.  </param>
Public Sub RemoveReference(ByVal a_project As VBIDE.VBProject, ByVal a_referenceName As String)

    Dim p_reference As VBIDE.Reference
    
    For Each p_reference In a_project.References
        If p_reference.Name = a_referenceName Then
            a_project.References.Remove p_reference
            Exit For
        End If
    Next
     
End Sub

''' <summary>   Returns true if the reference exists in the specified project. </summary>
''' <param name="a_project">    [VBIDE.VBProject] The project which references are to be found.  </param>
''' <returns>   [VBIDE.Reference] or [Nothing]. </returns>
Public Function FindReferenceByName(ByVal a_project As VBIDE.VBProject, _
        ByVal a_referenceName As String) As VBIDE.Reference

    Dim p_reference As VBIDE.Reference
    
    Set FindReferenceByName = Nothing
     
    For Each p_reference In a_project.References
        If p_reference.Name = a_referenceName Then
            Set FindReferenceByName = p_reference
            Exit For
        End If
    Next
     
End Function

''' <summary>   Add a reference to the project using the GUID for the reference library. </summary>
''' <param name="a_guid">    [String] The GUID.  </param>
''' <returns>   [VBIDE.Reference] or [Nothing]. </returns>
Public Function AddReferenceByGUID(ByVal a_project As VBIDE.VBProject, _
        ByVal a_guid As String, ByVal a_major As Long, ByVal a_minor As Long) As VBIDE.Reference
     
    Me.RemoveBrokenReferences a_project
    
    ' Add the reference
    Set AddReferenceByGUID = a_project.References.AddFromGuid(a_guid, a_major, a_minor)
   
End Function

''' <summary>   Adds a reference by name. Returns the reference if it was successfully added. </summary>
''' <param name="a_project">          [VBIDE.VBProject] The project which references are to be found.  </param>
''' <param name="a_referenceName">    [String] The name of the reference. </param>
''' <param name="a_referencePath">    [String] The reference path. </param>
''' <param name="a_force">            [Boolean, False] True to force replacing the reference if it
'''                                   exists. </param>
''' <returns>   [VBIDE.Reference] or [Nothing]. </returns>
Public Function AddReferenceByName(ByVal a_project As VBIDE.VBProject, _
        ByVal a_referenceName As String, _
        ByVal a_referencePath As String, _
        Optional ByVal a_force As Boolean = False) As VBIDE.Reference
        
    Dim p_reference As VBIDE.Reference
    
    Set p_reference = Me.FindReferenceByName(a_project, a_referenceName)
    
    Dim p_exists As Boolean
    p_exists = Not p_reference Is Nothing
    
    If Not p_reference Is Nothing Then
        ' if the existing reference has a different path than the requested path then force
        ' removal.
        a_force = 0 >= InStr(1, p_reference.FullPath, a_referencePath, vbTextCompare)
    End If
    
    If p_exists And a_force Then
        Me.RemoveReference a_project, a_referenceName
    End If
    
    Set p_reference = Me.FindReferenceByName(a_project, a_referenceName)
    p_exists = Not p_reference Is Nothing
    
    If p_exists And a_force Then
        
        ' if failed to remove then throw an exception
        Err.Raise VBA.vbObjectError + 9001, TypeName(Me) & ".AddReferenceByName", _
                  "Failed removing existing reference '" & a_referenceName & "'."
                  
    Else
    
        p_reference = a_project.References.AddFromFile(a_referencePath)
        Set AddReferenceByName = p_reference
        
        If p_reference Is Nothing Then
        
            ' if failed to add then throw an exception
            Err.Raise VBA.vbObjectError + 9001, TypeName(Me) & ".AddReferenceByName", _
                      "Failed adding reference '" & a_referenceName & "' From '" & _
                      a_referencePath & "'. Add returned an empty reference."
        End If
        
    End If

End Function

''' <summary>   Adds a project folder reference by name if the reference file exists in the
''' project folder. Returns the reference if it was successfully added. </summary>
''' <param name="a_project">          [VBIDE.VBProject] The project which references are to be found.  </param>
''' <param name="a_referenceName">    [String] The name of the reference. </param>
''' <param name="a_force">            [Boolean, False] True to force replacing the reference if it
'''                                   exists. </param>
''' <returns>   [VBIDE.Reference] or [Nothing]. </returns>
Public Function AddLocalReferenceByName(ByVal a_project As VBIDE.VBProject, _
        ByVal a_referenceName As String) As VBIDE.Reference
        
    ' get the reference path from the project path.
    ' e.g., C:\my\lib\vba\core\core\src\io\cc.isr.core.io.xlsm
    
    Dim p_projectPath As String: p_projectPath = a_project.FileName
    
    ' get the project folder from the file name.
    Dim p_folderPath As String: p_folderPath = PathExtensions.FileParentFolderPath(p_projectPath)
    
    Dim p_referencePath  As String: p_referencePath = p_folderPath & a_referenceName & ".xlsm"
    
    ' check if the file exists in the project folder.
    If PathExtensions.FileExists(p_referencePath) Then
        Set AddLocalReferenceByName = Me.AddReferenceByName(a_project, a_referenceName, p_referencePath)
    Else
        Set AddLocalReferenceByName = Nothing
    End If

End Function
