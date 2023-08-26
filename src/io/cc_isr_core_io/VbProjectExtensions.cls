VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbProjectExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   VB Project extensions. </summary>
''' <remarks> Dependencies: PathExtensions.cls, VbComponentExtesions.cls
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' Export project files
'
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>   Exports code files of the specified project to the active workbook path. </summary>
''' <remarks> About logical operations in VBA and MS Extensibility VB Component:
''' It's a non standard boolean implementation, where True is not -1. Consequently, logical operation
''' using True (-1) and False (0) fail.
''' Specifically, a <c>Not A</c> operation is the bitwise inversion of A (Boolean is stored as long).
''' If A=-1 its bits are 1111111111111111 and the inversion is 0000000000000000. But if A=1 then its bits
''' are 0000000000000001 and the inversion gets 1111111111111110 what is representing -2.
''' So <c>Not 1</c> is -2 that is True because it is not equal to 0!
''' That's why <c>Not VBComponent.Saved = True</c> (=-2) when <c>VBComponent.Saved = True</c> (=1) too!
''' <code>
''' -1 = CInt(True): 0 = CInt(False): False = Not True
''' 1 = CInt(p_vbComponent.Saved): -2 = CInt(Not p_vbComponent.Saved): True = CBool( -2)
''' </code>
''' </remarks>
''' <param name="a_project">           [VBProject, The VB Project which code files to export. </param>
''' <param name="a_useProjectName">    [Boolean, True] true to use the project name for the subfolder. </param>
''' <param name="a_subFolder">         [String, 'code'] Specifies the sub-folder were the files
'''                                    are to be stored. </param>
''' <param name="a_folderMustExist">   [Boolean, True] True if the folder must exists in order
'''                                    to prevent unintended storage into the default folder. </param>
''' <param name="a_skipSavedFile">     [Boolean, False] True to only save unsaved files. Problem is that files are
'''                                    still marked as saved if only code documents changed. </param>
Public Sub ExportCodeFiles(ByVal a_project As VBIDE.VBProject, _
        Optional ByVal a_useProjectName As Boolean = True, _
        Optional ByVal a_subFolder As String = "code", _
        Optional ByVal a_folderMustExist As Boolean = True, _
        Optional ByVal a_skipSavedFile As Boolean = False)

    Dim p_vbComponent As VBIDE.VBComponent
    Dim p_extension As String
    Dim p_fileCount As Integer
    
    ' set the subfolder tot he project name in lower case
    If a_useProjectName Then a_subFolder = VBA.LCase$(a_project.Name)
    
    ' set and, optionally, create the folder.
    Dim p_path As String: p_path = PathExtensions.GetParentFolderName(a_project.FileName)
    
    ' abort if the folder must exist but does not exist.
    If a_folderMustExist And Not PathExtensions.FolderExists(PathExtensions.Join(p_path, a_subFolder, False)) Then
        Debug.Print "No files were exported because the path '" & p_path & "' must exist."
        Exit Sub
    End If
    
    p_path = PathExtensions.Join(p_path, a_subFolder, True)
    
    Dim p_filePath As String
    Dim p_destPath As String
    
    For Each p_vbComponent In a_project.VBComponents
        DoEvents
        ' Because VB Component is Type 100 document, 1 is True in this boolean implementation; -2 is false.
        If Not (a_skipSavedFile And 1 = VBA.CInt(p_vbComponent.Saved)) Then
            Dim p_name As String
            p_name = p_vbComponent.Name
            p_extension = VbComponentExtensions.GetFileExtension(p_vbComponent)
            If p_extension <> VBA.vbNullString Then
                p_destPath = Me.GetWorkbookPath(p_vbComponent, p_path, "workbook")
                p_filePath = PathExtensions.JoinFile(p_destPath, p_name & p_extension)
                p_vbComponent.Export FileName:=p_filePath
                p_fileCount = p_fileCount + 1
            End If
        End If
    Next p_vbComponent
    
    Debug.Print "Exported " & VBA.CStr(p_fileCount) & " files to " & p_path
    
End Sub

''' <summary>   Gets the destination path for storing the project files and the sub path
''' for storing workbook and worksheet files. <summary>
''' <param name="a_component">      [<see cref="VBComponent"/>] A VB Component. </param>
''' <param name="a_path">           [String] The top path. </param>
''' <param name="a_workbookPath">   [String] The workbook and sheets sub path. </param>
Public Function GetWorkbookPath(ByVal a_component As VBIDE.VBComponent, ByVal a_path As String, _
        ByVal a_subPath As String)
    If VbComponentExtensions.IsSheet(a_component) Or _
        VbComponentExtensions.IsWorkbook(a_component) Or _
        VbComponentExtensions.IsChart(a_component) Or _
        (1 = InStr(1, a_component.Name, "This", vbTextCompare)) Then
        GetWorkbookPath = PathExtensions.Join(a_path, a_subPath, True)
    Else
        GetWorkbookPath = a_path
    End If
End Function

