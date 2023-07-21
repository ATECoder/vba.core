VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' WorkbookExport.cls
'
' Dependencies:
'
' PathExtensions.cls
' VbComponentExtesions.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Exports all code files to the active workbook path. </summary>
''' <para name="a_subFolder">   [Optional, 'code'] Specifies the sub-folder were the files are to be stored. </param>
Public Sub ExportCodeFiles(Optional ByVal a_subFolder As String = "code")

    Me.ExportProjectCodeFiles Excel.ActiveWorkbook.VBProject, a_subFolder
    
End Sub

''' <summary>   Exports code files of the specified project to the active workbook path. </summary>
''' <para name="project">          [VBProject, The VB Project which code files to export. </param>
''' <para name="useProjectName">   [Optional, Boolean, True] true to use the project name for the subfolder. </param>
''' <para name="subFolder">        [Optional, String, 'code'] Specifies the sub-folder were the files
'''                                are to be stored. </param>
''' <para name="mustExist">        [Optional, Boolean, True] true if the folder must exists in order to prevent
'''                                unintended storage into the default folder. </param>
Public Sub ExportProjectCodeFiles(ByVal a_project As VBIDE.VBProject, _
        Optional ByVal a_useProjectName As Boolean = True, _
        Optional ByVal a_subFolder As String = "code", _
        Optional ByVal a_mustExist As Boolean = True)

    Dim p_vbComponent As VBIDE.VBComponent
    Dim p_extension As String
    Dim p_fileCount As Integer
    
    ' set the subfolder tot he project name in lower case
    If a_useProjectName Then a_subFolder = VBA.LCase$(a_project.Name)
    
    ' set and, optionally, create the folder.
    Dim p_path As String: p_path = PathExtensions.GetParentFolderName(a_project.FileName)
    
    ' abort if the folder must exist but does not exist.
    If a_mustExist And Not PathExtensions.FolderExists(PathExtensions.Join(p_path, a_subFolder, False)) Then
        Debug.Print "No files were exported because the path '" & p_path & "' must exist."
        Exit Sub
    End If
    
    p_path = PathExtensions.Join(p_path, a_subFolder, True)
    
    Dim p_filePath As String
    Dim p_destPath As String
    
    For Each p_vbComponent In a_project.VBComponents
        DoEvents
        Dim p_name As String: p_name = p_vbComponent.Name
        p_extension = VbComponentExtensions.GetFileExtension(p_vbComponent)
        If p_extension <> VBA.vbNullString Then
            p_destPath = Me.GetWorkbookPath(p_vbComponent, p_path, "workbook")
            p_filePath = PathExtensions.JoinFile(p_destPath, p_name & p_extension)
            p_fileCount = p_fileCount + 1
            p_vbComponent.Export FileName:=p_filePath
        End If
    Next p_vbComponent
    
    Debug.Print "Exported " & CStr(p_fileCount) & " files to " & p_path
    
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

