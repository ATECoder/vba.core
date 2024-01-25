VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Manages paths. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Trims the path from '\' and '/' suffixes. </summary>
''' <param name="a_path">   [String] The path. </param>
''' <returns>   [String]. </returns>
Public Function TrimPath(ByVal a_path As String) As String

    a_path = VBA.Trim$(a_path)
    a_path = Me.TrimRight(a_path, "/")
    a_path = Me.TrimRight(a_path, "\")
    TrimPath = a_path
    
End Function

''' <summary>   Trims the given character from the end of the given text. </summary>
''' <param name="a_text">   [String] The text to Trim. </param>
''' <param name="a_char">   [String] The character to Trim. </param>
''' <returns>   [String]. <returns>
Public Function TrimRight(ByVal a_text As String, ByVal a_char As String) As String
    
    Dim p_textLength As Long: p_textLength = VBA.Len(a_text)
    
    Dim p_lastCharacter As String
        
    Do While p_textLength > 0
        p_lastCharacter = VBA.Right$(a_text, 1)
        
        If p_lastCharacter <> a_char Then
            Exit Do
        End If
        
        a_text = VBA.Left$(a_text, p_textLength - 1)
        p_textLength = VBA.Len(a_text)
    Loop
    
    TrimRight = a_text
    
End Function

''' <summary>   Returns the parent folder path of the specified file path. </summary>
''' <param name="a_filePath">   [String] The file path. </param>
''' <returns>   [String]. </returns>
Public Function FileParentFolderPath(ByVal a_filePath As String) As String

    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    FileParentFolderPath = p_fileInfo.GetFile(a_filePath).ParentFolder.Path
    
End Function


''' <summary>   Returns true if the folder was deleted and not longer exists. </summary>
''' <param name="a_path">   [String] The path. </param>
''' <returns>   [String]. </returns>
Public Function DeleteFolder(ByVal a_path As String) As Boolean

    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    If p_fileInfo.FolderExists(a_path) Then
        p_fileInfo.DeleteFolder a_path
    End If
    DeleteFolder = Not p_fileInfo.FolderExists(a_path)
    
End Function

''' <summary>   Returns true if the folder was deleted and not longer exists. </summary>
''' <param name="a_path">   [String] The path. </param>
''' <returns>   [String]. </returns>
Public Function DeleteFile(ByVal a_path As String) As Boolean

    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    If p_fileInfo.FileExists(a_path) Then
        p_fileInfo.DeleteFile a_path
    End If
    DeleteFile = Not p_fileInfo.FileExists(a_path)
    
End Function

''' <summary>   Returns true if the folder exists. </summary>
''' <param name="a_path">   [String] The path. </param>
''' <returns>   [String]. </returns>
Public Function FolderExists(ByVal a_path As String) As Boolean

    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    FolderExists = p_fileInfo.FolderExists(a_path)
    
End Function

''' <summary>   Returns true if the file exists. </summary>
''' <param name="a_path">   [String] The path. </param>
''' <returns>   [String]. </returns>
Public Function FileExists(ByVal a_path As String) As Boolean

    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    FileExists = p_fileInfo.FileExists(a_path)
    
End Function

''' <summary>   Returns true if the test file was created. </summary>
''' <param name="a_path">   [String] The path. </param>
''' <returns>   [String]. </returns>
Public Function CreateTextFile(ByVal a_path As String) As Boolean

    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    If Not p_fileInfo.FileExists(a_path) Then
        p_fileInfo.CreateTextFile a_path
    End If
    CreateTextFile = p_fileInfo.FileExists(a_path)
    
End Function

''' <summary>   Creates the directory if it does not exist. </summary>
''' <param name="a_path">   [String] The path. </param>
''' <returns>   [String]. </returns>
Public Function CreateFolder(ByVal a_path As String) As String

    If VBA.vbNullString <> a_path Then
        a_path = Me.TrimPath(a_path)
        Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
        If Not p_fileInfo.FolderExists(a_path) Then
            p_fileInfo.CreateFolder a_path
        End If
    End If
    CreateFolder = a_path
    
End Function

''' <summary>   Joins the file to the path; creates the path if it does not exist. </summary>
''' <param name="a_path">       [String] The path. </param>
''' <param name="a_fileName">   [String] The file name. </param>
''' <param name="a_create">     [Boolean, False] If true, creates the folder for the path. </param>
''' <returns>   [String]. </returns>
Public Function JoinFile(ByVal a_path As String, ByVal a_fileName As String, _
        Optional ByVal a_create As Boolean = False) As String

    a_path = TrimPath(a_path)
    If a_create Then a_path = Me.CreateFolder(a_path)
    If VBA.vbNullString <> a_path Then a_path = a_path & "\"
    JoinFile = a_path & a_fileName

End Function

''' <summary>   Joins the sub folder to the path and creates the path if it does not exist. </summary>
''' <param name="a_path">        [String] The path. </param>
''' <param name="a_subFolder">   [String] The sub folder to join to the path. </param>
''' <param name="a_create">      [Boolean, False] If true, creates the folder for the path. </param>
''' <returns>   [String]. </returns>
Public Function Join(ByVal a_path As String, ByVal a_subFolder As String, _
        Optional ByVal a_create As Boolean = False) As String

    a_path = Me.JoinFile(a_path, a_subFolder)
    If a_create Then
        a_path = Me.CreateFolder(a_path)
    End If
    Join = a_path

End Function

''' <summary>   Joins elements to a path creating the paths as needed. </summary>
''' <param name="create">      [Boolean] If true, creates the folder for the path. </param>
''' <param name="pathParts">   [Variant] The path parts to be joined. </param>
''' <returns>   [String]. </returns>
Public Function JoinAll(ByVal a_create As Boolean, ParamArray a_pathParts() As Variant) As String
    
    Dim p_joinedPath As String: p_joinedPath = VBA.vbNullString
    If Not VBA.IsEmpty(a_pathParts) Then
        Dim p_pathPart As Variant
        For Each p_pathPart In a_pathParts
            DoEvents
            p_joinedPath = Me.Join(p_joinedPath, p_pathPart, a_create)
        Next
    End If

    JoinAll = p_joinedPath
        
End Function

''' <summary>   Gets the parent folder path from the specified path. </summary>
''' <param name="a_path">   [String] The path. </param>
''' <returns>   [String]. </returns>
Public Function GetParentFolderName(ByVal a_path As String) As String
        
    GetParentFolderName = a_path
    If VBA.vbNullString <> a_path Then
        a_path = Me.TrimPath(a_path)
        Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
        GetParentFolderName = p_fileInfo.GetParentFolderName(a_path)
    End If
    
End Function
