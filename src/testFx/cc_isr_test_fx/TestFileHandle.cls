VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestFileHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' TestFileHandle.cls
'
' Dependencies:
'
' TBA
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type MyType
    FileName As String
    FileStream As TextStream
End Type

Private This As MyType


''' <summary>   Creates a temporary test file. </summary>
''' <returns>   A <see cref="TestFileHandle"/>. <returns>
Public Function CreateTestFile() As TestFileHandle
    
    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    ' Creates a specified file name and returns a TextStream object that can be used to read from
    ' or write to the file.
    Dim p_handle As TestFileHandle: Set p_handle = New TestFileHandle
    p_handle.FileName = p_fileInfo.BuildPath(ActiveWorkbook.path, p_fileInfo.GetTempName())
    Set p_handle.FileStream = p_fileInfo.CreateTextFile(p_handle.FileName, True)
    Set CreateTestFile = p_handle
End Function

''' <summary>   Deletes the given test file. </summary>
Public Sub DeleteTestFile()
    
    Dim p_fileInfo As FileSystemObject: Set p_fileInfo = New FileSystemObject
    p_fileInfo.DeleteFile Me.FileName

End Sub

''' <summary>   Deletes the test file. </summary>
''' <param name="a_handle">   <see cref="TestFileHandle"/> The handle to the file to be deleted. </param>
Public Sub Delete(ByVal a_handle As TestFileHandle)
    
    Dim p_fileInfo As FileSystemObject
    Set p_fileInfo = New Scripting.FileSystemObject
    p_fileInfo.DeleteFile a_handle.FileName

End Sub

''' <summary>   Gets the file name. </summary>
''' <value>   [String]. </value>
Public Property Get FileName() As String
    FileName = This.FileName
End Property

''' <summary>   Sets the file name. </summary>
''' <param name="a_value">   [String] The name of the file. </param>
Public Property Let FileName(ByVal a_value As String)
    This.FileName = a_value
End Property

''' <summary>   Gets the file <see cref="Scripting.TextStream"/>. </summary>
''' <value>   [TextStream]. </value>
Public Property Get FileStream() As TextStream
    Set FileStream = This.FileStream
End Property

''' <summary>   Sets the file <see cref="Scripting.TextStream"/>. </summary>
''' <param name="a_value">   [TextStream]. </param>
Public Property Set FileStream(ByVal a_value As TextStream)
    Set This.FileStream = a_value
End Property

