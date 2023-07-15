Attribute VB_Name = "PathExtensionsTests"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' PathExtensionsTests.bas
'
' Dependencies:
'
' Assert.cls
' PathExtensions.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Asserts that the path elements should join and create the directory. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestPathElementsShouldJoin() As Assert

    Dim p_element1 As String: p_element1 = ActiveWorkbook.path
    Dim p_element2 As String: p_element2 = "dummy"
    Dim p_element3 As String: p_element3 = "workbook"
    Dim p_fileName As String: p_fileName = "filename.txt"
    
    ' test joining without creating
    
    Dim p_expectedDummyPath As String: p_expectedDummyPath = p_element1 & "\" & p_element2
    Dim p_actualDummyPath As String: p_actualDummyPath = PathExtensions.Join(p_element1, p_element2)
    Set TestPathElementsShouldJoin = Assert.AreEqual(p_expectedDummyPath, p_actualDummyPath, "The path elements should be joined")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    Dim p_expectedPath As String: p_expectedPath = p_element1 & "\" & p_element2 & "\" & p_element3
    Dim p_expectedFilePath As String: p_expectedFilePath = p_expectedPath & "\" & p_fileName
   
    ' test joining without creating
    
    Dim p_actualPath As String: p_actualPath = PathExtensions.JoinAll(False, p_element1, p_element2, p_element3)
    Set TestPathElementsShouldJoin = Assert.AreEqual(p_expectedPath, p_actualPath, "The path elements should be joined")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test joining a file.
    
    Dim p_actualFilePath As String: p_actualFilePath = PathExtensions.JoinFile(p_actualPath, p_fileName)
    Set TestPathElementsShouldJoin = Assert.AreEqual(p_expectedFilePath, p_actualFilePath, _
        "The path path should be joined")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test deleting the folder if it exists.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.DeleteFolder(p_actualPath), _
        "The path " & p_actualPath & " should no longer exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test joining and creating.
    
    p_actualPath = PathExtensions.JoinAll(True, p_element1, p_element2, p_element3)
    
    Set TestPathElementsShouldJoin = Assert.AreEqual(p_expectedPath, p_actualPath, _
        "The path element should be joined")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test detecting the created folder.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.FolderExists(p_actualPath), _
        "The path " & p_actualPath & " should exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test creating the file.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.CreateTextFile(p_actualFilePath), _
        "The file " & p_actualFilePath & " should exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test checking if a file exists.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.FileExists(p_actualFilePath), _
        "The file " & p_actualFilePath & " should exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test deleting the file if it exists.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.DeleteFile(p_actualFilePath), _
        "The file " & p_actualFilePath & " should no longer exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test deleting the folder.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.DeleteFolder(p_actualPath), _
        "The path " & p_actualPath & " should no longer exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test deleting the dummy folder.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.DeleteFolder(p_actualDummyPath), _
        "The path " & p_actualDummyPath & " should no longer exist")

End Function

