Attribute VB_Name = "PathExtensionsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Path extension methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Asserts that the path elements should join and create the directory. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestPathElementsShouldJoin() As cc_isr_Test_Fx.Assert

    Dim p_element1 As String: p_element1 = Excel.ActiveWorkbook.path
    Dim p_element2 As String: p_element2 = "dummy"
    Dim p_element3 As String: p_element3 = "workbook"
    Dim p_fileName As String: p_fileName = "filename.txt"
    
    ' test joining without creating
    
    Dim p_expectedDummyPath As String: p_expectedDummyPath = p_element1 & "\" & p_element2
    Dim p_actualDummyPath As String: p_actualDummyPath = PathExtensions.Join(p_element1, p_element2)
    Set TestPathElementsShouldJoin = cc_isr_Test_Fx.Assert.AreEqual(p_expectedDummyPath, _
                                        p_actualDummyPath, _
                                        "The path elements should be joined")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    Dim p_expectedPath As String: p_expectedPath = p_element1 & "\" & p_element2 & "\" & p_element3
    Dim p_expectedFilePath As String: p_expectedFilePath = p_expectedPath & "\" & p_fileName
   
    ' test joining without creating
    
    Dim p_actualPath As String: p_actualPath = PathExtensions.JoinAll(False, p_element1, p_element2, p_element3)
    Set TestPathElementsShouldJoin = cc_isr_Test_Fx.Assert.AreEqual(p_expectedPath, _
                                        p_actualPath, _
                                        "The path elements should be joined")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test joining a file.
    
    Dim p_actualFilePath As String: p_actualFilePath = PathExtensions.JoinFile(p_actualPath, p_fileName)
    
    Set TestPathElementsShouldJoin = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFilePath, _
                                        p_actualFilePath, "The path path should be joined")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test deleting the folder if it exists.
    
    Set TestPathElementsShouldJoin = Assert.IsTrue(PathExtensions.DeleteFolder(p_actualPath), _
        "The path " & p_actualPath & " should no longer exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test joining and creating.
    
    p_actualPath = PathExtensions.JoinAll(True, p_element1, p_element2, p_element3)
    
    Set TestPathElementsShouldJoin = cc_isr_Test_Fx.Assert.AreEqual(p_expectedPath, _
                                        p_actualPath, "The path element should be joined")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test detecting the created folder.
    
    Set TestPathElementsShouldJoin = cc_isr_Test_Fx.Assert.IsTrue(PathExtensions.FolderExists(p_actualPath), _
                                            "The path " & p_actualPath & " should exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test creating the file.
    
    Set TestPathElementsShouldJoin = cc_isr_Test_Fx.Assert.IsTrue(PathExtensions.CreateTextFile(p_actualFilePath), _
                                        "The file " & p_actualFilePath & " should exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test checking if a file exists.
    
    Set TestPathElementsShouldJoin = cc_isr_Test_Fx.Assert.IsTrue(PathExtensions.FileExists(p_actualFilePath), _
                                        "The file " & p_actualFilePath & " should exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test deleting the file if it exists.
    
    Set TestPathElementsShouldJoin = cc_isr_Test_Fx.Assert.IsTrue(PathExtensions.DeleteFile(p_actualFilePath), _
                                        "The file " & p_actualFilePath & " should no longer exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test deleting the folder.
    
    Set TestPathElementsShouldJoin = cc_isr_Test_Fx.Assert.IsTrue(PathExtensions.DeleteFolder(p_actualPath), _
                                        "The path " & p_actualPath & " should no longer exist")
    
    If Not TestPathElementsShouldJoin.AssertSuccessful Then Exit Function
    
    ' test deleting the dummy folder.
    
    Set TestPathElementsShouldJoin = cc_isr_Test_Fx.Assert.IsTrue(PathExtensions.DeleteFolder(p_actualDummyPath), _
                                        "The path " & p_actualDummyPath & " should no longer exist")

End Function

