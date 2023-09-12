Attribute VB_Name = "PathExtensionsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Path extension methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Name As String
    TestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    ErrTracer As IErrTracer
    TestCount As Integer
    RunCount As Integer
    PassedCount As Integer
    FailedCount As Integer
    InconclusiveCount As Integer
End Type

Private This As this_

''' <summary>   Runs the specified test. </summary>
Public Function RunTest(ByVal a_testNumber As Integer) As cc_isr_Test_Fx.Assert
    Dim p_outcome As cc_isr_Test_Fx.Assert
    'BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestPathElementsShouldJoin
        Case Else
    End Select
    Set RunTest = p_outcome
    'AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    'BeforeAll
    RunTest 1
    'AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    This.Name = "PathExtensionTests"
    'BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 1
    Dim p_testNumber As Integer
    For p_testNumber = 1 To This.TestCount
        Set p_outcome = RunTest(p_testNumber)
        If Not p_outcome Is Nothing Then
            This.RunCount = This.RunCount + 1
            If p_outcome.AssertInconclusive Then
                This.InconclusiveCount = This.InconclusiveCount + 1
            ElseIf p_outcome.AssertSuccessful Then
                This.PassedCount = This.PassedCount + 1
            Else
                This.FailedCount = This.FailedCount + 1
            End If
        End If
        DoEvents
    Next p_testNumber
    'AfterAll
    Debug.Print "Ran " & VBA.CStr(This.RunCount) & " out of " & VBA.CStr(This.TestCount) & " tests."
    Debug.Print "Passed: " & VBA.CStr(This.PassedCount) & "; Failed: " & VBA.CStr(This.FailedCount) & _
                "; Inconclusive: " & VBA.CStr(This.InconclusiveCount) & "."
End Sub


''' <summary>   Unit test. Asserts that the path elements should join and create the directory. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestPathElementsShouldJoin() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_element1 As String: p_element1 = Excel.ActiveWorkbook.path
    Dim p_element2 As String: p_element2 = "dummy"
    Dim p_element3 As String: p_element3 = "workbook"
    Dim p_fileName As String: p_fileName = "filename.txt"
    
    ' test joining without creating
    
    Dim p_expectedDummyPath As String: p_expectedDummyPath = p_element1 & "\" & p_element2
    Dim p_actualDummyPath As String: p_actualDummyPath = PathExtensions.Join(p_element1, p_element2)
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedDummyPath, _
                                        p_actualDummyPath, _
                                        "The path elements should be joined")
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_expectedPath As String: p_expectedPath = p_element1 & "\" & p_element2 & "\" & p_element3
        Dim p_expectedFilePath As String: p_expectedFilePath = p_expectedPath & "\" & p_fileName
        
        ' test joining without creating
        
        Dim p_actualPath As String: p_actualPath = PathExtensions.JoinAll(False, p_element1, p_element2, p_element3)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedPath, _
                                                       p_actualPath, _
                                                       "The path elements should be joined")
     
    End If
    
    If p_outcome.AssertSuccessful Then
    
        ' test joining a file.
        
        Dim p_actualFilePath As String: p_actualFilePath = PathExtensions.JoinFile(p_actualPath, p_fileName)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFilePath, _
            p_actualFilePath, "The path path should be joined")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        ' test deleting the folder if it exists.
        
        Set p_outcome = Assert.IsTrue(PathExtensions.DeleteFolder(p_actualPath), _
            "The path " & p_actualPath & " should no longer exist")
        
    End If
    
    If p_outcome.AssertSuccessful Then
        
        ' test joining and creating.
        
        p_actualPath = PathExtensions.JoinAll(True, p_element1, p_element2, p_element3)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedPath, _
            p_actualPath, "The path element should be joined")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        ' test detecting the created folder.
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(PathExtensions.FolderExists(p_actualPath), _
            "The path " & p_actualPath & " should exist")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        ' test creating the file.
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(PathExtensions.CreateTextFile(p_actualFilePath), _
            "The file " & p_actualFilePath & " should exist")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        ' test checking if a file exists.
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(PathExtensions.FileExists(p_actualFilePath), _
            "The file " & p_actualFilePath & " should exist")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        ' test deleting the file if it exists.
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(PathExtensions.DeleteFile(p_actualFilePath), _
            "The file " & p_actualFilePath & " should no longer exist")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        ' test deleting the folder.
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(PathExtensions.DeleteFolder(p_actualPath), _
                "The path " & p_actualPath & " should no longer exist")
    End If
    
    
    If p_outcome.AssertSuccessful Then
        
        ' test deleting the dummy folder.
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(PathExtensions.DeleteFolder(p_actualDummyPath), _
                "The path " & p_actualDummyPath & " should no longer exist")
    End If
    

    Debug.Print p_outcome.BuildReport("TestPathElementsShouldJoin")
    
    Set TestPathElementsShouldJoin = p_outcome

End Function

