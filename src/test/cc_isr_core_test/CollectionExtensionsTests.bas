Attribute VB_Name = "CollectionExtensionsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Collection extension methods. </summary>
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
            Set p_outcome = TestCollectionShouldContain
        Case 2
            Set p_outcome = TestCollectionShouldNotContain
        Case 3
            Set p_outcome = TestCollectionShouldContainItself
        Case 4
            Set p_outcome = TestCollectionShouldBeEqual
        Case 5
            Set p_outcome = TestCollectionShouldNotBeEqual
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
    This.Name = "CollectionExtensionTests"
    'BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 5
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

''' <summary>   Unit test. Asserts that the collection contains an expected value. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldContain() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_col As VBA.Collection
    Set p_col = New VBA.Collection
    Dim p_expected As Variant: p_expected = "a"
    p_col.Add p_expected
    Set p_outcome = Assert.IsTrue(CollectionExtensions.ContainsKey(p_col, p_expected), "The collection should contain the value")

    Debug.Print p_outcome.BuildReport("TestCollectionShouldContain")
    
    Set TestCollectionShouldContain = p_outcome

End Function

''' <summary>   Unit test. Asserts that the collection does not contain a value. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldNotContain() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_col As VBA.Collection
    Set p_col = New VBA.Collection
    Dim p_expected As Variant: p_expected = "a"
    Dim notExpected As Variant: notExpected = "b"
    p_col.Add p_expected
    Set p_outcome = Assert.IsFalse(CollectionExtensions.ContainsKey(p_col, notExpected), "The collection should not contain a value")

    Debug.Print p_outcome.BuildReport("TestCollectionShouldNotContain")
    
    Set TestCollectionShouldNotContain = p_outcome

End Function

''' <summary>   Unit test. Asserts that the collection contains an expected value. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldContainItself() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_col As New VBA.Collection
    p_col.Add "a"
    p_col.Add "b"
    Set p_outcome = Assert.IsTrue(CollectionExtensions.ContainsAll(p_col, p_col), _
                                    "The collection should contain itself")

    Debug.Print p_outcome.BuildReport("TestCollectionShouldContainItself ")
    
    Set TestCollectionShouldContainItself = p_outcome

End Function

''' <summary>   Unit test. Asserts that collection are equal. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldBeEqual() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_left As VBA.Collection
    Set p_left = New VBA.Collection
    Dim p_right As VBA.Collection
    Set p_right = New VBA.Collection
    Dim p_item As String
    p_item = "a": p_left.Add p_item: p_right.Add p_item
    p_item = "b": p_left.Add p_item: p_right.Add p_item
    
    Set p_outcome = Assert.IsTrue(CollectionExtensions.areEqual(p_left, p_right), _
        "The collection should be equal")

    Debug.Print p_outcome.BuildReport("TestCollectionShouldBeEqual")
    
    Set TestCollectionShouldBeEqual = p_outcome

End Function

''' <summary>   Unit test. Asserts that collection are not equal. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldNotBeEqual() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_left As VBA.Collection
    Set p_left = New VBA.Collection
    Dim p_right As VBA.Collection
    Set p_right = New VBA.Collection
    Dim p_item As String
    p_item = "a": p_left.Add p_item: p_right.Add p_item
    p_item = "b": p_left.Add p_item: p_right.Add p_item
    p_item = "c": p_left.Add p_item
    
    Set p_outcome = Assert.IsFalse(CollectionExtensions.areEqual(p_left, p_right), _
        "The collection should not be equal because they are of difference length")

    If p_outcome.AssertSuccessful Then
        p_right.Add p_item & "d"
        Set p_outcome = Assert.IsFalse(CollectionExtensions.areEqual(p_left, p_right), _
            "The collection should not be equal because they have difference items")
    End If
    Debug.Print p_outcome.BuildReport("TestCollectionShouldNotBeEqual")
    
    Set TestCollectionShouldNotBeEqual = p_outcome

End Function



