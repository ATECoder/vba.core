Attribute VB_Name = "StackTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Stack methods. </summary>
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
            Set p_outcome = TestStackShouldConstruct
        Case 2
            Set p_outcome = TestStackShouldPush
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
    This.Name = "BinaryExtensionTests"
    'BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 2
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


''' <summary>   [Unit Test]. Tests Constructing the Stack. </summary>
Public Function TestStackShouldConstruct() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_expectedCount As Integer:    p_expectedCount = 0
    Dim p_expectedCapacity As Integer: p_expectedCapacity = 11
    Dim p_Stack As Stack: Set p_Stack = cc_isr_Core_IO.Factory.CreateStack(p_expectedCapacity)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_Stack, _
                                    "A Stack should be created.")

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                                        p_Stack.Count, _
                                        "The Stack count should initialize at " & VBA.CStr(p_expectedCount) & ".")
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCapacity, _
                                        p_Stack.Capacity, _
                                        "The Stack should have the expected capacity.")
    End If

    Debug.Print p_outcome.BuildReport("TestStackShouldConstruct")
    
    Set TestStackShouldConstruct = p_outcome

End Function

''' <summary>   [Unit Test]. The Stack should Push, Pop and peek. </summary>
Public Function TestStackShouldPush() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_expectedCount As Integer:    p_expectedCount = 0
    Dim p_expectedCapacity As Integer: p_expectedCapacity = 11
    Dim p_Stack As Stack: Set p_Stack = cc_isr_Core_IO.Factory.CreateStack(p_expectedCapacity)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_Stack, _
                                    "A Stack should be created.")

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                                        p_Stack.Count, _
                                        "The Stack count should initialize at " & VBA.CStr(p_expectedCount) & ".")
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCapacity, _
                                        p_Stack.Capacity, _
                                        "The Stack should have the expected capacity.")
    End If

    Dim p_lastItem As Integer
    
    If p_outcome.AssertSuccessful Then
        p_lastItem = 1
        p_expectedCount = p_expectedCount + 1
        p_Stack.Push VBA.CStr(p_lastItem)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                                        p_Stack.Count, _
                                        "The Stack count should increment after Pushing.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim i As Integer
        For i = 1 To p_Stack.Capacity - 1
            p_lastItem = p_lastItem + 1
            p_Stack.Push VBA.CStr(p_lastItem)
            DoEvents
        Next i
        p_expectedCount = p_Stack.Capacity
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                            p_Stack.Count, _
                            "The Stack count should set to full count after adding " & VBA.CStr(p_Stack.Capacity) & " items.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_lastItem = p_lastItem + 1
        p_Stack.Push VBA.CStr(p_lastItem)
        p_expectedCount = p_Stack.Capacity
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                            p_Stack.Count, _
                            "The Stack count should remain at full capacity after adding beyond the Stack capacity.")
    End If
    
    Dim p_index As Integer
    Dim p_expectedValue As Integer
    Dim p_actualValue As String
    If p_outcome.AssertSuccessful Then
        p_index = 1
        p_expectedValue = IIf(p_lastItem > p_Stack.Capacity, p_lastItem - p_Stack.Capacity + 1, p_index)
        p_actualValue = p_Stack.peek(p_index)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.CStr(p_expectedValue), _
                            p_actualValue, _
                            "The Stack should have the expected value at the " & VBA.CStr(p_index) & " index.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_index = p_Stack.Capacity
        p_expectedValue = IIf(p_lastItem > p_Stack.Capacity, p_lastItem, p_index)
        p_actualValue = p_Stack.peek(p_index)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.CStr(p_expectedValue), _
                            p_actualValue, _
                            "The Stack should have the expected value at the " & VBA.CStr(p_index) & " index.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_index = p_Stack.Capacity
        p_expectedValue = IIf(p_lastItem > p_Stack.Capacity, p_lastItem, p_index)
        While p_Stack.Count > 0
            DoEvents
            p_actualValue = p_Stack.pop()
            If p_outcome.AssertSuccessful Then
                Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.CStr(p_expectedValue), _
                                    p_actualValue, _
                                    "The Stack should have the expected value at the " & VBA.CStr(p_index) & " index after poping to " & p_Stack.Count & " items.")
                p_expectedValue = p_expectedValue - 1
                p_index = p_index - 1
            End If
        Wend
        
    End If
    
    
    Debug.Print p_outcome.BuildReport("TestStackShouldPush")
    
    Set TestStackShouldPush = p_outcome

End Function





