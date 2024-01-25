Attribute VB_Name = "QueueTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Queue methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Name As String
    TestNumber As Integer
    PreviousTestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    TestStopper As cc_isr_Core_IO.Stopwatch
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
    This.TestNumber = a_testNumber
    BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestQueueShouldConstruct
        Case 2
            Set p_outcome = TestQueueShouldEnqueue
        Case Else
    End Select
    Set RunTest = p_outcome
    AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 1
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    This.Name = "BinaryExtensionTests"
    BeforeAll
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
    AfterAll
    Debug.Print "Ran " & VBA.CStr(This.RunCount) & " out of " & VBA.CStr(This.TestCount) & " tests."
    Debug.Print "Passed: " & VBA.CStr(This.PassedCount) & "; Failed: " & VBA.CStr(This.FailedCount) & _
                "; Inconclusive: " & VBA.CStr(This.InconclusiveCount) & "."
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Tests initialize and cleanup.
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Prepares all tests. </summary>
''' <remarks>   This method sets up the 'Before All' <see cref="cc_isr_Test_Fx.Assert"/>
''' which serves to set the 'Before Each' <see cref="cc_isr_Test_Fx.Assert"/>.
''' The error object and user defined errors state are left clear after this method. </remarks>
Public Sub BeforeAll()

    Const p_procedureName As String = "BeforeAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = Assert.Pass("Primed to run all tests.")

    Set This.TestStopper = cc_isr_Core_IO.Factory.NewStopwatch
    Set This.ErrTracer = New ErrTracer
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState

    ' Prime all tests
    This.TestNumber = 0
    This.PreviousTestNumber = 0
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
            Set p_outcome = Assert.Pass("Primed to run all tests.")
        Else
            Set p_outcome = Assert.Inconclusive("Failed priming all tests;" & _
                VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    End If
    
    Set This.BeforeAllAssert = p_outcome
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Prepares each test before it is run. </summary>
''' <remarks>   This method sets up the 'Before Each' <see cref="cc_isr_Test_Fx.Assert"/>
''' which serves to initialize the <see cref="cc_isr_Test_Fx.Assert"/> of each test.
''' The error object and user defined errors state are left clear after this method. </remarks>
Public Sub BeforeEach()

    Const p_procedureName As String = "BeforeEach"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert

    If This.BeforeAllAssert.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber) & ".")
    Else
        Set p_outcome = Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeAllAssert.AssertMessage)
    End If
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
   
    ' Prepare the next test
    If This.TestNumber = This.PreviousTestNumber Then _
        This.TestNumber = This.PreviousTestNumber + 1
   
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
             Set p_outcome = Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber))
        Else
            Set p_outcome = Assert.Inconclusive("Failed priming pre-test #" & VBA.CStr(This.TestNumber) & _
                ";" & VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    End If
    
    Set This.BeforeEachAssert = p_outcome

    On Error GoTo 0
    
    This.TestStopper.Restart
    
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
                       
End Sub

''' <summary>   Releases test elements after each tests is run. </summary>
''' <remarks>   This method uses the <see cref="ErrTracer"/> to report any leftover errors
''' in the user defined errors queue and stack. The error object and user defined errors
''' state are left clear after this method. </remarks>
Public Sub AfterEach()
    
    Const p_procedureName As String = "AfterEach"
    
    ' Trap errors to the error handler.
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")

    ' check if we can proceed with cleanup.
    
    If Not This.BeforeEachAssert.AssertSuccessful Then _
        Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Unable to cleanup test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeEachAssert.AssertMessage)

    ' cleanup after each test.
    This.PreviousTestNumber = This.TestNumber
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before Each' cc_isr_Test_Fx.Assert.
    Set This.BeforeEachAssert = Nothing

    If p_outcome.AssertSuccessful Then
    
        ' report any leftover errors.
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        If p_outcome.AssertSuccessful Then
            Set p_outcome = cc_isr_Test_Fx.Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
        Else
            Set p_outcome = cc_isr_Test_Fx.Assert.Inconclusive("Errors reported cleaning up test #" & VBA.CStr(This.TestNumber) & _
                ";" & VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    
    End If

    If Not p_outcome.AssertSuccessful Then _
        This.ErrTracer.TraceError p_outcome.AssertMessage
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Releases the test class after all tests run. </summary>
''' <remarks>   This method uses the <see cref="ErrTracer"/> to report any leftover errors
''' in the user defined errors queue and stack. The error object and user defined errors
''' state are left clear after this method. </remarks>
Public Sub AfterAll()
    
    Const p_procedureName As String = "AfterAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = Assert.Pass("All tests cleaned up.")
    
    ' cleanup after all tests.
    If This.BeforeAllAssert.AssertSuccessful Then
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before All' assert.
    Set This.BeforeAllAssert = Nothing

    ' report any leftover errors.
    Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
    Else
        Set p_outcome = Assert.Inconclusive("Errors reported cleaning up all tests;" & _
            VBA.vbCrLf & p_outcome.AssertMessage)
    End If
    
    If Not p_outcome.AssertSuccessful Then _
        This.ErrTracer.TraceError p_outcome.AssertMessage
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Tests
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   [Unit Test]. Tests Constructing the Queue. </summary>
Public Function TestQueueShouldConstruct() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "TestQueueShouldConstruct"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_expectedCount As Integer:    p_expectedCount = 0
    Dim p_expectedCapacity As Integer: p_expectedCapacity = 11
    Dim p_queue As Queue: Set p_queue = cc_isr_Core_IO.Factory.CreateQueue(p_expectedCapacity)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_queue, _
                                    "A queue should be created.")

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                                        p_queue.Count, _
                                        "The Queue count should initialize at " & VBA.CStr(p_expectedCount) & ".")
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCapacity, _
                                        p_queue.Capacity, _
                                        "The Queue should have the expected capacity.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestQueueShouldConstruct = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    

End Function

''' <summary>   [Unit Test]. The Queue should enqueue, dequeue and peek. </summary>
Public Function TestQueueShouldEnqueue() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "TestQueueShouldEnqueue"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_expectedCount As Integer:    p_expectedCount = 0
    Dim p_expectedCapacity As Integer: p_expectedCapacity = 11
    Dim p_queue As Queue: Set p_queue = cc_isr_Core_IO.Factory.CreateQueue(p_expectedCapacity)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_queue, _
                                    "A queue should be created.")

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                                        p_queue.Count, _
                                        "The Queue count should initialize at " & VBA.CStr(p_expectedCount) & ".")
    End If

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCapacity, _
                                        p_queue.Capacity, _
                                        "The Queue should have the expected capacity.")
    End If

    Dim p_lastItem As Integer
    
    If p_outcome.AssertSuccessful Then
        p_lastItem = 1
        p_expectedCount = p_expectedCount + 1
        p_queue.Enqueue VBA.CStr(p_lastItem)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                                        p_queue.Count, _
                                        "The Queue count should increment after enqueueing.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim i As Integer
        For i = 1 To p_queue.Capacity - 1
            p_lastItem = p_lastItem + 1
            p_queue.Enqueue VBA.CStr(p_lastItem)
            DoEvents
        Next i
        p_expectedCount = p_queue.Capacity
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                            p_queue.Count, _
                            "The Queue count should set to full count after adding " & VBA.CStr(p_queue.Capacity) & " items.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_lastItem = p_lastItem + 1
        p_queue.Enqueue VBA.CStr(p_lastItem)
        p_expectedCount = p_queue.Capacity
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedCount, _
                            p_queue.Count, _
                            "The Queue count should remain at full capacity after adding beyond the queue capacity.")
    End If
    
    Dim p_index As Integer
    Dim p_expectedValue As Integer
    Dim p_actualValue As String
    If p_outcome.AssertSuccessful Then
        p_index = 1
        p_expectedValue = IIf(p_lastItem > p_queue.Capacity, p_lastItem - p_queue.Capacity + 1, p_index)
        p_actualValue = p_queue.peek(p_index)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.CStr(p_expectedValue), _
                            p_actualValue, _
                            "The Queue should have the expected value at the " & VBA.CStr(p_index) & " index.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_index = p_queue.Capacity
        p_expectedValue = IIf(p_lastItem > p_queue.Capacity, p_lastItem, p_index)
        p_actualValue = p_queue.peek(p_index)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.CStr(p_expectedValue), _
                            p_actualValue, _
                            "The Queue should have the expected value at the " & VBA.CStr(p_index) & " index.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_index = 1
        p_expectedValue = IIf(p_lastItem > p_queue.Capacity, p_lastItem - p_queue.Capacity + 1, p_index)
        While p_queue.Count > 0
            DoEvents
            p_actualValue = p_queue.Dequeue()
            If p_outcome.AssertSuccessful Then
                Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.CStr(p_expectedValue), _
                                    p_actualValue, _
                                    "The Queue should have the expected value at the " & VBA.CStr(p_index) & " index after dequeuing to " & p_queue.Count & " items.")
                p_expectedValue = p_expectedValue + 1
                p_index = p_index + 1
            End If
        Wend
        
    End If
    
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestQueueShouldEnqueue = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    

End Function



