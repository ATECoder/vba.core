Attribute VB_Name = "UserDefinedErrorsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. User defined errors methods. </summary>
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
            Set p_outcome = TestErrorMessageShouldBuild
        Case 2
            Set p_outcome = TestRaisedErrorShouldBeReported
        Case 3
            Set p_outcome = TestUserDefinedErrorShouldExist
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
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 3
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

''' <summary>   Unit test. Asserts the existing of a user defined error. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of
''' <see cref="cc_isr_Test_Fx.Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestUserDefinedErrorShouldExist() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "TestUserDefinedErrorShouldExist"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    ' this should be added to the activate event of the workbook
    ' cc_isr_Core_IO.UserDefinedErrors.Initialize
    Dim p_userError As UserDefinedError
    Set p_userError = cc_isr_Core_IO.UserDefinedErrors.SocketConnectionError
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(UserDefinedErrors.UserDefinedErrorExists(p_userError), _
                                                        p_userError.ToString(" should exist"))
                                                        
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestUserDefinedErrorShouldExist = p_outcome

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

''' <summary>   Unit test. Asserts building the error message. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestErrorMessageShouldBuild() As cc_isr_Test_Fx.Assert

    Const p_procedureName = "TestErrorMessageShouldBuild"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_errorNumber As Long
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If Not p_outcome.AssertSuccessful Then GoTo exit_Handler
    
    ' create an error
    Dim p_zero As Double: p_zero = 0
    Dim p_value As Double: p_value = 1 / p_zero
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestErrorMessageShouldBuild = p_outcome
   
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    p_errorNumber = VBA.Err.Number
    
    ' build the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, "UserDefinedErrorsTests", ThisWorkbook
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(Len(Err.Source) > 0, _
            "VBA.Err.Source should not be empty.")
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_expectedErrorSource As String
        p_expectedErrorSource = ThisWorkbook.VBProject.Name & ".UserDefinedErrorsTests." & p_procedureName
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedErrorSource, _
                VBA.Err.Source, "VBA.Err.Source should equal the expected value")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_errorMessage As String
        p_errorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(Len(p_errorMessage) > 0, _
                "error message should build.")
    
    End If
   
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1, _
            cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount, _
            "VBA Error should be added to the error archive.")
    
    End If
    
    Dim p_error As cc_isr_Core_IO.UserDefinedError
    
    If p_outcome.AssertSuccessful Then
    
        Set p_error = cc_isr_Core_IO.UserDefinedErrors.PeekArchive
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_errorNumber, p_error.Number, _
            "VBA Error should be the same as the error from the top of the stack.")
    
    End If
    
    On Error Resume Next
    
    GoTo exit_Handler

End Function

''' <summary>   Unit test. Asserts building an error message using the raised user defined error. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestRaisedErrorShouldBeReported() As cc_isr_Test_Fx.Assert

    Const p_procedureName = "TestRaisedErrorShouldBeReported"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = This.BeforeEachAssert
    
    If Not p_outcome.AssertSuccessful Then GoTo exit_Handler
    
    Dim p_expectedArchivedErrorsCount As Integer
    p_expectedArchivedErrorsCount = 0
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedArchivedErrorsCount, _
            cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount, _
            "User defined errors error archive should be empty before buidlding the first standard error message.")
    End If
    
    Dim p_expectedQueuedErrorsCount As Integer
    p_expectedQueuedErrorsCount = 0
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedQueuedErrorsCount, _
            cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount, _
            "User defined errors error queue should be empty before enqueueing the first error.")
    End If
    
    If Not p_outcome.AssertSuccessful Then GoTo exit_Handler
    
    ' save the current error counts
    Dim p_queuedErrorsCount As Integer
    p_queuedErrorsCount = cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount
    p_expectedArchivedErrorsCount = p_queuedErrorsCount
    
    Dim p_archivedErrorsCount As Integer
    p_archivedErrorsCount = cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount
    p_expectedQueuedErrorsCount = p_archivedErrorsCount
    
    ' raise a user defined error
    cc_isr_Core_IO.GuardClauses.GuardNullReference Nothing, _
        ThisWorkbook.VBProject.Name & ".UserDefinedErrorsTests." & p_procedureName

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestRaisedErrorShouldBeReported = p_outcome
   
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' build the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, "UserDefinedErrorsTests", ThisWorkbook
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(Len(Err.Source) > 0, _
            "VBA.Err.Source should not be empty.")
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_expectedErrorSource As String
        p_expectedErrorSource = ThisWorkbook.VBProject.Name & ".UserDefinedErrorsTests." & p_procedureName
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedErrorSource, _
                VBA.Err.Source, "VBA.Err.Source should equal the expected value")
    
    End If
    
    p_expectedQueuedErrorsCount = p_expectedQueuedErrorsCount + 1
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedQueuedErrorsCount, _
            cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount, _
            "User defined errors error queue should increment by one after raising an error.")
    End If
    
    
    Dim p_lastError As cc_isr_Core_IO.UserDefinedError
    
    If p_outcome.AssertSuccessful Then
    
        Set p_lastError = cc_isr_Core_IO.Factory.NewUserDefinedError.FromUserDefinedError(cc_isr_Core_IO.UserDefinedErrors.PeekQueue())
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_lastError, _
                "User defined errors should initialize from the last queued error.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_errorMessage As String: p_errorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(Len(p_errorMessage) > 0, _
                "error message should build.")
    
    End If
   
    p_expectedArchivedErrorsCount = p_expectedArchivedErrorsCount + 1
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedArchivedErrorsCount, _
            cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount, _
            "User defined errors error archive stack should have a single error after buidlding the standard error message.")
    End If
   
    p_expectedQueuedErrorsCount = p_expectedQueuedErrorsCount - 1
   
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedQueuedErrorsCount, _
            cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount, _
            "User defined errors error queue should be empty after buidlding the standard error message.")
    End If
   
    Dim p_stackError As cc_isr_Core_IO.UserDefinedError
    
    If p_outcome.AssertSuccessful Then
    
    Set p_stackError = cc_isr_Core_IO.UserDefinedErrors.PeekArchive
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_stackError, _
                "Last archived error should get peeked from the User defined errors error archive.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_lastError.Number, p_stackError.Number, _
                "User defined errors stack should have the same error Number as the last error.")
    End If
   
    
    On Error Resume Next
    
    GoTo exit_Handler

End Function

