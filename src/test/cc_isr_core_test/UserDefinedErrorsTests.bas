Attribute VB_Name = "UserDefinedErrorsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. User defined errors methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    TestNumber As Integer
    BeforeAllAssert As Assert
    BeforeEachAssert As Assert
End Type

Private This As this_

Public Sub BeforeAll()

    This.TestNumber = 0
    
    Set This.BeforeAllAssert = Assert.IsTrue(True, "initialize the overall assert.")
    
    If cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack Is Nothing Then
        Set This.BeforeAllAssert = Assert.Inconclusive("User defined errors should have an error archive.")
    End If
    
    If This.BeforeAllAssert.AssertSuccessful Then
    
        ' clear the error archive
        cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Clear
    
        If cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count <> 0 Then
            Set This.BeforeAllAssert = Assert.Inconclusive("User defined errors error archive should be empty.")
        End If
    
    End If
    
    If This.BeforeAllAssert.AssertSuccessful Then
    
        If cc_isr_Core_IO.UserDefinedErrors.ErrorsQueue Is Nothing Then
            Set This.BeforeAllAssert = Assert.Inconclusive("User defined errors should have an error queue")
        End If
        
    End If
    
    If This.BeforeAllAssert.AssertSuccessful Then
    
        ' clear the error queue
        cc_isr_Core_IO.UserDefinedErrors.ErrorsQueue.Clear
        
        If cc_isr_Core_IO.UserDefinedErrors.ErrorsQueue.Count <> 0 Then
            Set This.BeforeAllAssert = Assert.Inconclusive("User defined errors error queue should be empty.")
        End If
        
    End If
   
End Sub

Public Sub BeforeEach()

    If This.BeforeAllAssert.AssertSuccessful Then
    
        Set This.BeforeEachAssert = Assert.IsTrue(True, "initialize the pre-test assert.")
    
    Else
        
        Set This.BeforeEachAssert = Assert.Inconclusive(This.BeforeAllAssert.AssertMessage)
    
    End If
    
    This.TestNumber = This.TestNumber + 1
    
End Sub

Public Sub AfterEach()
    Set This.BeforeEachAssert = Nothing
End Sub

Public Sub AfterAll()
    
    Set This.BeforeAllAssert = Nothing

End Sub

''' <summary>   Unit test. Asserts the existing of a user defined error. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of
''' <see cref="cc_isr_Test_Fx.Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestUserDefinedErrorShouldExist() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    ' this should be added to the activate event of the workbook
    ' cc_isr_Core_IO.UserDefinedErrors.Initialize
    Dim p_userError As UserDefinedError
    Set p_userError = cc_isr_Core_IO.UserDefinedErrors.SocketConnectionError
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(UserDefinedErrors.UserDefinedErrorExists(p_userError), _
                                                        p_userError.ToString(" should exist"))
                                                        
    Debug.Print p_outcome.BuildReport("TestUserDefinedErrorShouldExist")
    
    Set TestUserDefinedErrorShouldExist = p_outcome

End Function

''' <summary>   Unit test. Asserts building the error message. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestErrorMessageShouldBuild() As cc_isr_Test_Fx.Assert

    Const thisProcedureName = "TestErrorMessageShouldBuild"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If Not p_outcome.AssertSuccessful Then GoTo exit_Handler
    
    ' create an error
    Dim p_zero As Double: p_zero = 0
    Dim p_value As Double: p_value = 1 / p_zero
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Set TestErrorMessageShouldBuild = p_outcome
    Debug.Print p_outcome.BuildReport("TestErrorMessageShouldBuild")
   
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' build the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource thisProcedureName, "UserDefinedErrorsTests", ThisWorkbook
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(Len(Err.Source) > 0, _
            "VBA.Err.Source should not be empty.")
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_expectedErrorSource As String
        p_expectedErrorSource = ThisWorkbook.VBProject.Name & ".UserDefinedErrorsTests." & thisProcedureName
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedErrorSource, _
                VBA.Err.Source, "VBA.Err.Source should equal the expected value")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_errorMessage As String: p_errorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(Len(p_errorMessage) > 0, _
                "error message should build.")
    
    End If
   
    
    On Error Resume Next
    
    GoTo exit_Handler

End Function

''' <summary>   Unit test. Asserts building an error message using the raised user defined error. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestRaisedErrorShouldBeReported() As cc_isr_Test_Fx.Assert

    Const thisProcedureName = "TestRaisedErrorShouldBeReported"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
    If Not p_outcome.AssertSuccessful Then GoTo exit_Handler
    
    Dim p_expectedArchivedErrorsCount As Integer
    p_expectedArchivedErrorsCount = 0
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedArchivedErrorsCount, _
            cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count, _
            "User defined errors error archive should be empty before buidlding the first standard error message.")
    End If
    
    Dim p_expectedQueuedErrorsCount As Integer
    p_expectedQueuedErrorsCount = 0
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedQueuedErrorsCount, _
            cc_isr_Core_IO.UserDefinedErrors.ErrorsQueue.Count, _
            "User defined errors error queue should be empty before enqueueing the first error.")
    End If
    
    If Not p_outcome.AssertSuccessful Then GoTo exit_Handler
    
    ' save the current error counts
    Dim p_queuedErrorsCount As Integer
    p_queuedErrorsCount = cc_isr_Core_IO.UserDefinedErrors.ErrorsQueue.Count
    p_expectedArchivedErrorsCount = p_queuedErrorsCount
    
    Dim p_archivedErrorsCount As Integer
    p_archivedErrorsCount = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count
    p_expectedQueuedErrorsCount = p_archivedErrorsCount
    
    ' raise a user defined error
    cc_isr_Core_IO.GuardClauses.GuardNullReference Nothing, _
        ThisWorkbook.VBProject.Name & ".UserDefinedErrorsTests." & thisProcedureName

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Set TestRaisedErrorShouldBeReported = p_outcome
    Debug.Print p_outcome.BuildReport("TestRaisedErrorShouldBeReported")
   
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' build the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource thisProcedureName, "UserDefinedErrorsTests", ThisWorkbook
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(Len(Err.Source) > 0, _
            "VBA.Err.Source should not be empty.")
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_expectedErrorSource As String
        p_expectedErrorSource = ThisWorkbook.VBProject.Name & ".UserDefinedErrorsTests." & thisProcedureName
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedErrorSource, _
                VBA.Err.Source, "VBA.Err.Source should equal the expected value")
    
    End If
    
    p_expectedQueuedErrorsCount = p_expectedQueuedErrorsCount + 1
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedQueuedErrorsCount, _
            cc_isr_Core_IO.UserDefinedErrors.ErrorsQueue.Count, _
            "User defined errors error queue should increment by one after raising an error.")
    End If
    
    
    Dim p_lastError As cc_isr_Core_IO.UserDefinedError
    
    If p_outcome.AssertSuccessful Then
    
        Set p_lastError = cc_isr_Core_IO.Factory.NewUserDefinedError.Clone(cc_isr_Core_IO.UserDefinedErrors.ErrorsQueue.Peek())
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_lastError, _
                "User defined errors should clone the last error.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_errorMessage As String: p_errorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(Len(p_errorMessage) > 0, _
                "error message should build.")
    
    End If
   
    p_expectedArchivedErrorsCount = p_expectedArchivedErrorsCount + 1
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedArchivedErrorsCount, _
            cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count, _
            "User defined errors error archive stack should have a single error after buidlding the standard error message.")
    End If
   
    p_expectedQueuedErrorsCount = p_expectedQueuedErrorsCount - 1
   
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedQueuedErrorsCount, _
            cc_isr_Core_IO.UserDefinedErrors.ErrorsQueue.Count, _
            "User defined errors error queue should be empty after buidlding the standard error message.")
    End If
   
    Dim p_stackError As cc_isr_Core_IO.UserDefinedError
    
    If p_outcome.AssertSuccessful Then
    
    Set p_stackError = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_stackError, _
                "Last archived error should pop from the User defined errors error archive.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_lastError.Code, p_stackError.Code, _
                "User defined errors stack should have the same error code as the last error.")
    End If
   
    
    On Error Resume Next
    
    GoTo exit_Handler

End Function

Public Sub RunTests()
    BeforeAll
    BeforeEach
    TestRaisedErrorShouldBeReported
    AfterEach
    AfterAll
End Sub





