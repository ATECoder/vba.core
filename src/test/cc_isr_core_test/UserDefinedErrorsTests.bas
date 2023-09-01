Attribute VB_Name = "UserDefinedErrorsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. User defined errors methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

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
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
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

''' <summary>   Unit test. Asserts building an error message using the last user defined error. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestLastErrorShouldBuild() As cc_isr_Test_Fx.Assert

    Const thisProcedureName = "TestLastErrorShouldBuild"
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' raise a user defined error
    cc_isr_Core_IO.GuardClauses.GuardNullReference Nothing, _
        ThisWorkbook.VBProject.Name & ".UserDefinedErrorsTests." & thisProcedureName

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    Set TestLastErrorShouldBuild = p_outcome
    Debug.Print p_outcome.BuildReport("TestLastErrorShouldBuild")
   
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
    
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(cc_isr_Core_IO.UserDefinedErrors.LastError, _
                "User defined errors should have a last error.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(cc_isr_Core_IO.UserDefinedErrors.LastErrorsStack, _
                "User defined errors should have a last errors stack.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(0, cc_isr_Core_IO.UserDefinedErrors.LastErrorsStack.Count, _
                "User defined errors stack should be empty before buidlding the standard error message.")
    End If
    
    Dim p_lastError As cc_isr_Core_IO.UserDefinedError
    Set p_lastError = cc_isr_Core_IO.Factory.NewUserDefinedError.Clone(cc_isr_Core_IO.UserDefinedErrors.LastError)
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_lastError, _
                "User defined errors should clone the last error.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Dim p_errorMessage As String: p_errorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(Len(p_errorMessage) > 0, _
                "error message should build.")
    
    End If
   
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1, cc_isr_Core_IO.UserDefinedErrors.LastErrorsStack.Count, _
                "User defined errors stack should have a single error after buidlding the standard error message.")
    End If
   
    Dim p_stackError As cc_isr_Core_IO.UserDefinedError
    Set p_stackError = cc_isr_Core_IO.UserDefinedErrors.LastErrorsStack.Pop
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNothing(p_stackError, _
                "Last stacked error should be set from the User defined errors stack.")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_lastError.Code, p_stackError.Code, _
                "User defined errors stack should have the same error code as the last error.")
    End If
   
    
    On Error Resume Next
    
    GoTo exit_Handler

End Function





