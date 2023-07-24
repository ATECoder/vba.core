Attribute VB_Name = "UserDefinedErrorsTests"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' UserDefinedErrorsTests.bas
'
' Dependencies:
'
' Assert.cls
' UserDefinedError.cls
' UserDefinedErrors.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_moduleName As String = "UserDefinedErrorsTests"

''' <summary>   Unit test. Asserts the existing of a user defined error. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of
''' <see cref="cc_isr_Test_Fx.Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestUserDefinedErrorShouldExist() As cc_isr_Test_Fx.Assert
    
    ' this should be added to the activate event of the workbook
    ' cc_isr_Core.UserDefinedErrors.Initialize
    Dim p_userError As UserDefinedError
    Set p_userError = cc_isr_Core.UserDefinedErrors.SocketConnectionError
    
    Set TestUserDefinedErrorShouldExist = cc_isr_Test_Fx.Assert.IsTrue(UserDefinedErrors.UserDefinedErrorExists(p_userError), _
                                                        p_userError.ToString(" should exist"))
End Function

''' <summary>   Unit test. Asserts buidling the error message. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestErrorMessageShouldBuild() As cc_isr_Test_Fx.Assert

    Const thisProcedureName = "TestErrorMessageShouldBuild"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' create an error
    Dim p_zero As Double: p_zero = 0
    Dim p_value As Double: p_value = 1 / p_zero
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' build the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource thisProcedureName, m_moduleName
    
    Set TestErrorMessageShouldBuild = cc_isr_Test_Fx.Assert.IsTrue(Len(Err.Source) > 0, _
            "VBA.Err.Source should not be empty")
    
    If TestErrorMessageShouldBuild.AssertSuccessful Then
    
        Dim p_expectedErrorSource As String
        p_expectedErrorSource = ThisWorkbook.VBProject.name & "." & m_moduleName & "." & thisProcedureName
        
        Set TestErrorMessageShouldBuild = cc_isr_Test_Fx.Assert.areEqual(p_expectedErrorSource, _
                VBA.Err.Source, "VBA.Err.Source should equal the expected value")
    
    End If
    
    If TestErrorMessageShouldBuild.AssertSuccessful Then
    
        Dim p_errorMessage As String: p_errorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
        
        Set TestErrorMessageShouldBuild = cc_isr_Test_Fx.Assert.IsTrue(Len(p_errorMessage) > 0, _
                "error message should build")
    
    End If
   
    
    On Error Resume Next
    
    GoTo exit_Handler

End Function



