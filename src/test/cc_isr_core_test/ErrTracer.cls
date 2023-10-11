VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrTracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An object that can trace error messages. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Implements cc_isr_Test_Fx.IErrTracer

''' <summary>   Asserts and traces leftover errors from the error queue. </summary>
''' <para name="a_delimiter">   [String] a delimiter for concatenating error messages
'''                             from the error queue. </param>
Private Function IErrTracer_AssertLeftoverErrors(Optional ByVal a_delimiter As String = VBA.vbCrLf) As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = Assert.AreEqual(0, _
        cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount, _
        "Error stack should be empty upon completion this test.")

    If p_outcome.AssertSuccessful Then
        Dim p_leftoverErrorMessage As String
        If cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount > 0 Then
            p_leftoverErrorMessage = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(a_delimiter)
            IErrTracer_TraceError p_leftoverErrorMessage
            Set p_outcome = Assert.fail("Found errors in the error queue: " & VBA.vbCrLf & _
                p_leftoverErrorMessage)
        End If
    End If
    
    Set IErrTracer_AssertLeftoverErrors = p_outcome

End Function

''' <summary>   Traces an error message.  </summary>
''' <para name="a_message">   [String] the message to trace. </param>
Private Sub IErrTracer_TraceError(ByVal a_message As String)
    TestSheet.LastErrorMessage = a_message
End Sub

