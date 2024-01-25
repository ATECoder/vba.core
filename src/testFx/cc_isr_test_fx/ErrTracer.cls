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

''' <summary>   Asserts and traces leftover errors from the error queue and error stack. </summary>
''' <para name="a_delimiter">   [String] a delimiter for concatenating error messages
'''                             from the error queue. </param>
Private Function IErrTracer_AssertLeftoverErrors(Optional ByVal a_delimiter As String = VBA.vbCrLf) As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_archivedErrorCount As Integer
    p_archivedErrorCount = cc_isr_Core_IO.UserDefinedErrors.ArchivedErrorCount
    Dim p_queuedErrorCount As Integer
    p_queuedErrorCount = cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount
    If p_archivedErrorCount > 0 Or p_queuedErrorCount > 0 Then
        
        ' if we have leftover errors, move the queued errors to the stack and
        ' report the stacked errors.
        Dim p_queuedErrors As String
        p_queuedErrors = cc_isr_Core_IO.UserDefinedErrors.ArchiveQueuedErrors()
        
        Dim p_archivedErrors As String
        p_archivedErrors = cc_isr_Core_IO.UserDefinedErrors.PopArchivedErrors()
        Set p_outcome = Assert.Fail( _
            VBA.CStr(p_queuedErrorCount) & " errors were found in the error queue and " & _
            VBA.CStr(p_archivedErrorCount) & " errors were found in the error stack:" & _
            VBA.vbCrLf & p_archivedErrors)
    Else
        
        Set p_outcome = Assert.Pass("There were no leftover errors to assert.")
        
    End If
    
    Set IErrTracer_AssertLeftoverErrors = p_outcome

End Function

''' <summary>   Traces an error message.  </summary>
''' <para name="a_message">   [String] the message to trace. </param>
Private Sub IErrTracer_TraceError(ByVal a_message As String)
    UnitTestSheet.LastErrorMessage = a_message
End Sub


