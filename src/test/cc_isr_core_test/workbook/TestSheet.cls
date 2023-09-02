VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const m_lastErrorMessageCell As String = "C4"

''' <summary>   Gets the Last Error Message. </summary>
Friend Property Get LastErrorMessage() As String
    LastErrorMessage = Range(m_lastErrorMessageCell).value
End Property

''' <summary>   Sets the Last Error Message. </summary>
Friend Property Let LastErrorMessage(ByVal a_value As String)
    Range(m_lastErrorMessageCell).value = a_value
End Property

''' <summary>   Populates a combo box with the list of tests. </summary>
Private Sub ListTestsButton_Click()
    
    Dim p_module As ModuleInfo
    testNamesCombo.Clear
    For Each p_module In cc_isr_Test_Fx.WorkbookUtilities.EnumerateProjectModules(cc_isr_Core_Test.ThisWorkbook.VBProject)
        DoEvents
        testNamesCombo.AddItem p_module.FullName
    Next p_module

End Sub

''' <summary>   Runs the selected test. </summary>
Private Sub RunTestsButton_Click()

    cc_isr_Test_Fx.TestExecutive.Execute TestSheet, True

End Sub

Private Sub TestRaisingRemoteError()

    Const p_procedureName As String = "TestRaisingRemoteError"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_object As Object
    
    cc_isr_Core_IO.UserDefinedErrors.TestRaisingNullArgumentError p_object
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    MsgBox cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub
