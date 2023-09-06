''' <summary>   This class properties. </summary>
Private Type this_
	Name as string
    TestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    ErrTracer As IErrTracer
End Type

Private This As this_

''' <summary>   Runs the specified test. </summary>
Public Sub RunTest(ByVal a_testNumber As Integer)
    BeforeEach
    Select Case a_testNumber
        Case 1
            TestInitializeAndDispose
        Case 2
            TestGettingLastErrorDescription
        Case Else
    End Select
    AfterEach
End Sub

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 1
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    BeforeAll
    Dim p_testNumber As Integer
    For p_testNumber = 1 To 2
        RunTest p_testNumber
        DoEvents
    Next p_testNumber
    AfterAll
End Sub

''' <summary>   Prepares all tests. </summary>
Public Sub BeforeAll()

    Const p_procedureName As String = "BeforeAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

	this.Name = "UnitTestSkeleton"
	
    This.TestNumber = 0
    
    Set This.ErrTracer = New ErrTracer
    
    Set This.BeforeAllAssert = Assert.Pass("initialize the overall assert.")
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count > 0 Then
        
        Dim p_leftoverErrorMessage As String
        p_leftoverErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop().ToString()
        Set This.BeforeAllAssert = Assert.Inconclusive("Failed preparing all tests: " & _
            p_leftoverErrorMessage)
        This.ErrTracer.TraceError p_leftoverErrorMessage
    
    End If

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error if not user defined error
    If Not cc_isr_Core_IO.UserDefinedErrors.IsUserDefinedError(VBA.Err.Number) Then _
        cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
	
End Sub

''' <summary>   Prepares each test before it is run. </summary>
Public Sub BeforeEach()

    Const p_procedureName As String = "BeforeEach"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    If This.BeforeAllAssert.AssertSuccessful Or This.TestNumber > 0 Then
        
        Set This.BeforeEachAssert = Assert.Pass("initialize the pre-test assert.")
    
    Else
    
        Set This.BeforeEachAssert = Assert.Inconclusive(This.BeforeAllAssert.AssertMessage)
    
    End If

    This.TestNumber = This.TestNumber + 1

    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    
    If This.BeforeEachAssert.AssertSuccessful Then
    
        Set This.BeforeEachAssert = Assert.AreEqual(0, Err.Number, _
            "Error Number should be 0.")
            
    End If
    
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count > 0 Then
        
        Dim p_leftoverErrorMessage As String
        p_leftoverErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop().ToString()
        Set This.BeforeAllAssert = Assert.Inconclusive("Failed preparing test #" & vba.cstr(this.TestNumber) & ": " & _
            p_leftoverErrorMessage)
        This.ErrTracer.TraceError p_leftoverErrorMessage
    
    End If

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error if not user defined error
    If Not cc_isr_Core_IO.UserDefinedErrors.IsUserDefinedError(VBA.Err.Number) Then _
        cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Releases test elements after each tests is run. </summary>
Public Sub AfterEach()
    
    Const p_procedureName As String = "AfterEach"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Set This.BeforeEachAssert = Nothing
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count > 0 Then
        
        Dim p_leftoverErrorMessage As String
        p_leftoverErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop().ToString()
        This.ErrTracer.TraceError "Error(s) were stacked unwinding test #" & vba.cstr(this.TestNumber) & ": " & p_leftoverErrorMessage
    
    End If

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error if not user defined error
    If Not cc_isr_Core_IO.UserDefinedErrors.IsUserDefinedError(VBA.Err.Number) Then _
        cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Releases the test class after all tests run. </summary>
Public Sub AfterAll()
    
    Const p_procedureName As String = "AfterAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Set This.BeforeAllAssert = Nothing
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Count > 0 Then
        
        Dim p_leftoverErrorMessage As String
        p_leftoverErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ErrorsArchiveStack.Pop().ToString()
        This.ErrTracer.TraceError "Errors were stacked unwinding all tests: " & p_leftoverErrorMessage
    
    End If

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error if not user defined error
    If Not cc_isr_Core_IO.UserDefinedErrors.IsUserDefinedError(VBA.Err.Number) Then _
        cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub


''' <summary>   Unit test. Asserts instantiating and disposing of the Winsock framework. </summary>
''' <returns>   [<see cref="cc_isr_Test_Fx.Assert"/>] instance where
''' <see cref="Assert.AssertSuccessful"/> is <c>True</c> if the test passed. </returns>
Public Function UnitTestSkeleton() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "UnitTestSkeleton"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = This.BeforeEachAssert
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestCreateSocket")
    
    Set TestCreateSocket = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error if not user defined error
    If Not cc_isr_Core_IO.UserDefinedErrors.IsUserDefinedError(VBA.Err.Number) Then _
        cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Function

