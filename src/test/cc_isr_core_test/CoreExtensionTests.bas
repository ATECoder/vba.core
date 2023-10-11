Attribute VB_Name = "CoreExtensionTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Core extension methods. </summary>
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

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Test runners
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Runs the specified test. </summary>
Public Function RunTest(ByVal a_testNumber As Integer) As cc_isr_Test_Fx.Assert
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.TestNumber = a_testNumber
    BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestWaitShouldEqualOrExceedDuration
        Case 2
            Set p_outcome = TestNowResolution
        Case 3
            Set p_outcome = TestDefaultValues
        Case 4
            Set p_outcome = TestParameterArrayPropagated
        Case 5
            Set p_outcome = TestByteShouldClamp
        Case 6
            Set p_outcome = TestDoubleShouldClamp
        Case 7
            Set p_outcome = TestIntegerShouldClamp
        Case 8
            Set p_outcome = TestLongShouldClamp
        Case 9
            Set p_outcome = TestSingleShouldClamp
        Case 10
            Set p_outcome = TestVariantShouldClamp
        Case Else
    End Select
    Set RunTest = p_outcome
    AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 10
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    This.Name = "CoreExtensionTests"
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 10
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

''' <summary>   Unit test. Asserts that a wait time should be longer or equal to the expected duration. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestWaitShouldEqualOrExceedDuration() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestWaitShouldEqualOrExceedDuration"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_expectedDuration As Double
    p_expectedDuration = 0.1
    Dim p_actualDuration As Double: p_actualDuration = cc_isr_Core_IO.CoreExtensions.Wait(p_expectedDuration)
    Set p_outcome = Assert.IsTrue(p_expectedDuration <= p_actualDuration, _
        "Wait time " & CStr(p_actualDuration) & " should be equal ot longer than the specified duration of " & CStr(p_expectedDuration) & " .")

    If p_outcome.AssertSuccessful Then
        p_actualDuration = cc_isr_Core_IO.CoreExtensions.Wait(p_expectedDuration)
        Set p_outcome = Assert.IsTrue(2 * p_expectedDuration > p_actualDuration, _
            "Wait time " & CStr(p_actualDuration) & " should be shorter that double the specified duration of " & CStr(2 * p_expectedDuration) & " .")
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestWaitShouldEqualOrExceedDuration = p_outcome
    
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

''' <summary>   Unit test. Asserts that the resolution of VBA.Now() should be longer or equal
'''             to the expected resolution but smaller than double of that resoltion. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestNowResolution() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestNowResolution"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    '  loop until now changes
    Dim p_startTime As Double: p_startTime = cc_isr_Core_IO.CoreExtensions.DaysNow()
    Dim p_endTime As Double: p_endTime = cc_isr_Core_IO.CoreExtensions.DaysNow()
    While p_startTime = p_endTime
        'DoEvents
        p_endTime = cc_isr_Core_IO.CoreExtensions.DaysNow()
    Wend
    
    p_startTime = cc_isr_Core_IO.CoreExtensions.DaysNow()
    p_endTime = cc_isr_Core_IO.CoreExtensions.DaysNow()
    While p_startTime = p_endTime
        'DoEvents
        p_endTime = cc_isr_Core_IO.CoreExtensions.DaysNow()
    Wend
    Dim p_resolution As Double
    p_resolution = cc_isr_Core_IO.CoreExtensions.SecondsPerDay() * (p_endTime - p_startTime)
    
    Dim p_expectedResolution As Double: p_expectedResolution = 0.95 * cc_isr_Core_IO.CoreExtensions.TimerResolution
    Set p_outcome = Assert.IsTrue(p_expectedResolution <= p_resolution, _
        "Actual resolution " & CStr(p_resolution) & " should be equal ot larger than the adjusted expected resolution of " & CStr(p_expectedResolution) & ".")
    
    ' Debug.Print p_resolution, p_expectedResolution
    
    If p_outcome.AssertSuccessful Then
    
        p_expectedResolution = 2# * cc_isr_Core_IO.CoreExtensions.TimerResolution
        Set p_outcome = Assert.IsTrue(p_expectedResolution >= p_resolution, _
            "Actual resolution " & CStr(p_resolution) & " should be equal ot smaller than the twice the expected resolution of " & CStr(p_expectedResolution) & ".")
    
        ' Debug.Print p_resolution, p_expectedResolution
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestNowResolution = p_outcome
    
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

''' <summary>   Unit test. Asserts that default values are as expected. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestDefaultValues() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestDefaultValues"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = Assert.AreEqual(False, cc_isr_Core_IO.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbBoolean), _
        "The default value of VBA.VbVarType.vbBoolean should equal.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(0, cc_isr_Core_IO.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbByte), _
            "The default value of VBA.VbVarType.vbByte should equal.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(Empty, cc_isr_Core_IO.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbArray), _
            "The default value of VBA.VbVarType.vbArray should equal.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsNull(cc_isr_Core_IO.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbNull), _
            "The default value of VBA.VbVarType.vbNull should equal.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsNull(cc_isr_Core_IO.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbNull), _
            "The default value of VBA.VbVarType.vbNull should be Null.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.IsTrue(cc_isr_Core_IO.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbObject) Is Nothing, _
            "The default value of VBA.VbVarType.vbObject should be nothing.")
    
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = Assert.AreEqual(0, cc_isr_Core_IO.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbLongLong), _
            "The default value of VBA.VbVarType.vbLongLong should equal.")
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestDefaultValues = p_outcome
    
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

''' <summary>   Unit test. Asserts that the paramter array propagated through nested methods
''' without errors. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestParameterArrayPropagated() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "TestParameterArrayPropagated"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_dummyVariant As Variant
    p_dummyVariant = "a"
    Dim p_dummyArray() As Variant
    Dim p_unboxedTokens As Variant
    
    Dim p_tokens As Variant
    p_tokens = Method1("a", "b", "c")
    
    Set p_outcome = Assert.AreEqual(VBA.Err.Number, 0, _
        "The parameter array should pass without errors")
        
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.AreEqual(TypeName(p_dummyArray), TypeName(p_tokens), _
        "The nested parameter array type should match the expected type")
    
    End If
        
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.AreEqual(TypeName(p_dummyArray), TypeName(p_tokens(0)), _
        "The first element of the nested parameter array type should match the expected type")
    
    End If
        
    If p_outcome.AssertSuccessful Then
    
        p_unboxedTokens = CoreExtensions.UnboxParameterArray(p_tokens)
        
        Set p_outcome = Assert.AreEqual(TypeName(p_dummyArray), TypeName(p_unboxedTokens), _
        "The unboxed parameter array type should match the expected type")
        
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.AreEqual(TypeName(p_dummyVariant), TypeName(p_unboxedTokens(0)), _
        "The first element of the nested parameter array type should match the expected type")
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestParameterArrayPropagated = p_outcome
    
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

Public Function Method1(ParamArray a_tokens() As Variant) As Variant
    Method1 = Method2(a_tokens)
End Function

Public Function Method2(ParamArray a_tokens() As Variant) As Variant
    Method2 = Method3(a_tokens)
End Function

Public Function Method3(ParamArray a_tokens() As Variant) As Variant
    Method3 = a_tokens
End Function

Public Sub MethodA(ParamArray a_tokens() As Variant)
    Dim p_tokens() As Variant
    p_tokens = CoreExtensions.UnboxParameterArray(a_tokens)
    MethodB a_tokens
End Sub

Public Sub MethodB(ParamArray a_tokens() As Variant)
    Dim i As Integer, p_tokens() As Variant
    p_tokens = CoreExtensions.UnboxParameterArray(a_tokens)
    For i = 0 To UBound(p_tokens)
        Debug.Print StringExtensions.StringFormat("i: {0} prm: {1} ", i, p_tokens(i))
    Next i
    MethodC a_tokens
End Sub

Public Sub MethodC(ParamArray a_tokens() As Variant)
    Dim i As Integer, p_tokens() As Variant
    p_tokens = CoreExtensions.UnboxParameterArray(a_tokens)
    For i = 0 To UBound(p_tokens)
        Debug.Print StringExtensions.StringFormat("i: {0} prm: {1} ", i, p_tokens(i))
    Next i
End Sub

''' <summary>   Unit test. Asserts that a byte should clamp. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestByteShouldClamp() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestByteShouldClamp"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    Dim p_value As Byte
    Dim p_min As Byte
    Dim p_max As Byte
    Dim p_expected As Byte
    
    If p_outcome.AssertSuccessful Then
        p_value = 2
        p_min = 1
        p_max = 10
        p_expected = 2
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampByte(p_value, p_min, p_max), _
            "ClampByte( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
        
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 0
        p_min = 1
        p_max = 10
        p_expected = 1
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampByte(p_value, p_min, p_max), _
            "ClampByte( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 1
        p_min = 1
        p_max = 10
        p_expected = 1
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampByte(p_value, p_min, p_max), _
            "ClampByte( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 11
        p_min = 1
        p_max = 10
        p_expected = 10
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampByte(p_value, p_min, p_max), _
            "ClampByte( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 10
        p_min = 1
        p_max = 10
        p_expected = 10
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampByte(p_value, p_min, p_max), _
            "ClampByte( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestByteShouldClamp = p_outcome

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

''' <summary>   Unit test. Asserts that a Double should clamp. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestDoubleShouldClamp() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestDoubleShouldClamp"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    Dim p_value As Double
    Dim p_min As Double
    Dim p_max As Double
    Dim p_expected As Double
    
    If p_outcome.AssertSuccessful Then
        p_value = 2
        p_min = 1
        p_max = 10
        p_expected = 2
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampDouble(p_value, p_min, p_max), _
            "ClampDouble( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
        
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 0
        p_min = 1
        p_max = 10
        p_expected = 1
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampDouble(p_value, p_min, p_max), _
            "ClampDouble( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 1
        p_min = 1
        p_max = 10
        p_expected = 1
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampDouble(p_value, p_min, p_max), _
            "ClampDouble( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 11
        p_min = 1
        p_max = 10
        p_expected = 10
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampDouble(p_value, p_min, p_max), _
            "ClampDouble( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 10
        p_min = 1
        p_max = 10
        p_expected = 10
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampDouble(p_value, p_min, p_max), _
            "ClampDouble( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestDoubleShouldClamp = p_outcome

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

''' <summary>   Unit test. Asserts that a Integer should clamp. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestIntegerShouldClamp() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestIntegerShouldClamp"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    Dim p_value As Integer
    Dim p_min As Integer
    Dim p_max As Integer
    Dim p_expected As Integer
    
    If p_outcome.AssertSuccessful Then
        p_value = 2
        p_min = 1
        p_max = 10
        p_expected = 2
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampInteger(p_value, p_min, p_max), _
            "ClampInteger( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
        
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 0
        p_min = 1
        p_max = 10
        p_expected = 1
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampInteger(p_value, p_min, p_max), _
            "ClampInteger( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 1
        p_min = 1
        p_max = 10
        p_expected = 1
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampInteger(p_value, p_min, p_max), _
            "ClampInteger( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 11
        p_min = 1
        p_max = 10
        p_expected = 10
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampInteger(p_value, p_min, p_max), _
            "ClampInteger( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 10
        p_min = 1
        p_max = 10
        p_expected = 10
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampInteger(p_value, p_min, p_max), _
            "ClampInteger( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestIntegerShouldClamp = p_outcome

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

''' <summary>   Unit test. Asserts that a Long should clamp. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestLongShouldClamp() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestLongShouldClamp"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    Dim p_value As Long
    Dim p_min As Long
    Dim p_max As Long
    Dim p_expected As Long
    
    If p_outcome.AssertSuccessful Then
        p_value = 2
        p_min = 1
        p_max = 10
        p_expected = 2
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampLong(p_value, p_min, p_max), _
            "ClampLong( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
        
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 0
        p_min = 1
        p_max = 10
        p_expected = 1
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampLong(p_value, p_min, p_max), _
            "ClampLong( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 1
        p_min = 1
        p_max = 10
        p_expected = 1
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampLong(p_value, p_min, p_max), _
            "ClampLong( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 11
        p_min = 1
        p_max = 10
        p_expected = 10
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampLong(p_value, p_min, p_max), _
            "ClampLong( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 10
        p_min = 1
        p_max = 10
        p_expected = 10
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampLong(p_value, p_min, p_max), _
            "ClampLong( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestLongShouldClamp = p_outcome

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

''' <summary>   Unit test. Asserts that a Single should clamp. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestSingleShouldClamp() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestSingleShouldClamp"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    Dim p_value As Single
    Dim p_min As Single
    Dim p_max As Single
    Dim p_expected As Single
    
    If p_outcome.AssertSuccessful Then
        p_value = 2
        p_min = 1
        p_max = 10
        p_expected = 2
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampSingle(p_value, p_min, p_max), _
            "ClampSingle( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
        
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 0
        p_min = 1
        p_max = 10
        p_expected = 1
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampSingle(p_value, p_min, p_max), _
            "ClampSingle( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 1
        p_min = 1
        p_max = 10
        p_expected = 1
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampSingle(p_value, p_min, p_max), _
            "ClampSingle( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 11
        p_min = 1
        p_max = 10
        p_expected = 10
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampSingle(p_value, p_min, p_max), _
            "ClampSingle( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 10
        p_min = 1
        p_max = 10
        p_expected = 10
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.ClampSingle(p_value, p_min, p_max), _
            "ClampSingle( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestSingleShouldClamp = p_outcome

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

''' <summary>   Unit test. Asserts that a Variant should clamp. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestVariantShouldClamp() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestVariantShouldClamp"

    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As Assert: Set p_outcome = Assert.Pass("Entered the " & p_procedureName & " test.")
    
    Dim p_value As Variant
    Dim p_min As Variant
    Dim p_max As Variant
    Dim p_expected As Variant
    
    If p_outcome.AssertSuccessful Then
        p_value = 2
        p_min = 1
        p_max = 10
        p_expected = 2
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.Clamp(p_value, p_min, p_max), _
            "Clamp( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
        
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 0
        p_min = 1
        p_max = 10
        p_expected = 1
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.Clamp(p_value, p_min, p_max), _
            "Clamp( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 1
        p_min = 1
        p_max = 10
        p_expected = 1
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.Clamp(p_value, p_min, p_max), _
            "Clamp( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 11
        p_min = 1
        p_max = 10
        p_expected = 10
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.Clamp(p_value, p_min, p_max), _
            "Clamp( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_value = 10
        p_min = 1
        p_max = 10
        p_expected = 10
        Set p_outcome = Assert.AreEqual(p_expected, cc_isr_Core_IO.CoreExtensions.Clamp(p_value, p_min, p_max), _
            "Clamp( " & VBA.CStr(p_value) & ", " & VBA.CStr(p_min) & ", " & VBA.CStr(p_max) & _
            ") should equal expected value.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestVariantShouldClamp = p_outcome

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


