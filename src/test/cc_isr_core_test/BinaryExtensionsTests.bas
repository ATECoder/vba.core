Attribute VB_Name = "BinaryExtensionsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Binary extension methods. </summary>
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
            Set p_outcome = TestBitsShouldInvert
        Case 2
            Set p_outcome = TestBinaryValuesShouldAdd
        Case 3
            Set p_outcome = TestLongValueShouldConvertToBinary
        Case 4
            Set p_outcome = TestFractionalValueShouldConvertToBinary
        Case 5
            Set p_outcome = TestDoubleValueShouldConvertToBinary
        Case Else
    End Select
    Set RunTest = p_outcome
    AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 4
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
    This.TestCount = 5
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
    
    ' clear overall
    
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


''' <summary>   Unit test. Binary bits should invert. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestBitsShouldInvert() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestBitsShouldInvert"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_initialValue As String
    Dim p_expectedValue As String
    Dim p_actualValue As String
    
    p_initialValue = "101"
    p_expectedValue = "010"
    p_actualValue = cc_isr_Core.BinaryExtensions.InvertBits(p_initialValue)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "Actual value should equal the inverted bits of '" & p_initialValue & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = "0"
        p_expectedValue = "1"
        p_actualValue = cc_isr_Core.BinaryExtensions.InvertBits(p_initialValue)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the inverted bits of '" & p_initialValue & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = "00"
        p_expectedValue = "11"
        p_actualValue = cc_isr_Core.BinaryExtensions.InvertBits(p_initialValue)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the inverted bits of '" & p_initialValue & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = "11"
        p_expectedValue = "00"
        p_actualValue = cc_isr_Core.BinaryExtensions.InvertBits(p_initialValue)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the inverted bits of '" & p_initialValue & "'.")
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestBitsShouldInvert = p_outcome
    
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

''' <summary>   Unit test. Binary values should add. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestBinaryValuesShouldAdd() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestBinaryValuesShouldAdd"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_initialValue1 As String
    Dim p_initialValue2 As String
    Dim p_expectedValue As String
    Dim p_actualValue As String
    
    p_initialValue1 = "1"
    p_initialValue2 = "0"
    p_expectedValue = "1"
    p_actualValue = cc_isr_Core.BinaryExtensions.AddBinary(p_initialValue1, p_initialValue2)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "Actual value should equal the sum of '" & p_initialValue1 & "' and '" & p_initialValue2 & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue1 = "10"
        p_initialValue2 = "0"
        p_expectedValue = "10"
        p_actualValue = cc_isr_Core.BinaryExtensions.AddBinary(p_initialValue1, p_initialValue2)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the sum of '" & p_initialValue1 & "' and '" & p_initialValue2 & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue1 = "10"
        p_initialValue2 = "10"
        p_expectedValue = "100"
        p_actualValue = cc_isr_Core.BinaryExtensions.AddBinary(p_initialValue1, p_initialValue2)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the sum of '" & p_initialValue1 & "' and '" & p_initialValue2 & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue1 = "11"
        p_initialValue2 = "10"
        p_expectedValue = "101"
        p_actualValue = cc_isr_Core.BinaryExtensions.AddBinary(p_initialValue1, p_initialValue2)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the sum of '" & p_initialValue1 & "' and '" & p_initialValue2 & "'.")
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestBinaryValuesShouldAdd = p_outcome
    
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

''' <summary>   Unit test. Long decimal value should convert to binary. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestLongValueShouldConvertToBinary() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestLongValueShouldConvertToBinary"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_initialValue As Integer
    Dim p_expectedValue As String
    Dim p_actualValue As String
    
    p_initialValue = 0
    p_expectedValue = "0000"
    p_actualValue = cc_isr_Core.BinaryExtensions.LongToBinary(p_initialValue, 4)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 1
        p_expectedValue = "0001"
        p_actualValue = cc_isr_Core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 5
        p_expectedValue = "0101"
        p_actualValue = cc_isr_Core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -1
        p_expectedValue = 1111
        p_actualValue = cc_isr_Core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -2
        p_expectedValue = 1110
        p_actualValue = cc_isr_Core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -3
        p_expectedValue = 1101
        p_actualValue = cc_isr_Core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestLongValueShouldConvertToBinary = p_outcome
    
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

''' <summary>   Unit test. Fractional decimal value should convert to binary. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestFractionalValueShouldConvertToBinary() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestFractionalValueShouldConvertToBinary"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_initialValue As Double
    Dim p_expectedValue As String
    Dim p_actualValue As String
    
    p_initialValue = 0
    p_expectedValue = "0000"
    p_actualValue = cc_isr_Core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 0.5
        p_expectedValue = "1000"
        p_actualValue = cc_isr_Core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 0.25
        p_expectedValue = "0100"
        p_actualValue = cc_isr_Core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 0.75
        p_expectedValue = 1100
        p_actualValue = cc_isr_Core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    ' assert any errors here before throwing an exception in the next test.
    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    If p_outcome.AssertSuccessful Then
    
        On Error Resume Next
        p_initialValue = -0.75
        p_actualValue = cc_isr_Core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(cc_isr_Core_IO.UserDefinedErrors.InvalidArgumentError.Number, _
                Err.Number, _
                "Attempting to convert a negative fractional should raise the Invalid Argument Error code.")
        
        On Error GoTo 0
    
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(1, cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount, _
                "A single error should be enqueued after testing for invalid argument exception.")
    End If
    
    If p_outcome.AssertSuccessful Then
        ' clear the error state.
        cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestFractionalValueShouldConvertToBinary = p_outcome
    
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

''' <summary>   Unit test. Double decimal value should convert to binary. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestDoubleValueShouldConvertToBinary() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestDoubleValueShouldConvertToBinary"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_initialValue As Double
    Dim p_expectedValue As String
    Dim p_actualValue As String
    
    p_initialValue = 0
    p_expectedValue = "0000.0000"
    p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 1.5
        p_expectedValue = "0001.1000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 1.25
        p_expectedValue = "0001.0100"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -1#
        p_expectedValue = "1111.0000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -1.5
        p_expectedValue = "1110.1000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 12.5
        p_expectedValue = "01100.10000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 12.25
        p_expectedValue = "01100.01000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 12.75
        p_expectedValue = "01100.11000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -12.5
        p_expectedValue = "10011.10000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -12.25
        p_expectedValue = "10011.11000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then

        p_initialValue = -12.75
        p_expectedValue = "10011.01000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestDoubleValueShouldConvertToBinary = p_outcome
    
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


