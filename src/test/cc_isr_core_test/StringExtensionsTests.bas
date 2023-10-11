Attribute VB_Name = "StringExtensionsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. String extension methods. </summary>
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
            Set p_outcome = TestCharAt
        Case 2
            Set p_outcome = TestEndsWith
        Case 3
            Set p_outcome = TestEscapeSequences
        Case 4
            Set p_outcome = TestInsertRepelaceEscapeSequences
        Case 5
            Set p_outcome = TestInsert
        Case 6
            Set p_outcome = TestPop
        Case 7
            Set p_outcome = TestRemoveWhiteSpaces
        Case 8
            Set p_outcome = TestRepeat
        Case 9
            Set p_outcome = TestStartsWith
        Case 10
            Set p_outcome = TestFormatStringParser
        Case 11
            Set p_outcome = TestDateStringFormat
        Case 12
            Set p_outcome = TestStringFormat
        Case 13
            Set p_outcome = TestStringFormatReplace
        Case 14
            Set p_outcome = TestStringContains
        Case 15
            Set p_outcome = TestStringContainsAny
        Case 16
            Set p_outcome = TestSubstring
        Case 17
            Set p_outcome = TestToBinary
        Case 18
            Set p_outcome = TestTrimLeft
        Case 19
            Set p_outcome = TestTrimRight
        Case 20
            Set p_outcome = TestTextShouldParseToBoolean
        Case 21
            Set p_outcome = TestTextShouldParseToDouble
        Case 22
            Set p_outcome = TestTextShouldParseToLong
        Case 23
            Set p_outcome = TestTrimEnd
        Case 24
            Set p_outcome = TestTrimStart
        Case Else
    End Select
    Set RunTest = p_outcome
    AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 20
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
    This.TestCount = 24
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

''' <summary>   Unit test. Asserts character at an index position. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCharAt() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestCharAt"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("a", StringExtensions.CharAt("foobar", 5), _
            "Should get the expected character from the string")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestCharAt = p_outcome
    
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

''' <summary>   Unit test. Asserts end width. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestEndsWith() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestEndsWith"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(StringExtensions.EndsWith("foobar", "bar"), _
            "String should end with the expected value")
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestEndsWith = p_outcome
    
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

''' <summary>   Unit test. Asserts escape sequences existence and values. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestEscapeSequences() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestEscapeSequences"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_escapes As Collection
    Set p_escapes = cc_isr_Core.StringExtensions.EscapeSequences
    Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNull(p_escapes, _
            "Escape sequences should be created")
            
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(8, p_escapes.Count, _
                "Number of escape sequences should match")
                
    End If

    If p_outcome.AssertSuccessful Then
    
        Dim p_escape As EscapeSequence
        Dim p_item As EscapeSequence
        For Each p_escape In p_escapes
        
            Set p_item = p_escapes(p_escape.value)
            
            Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_escape.value, p_item.value, _
                    "For each escape value must match collection item value")
                    
            If Not p_outcome.AssertSuccessful Then Exit For
        
            Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_escape.Literal, p_item.Literal, _
                    "For each escape replacement value must match collection item replacement value")
                    
            If Not p_outcome.AssertSuccessful Then Exit For
        
        Next
        
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestEscapeSequences = p_outcome
    
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

''' <summary>   Unit test. Asserts inserting. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestInsertRepelaceEscapeSequences() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestInsertRepelaceEscapeSequences"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_original As String: p_original = "\t1234\r\n"
    Dim p_expected As String: p_expected = VBA.vbTab & "1234" & VBA.Chr$(13) & VBA.Chr$(10)
    Dim p_actual As String: p_actual = cc_isr_Core.StringExtensions.ReplaceEscapeSequences(p_original)

    Dim p_areEqual As Boolean: p_areEqual = cc_isr_Core.StringExtensions.AreEqualDebug(p_expected, p_actual)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, p_actual, _
            "Literal values should be inserted in place of the escaped sequences.")
            
    If p_outcome.AssertSuccessful Then
    
        p_expected = p_original
        p_original = p_actual
        p_actual = cc_isr_Core.StringExtensions.InsertEscapeSequences(p_original)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, p_actual, _
                "Escape sequences should be inserted in place of the literal characters.")
    End If

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestInsertRepelaceEscapeSequences = p_outcome
    
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

''' <summary>   Unit test. Asserts inserting. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestInsert() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "TestInsert"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_original As String: p_original = "1234"
    Dim p_added As String: p_added = "99"
    
    Dim p_expected As String
    Dim p_position As Long
    Dim p_suffix As String
    
    p_position = 0: p_expected = "991234": p_suffix = "-th"
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
            StringExtensions.Insert(p_original, p_added, p_position), _
            "Added string '" & p_added & "' should be inserted into '" & _
            p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
            
    If p_outcome.AssertSuccessful Then
        p_position = 1: p_expected = "991234": p_suffix = "-st"
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_position = 2: p_expected = "199234": p_suffix = "-nd"
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_position = 3: p_expected = "129934": p_suffix = "-rd"
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_position = 4: p_expected = "123994": p_suffix = "-th"
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_position = 100: p_expected = "123499": p_suffix = "-th (after the last)"
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestInsert = p_outcome
    
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

''' <summary>   Unit test. Asserts delimited string element should pop. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestPop() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "TestPop"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_delimitedString As String: p_delimitedString = "a,b,c"
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("a", _
            StringExtensions.pop(p_delimitedString, ","), _
            "First element in " & p_delimitedString & " should pop")
            
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("b", _
            StringExtensions.pop(p_delimitedString, ","), _
            "Second element in " & p_delimitedString & " should pop")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("c", _
            StringExtensions.pop(p_delimitedString, ","), _
            "Third element in " & p_delimitedString & " should pop")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, _
            StringExtensions.pop(p_delimitedString, ","), _
            "No element in " & p_delimitedString & " should pop")
            
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestPop = p_outcome
    
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

''' <summary>   Unit test. Asserts removing white spaces. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestRemoveWhiteSpaces() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "TestRemoveWhiteSpaces"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_original As String
    Dim p_expected As String
    Dim p_actual As String
    
    p_original = "Hello world!"
    p_expected = "Helloworld!"
    p_actual = cc_isr_Core.StringExtensions.RemoveWhiteSpace(p_original)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, p_actual, _
            "Should remove all white spaces from '" & p_original & "'.")
            
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestRemoveWhiteSpaces = p_outcome
    
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

''' <summary>   Unit test. Asserts creating a repeated string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestRepeat() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "TestRepeat"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("aaa", StringExtensions.Repeat("a", 3), _
            "Should constract a string with repreated strings")
            
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestRepeat = p_outcome
    
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

''' <summary>   Unit test. Asserts start with. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStartsWith() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestStartsWith"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(StringExtensions.StartsWith("foobar", "foo"), _
            "String should start with the expected value.")
            
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestStartsWith = p_outcome
    
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

''' <summary>   Unit test. Asserts parsing format string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestFormatStringParser() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestFormatStringParser"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_format As String
    Dim p_expected As String
    Dim p_actual As String
    Dim p_dateTime As Date
    Dim p_stringValue As String
    
    Dim p_formatGroup As String, p_precisionSpecifier As Integer
    Dim p_formatSpecifier As String, p_precisionString As String
    Dim p_itemIndex As Integer, p_success As Boolean, p_message As String

    Dim p_expectedFormatGroup As String
    Dim p_expectedFormatSpecifier As String, p_expectedPrecisionString As String
    Dim p_expectedItemIndex As Integer

    p_format = "{0:F11}"
    p_expectedItemIndex = 0
    p_expectedFormatGroup = "F11"
    p_expectedPrecisionString = "11"
    p_expectedFormatSpecifier = "F"
    
    p_success = cc_isr_Core.StringExtensions.ParseFormatSpecification(p_format, p_itemIndex, p_formatGroup, _
                                         p_precisionString, p_formatSpecifier, _
                                         p_message)
        
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, "'" & p_format & "' should parse: " & p_message)
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedItemIndex, p_itemIndex, _
        "'" & p_format & "' item index should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFormatGroup, p_formatGroup, _
            "'" & p_format & "' format group should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedPrecisionString, p_precisionString, _
            "'" & p_format & "' precision string should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFormatSpecifier, p_formatSpecifier, _
            "'" & p_format & "' format specifier should equal.")
    End If
    
    p_format = "{1:F1}"
    p_expectedItemIndex = 1
    p_expectedFormatGroup = "F1"
    p_expectedPrecisionString = "1"
    p_expectedFormatSpecifier = "F"
    
    p_success = cc_isr_Core.StringExtensions.ParseFormatSpecification(p_format, p_itemIndex, p_formatGroup, _
                                         p_precisionString, p_formatSpecifier, _
                                         p_message)
        
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, "'" & p_format & "' should parse: " & p_message)
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedItemIndex, p_itemIndex, _
        "'" & p_format & "' item index should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFormatGroup, p_formatGroup, _
            "'" & p_format & "' format group should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedPrecisionString, p_precisionString, _
            "'" & p_format & "' precision string should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFormatSpecifier, p_formatSpecifier, _
            "'" & p_format & "' format specifier should equal.")
    End If
    
    p_format = "{2:MMMM dd, yyyy}"
    p_expectedItemIndex = 2
    p_expectedFormatGroup = "MMMM dd, yyyy"
    p_expectedPrecisionString = ""
    p_expectedFormatSpecifier = "MMMM dd, yyyy"
    
    p_success = cc_isr_Core.StringExtensions.ParseFormatSpecification(p_format, p_itemIndex, p_formatGroup, _
                                         p_precisionString, p_formatSpecifier, _
                                         p_message)
        
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, "'" & p_format & "' should parse: " & p_message)
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedItemIndex, p_itemIndex, _
        "'" & p_format & "' item index should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFormatGroup, p_formatGroup, _
            "'" & p_format & "' format group should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedPrecisionString, p_precisionString, _
            "'" & p_format & "' precision string should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFormatSpecifier, p_formatSpecifier, _
            "'" & p_format & "' format specifier should equal.")
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestFormatStringParser = p_outcome
    
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


''' <summary>   Unit test. Asserts creating formatted date strings. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance where
'''             <see cref="Assert.AssertSuccessful"/> is True if the test passed. </returns>
Public Function TestDateStringFormat() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestDateStringFormat"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_format As String
    Dim p_expected As String
    Dim p_actual As String
    Dim p_dateTime As Date
    Dim p_stringValue As String
    
    p_stringValue = "12:00:00 AM"
    p_dateTime = CDate(p_stringValue)
    p_format = "{0:MMMM dd, yyyy}"
    p_expected = "December 30, 1899"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_dateTime)
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
            p_actual, "'" & p_format & "' should format CDate('" & p_stringValue & "' as expected.")
    
    p_stringValue = "12:00:00 AM"
    p_format = "{0:MMMM dd, yyyy}"
    p_expected = "December 30, 1899"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_stringValue)
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
            p_actual, "'" & p_format & "' should format '" & p_stringValue & "' as expected.")
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestDateStringFormat = p_outcome
    
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

''' <summary>   Unit test. Asserts creating a formatted string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance where
'''             <see cref="Assert.AssertSuccessful"/> is True if the test passed. </returns>
Public Function TestStringFormat() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestStringFormat"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_format As String
    Dim p_expected As String
    Dim p_actual As String
    
    p_format = "a{0}{1}"
    p_expected = "abc"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, "b", "c")
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")
            
    If p_outcome.AssertSuccessful Then
        p_format = "(B) Binary: {0:B}"
        p_expected = "(B) Binary: 10000101"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(B) Binary: {0:B16}"
        p_expected = "(B) Binary: 1111111110000101"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If
            
    If p_outcome.AssertSuccessful Then
        p_format = "(C) Currency: {0:C}\n"
        p_expected = "(C) Currency: -123.45$" & VBA.vbLf
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(C) Currency: {0:C}"
        p_expected = "(C) Currency: -123.00$"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_format = "(D) Decimal:. . . . . . . . . {0:D}"
        p_expected = "(D) Decimal:. . . . . . . . . -123"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If
    

    If p_outcome.AssertSuccessful Then
        p_format = "(E) Scientific: . . . . . . . {0:E}"
        p_expected = "(E) Scientific: . . . . . . . -1.23450E2"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_format = "(F) Fixed point:. . . . . . . {0:F}"
        p_expected = "(F) Fixed point:. . . . . . . -123.45"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_format = "(F) Fixed point:. . . . . . . {0:F1}"
        p_expected = "(F) Fixed point:. . . . . . . -123.5"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(G) General:. . . . . . . {0:G}"
        p_expected = "(G) General:. . . . . . . -1.23450E2"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(G) General:. . . . . . . {0:G4}"
        p_expected = "(G) General:. . . . . . . -123.5"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(N) Number: . . . . . . . . . {0:N}"
        p_expected = "(N) Number: . . . . . . . . . -123"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(P) Percent:. . . . . . . . . {0:P}"
        p_expected = "(P) Percent:. . . . . . . . . -12,345%"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(R) Round-trip: . . . . . . . {0:R}"
        p_expected = "(R) Round-trip: . . . . . . . -123.45"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(X) Hexadecimal:. . . . . . . {0:X}"
        p_expected = "(X) Hexadecimal:. . . . . . . FF85"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, CInt(-123))
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(X) Hexadecimal:. . . . . . . {0:x}"
        p_expected = "(X) Hexadecimal:. . . . . . . ff85"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, CInt(-123))
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    Dim p_date As Date: p_date = DateValue("January 26, 2013") + TimeValue("8:28:11 PM")

    If p_outcome.AssertSuccessful Then
        p_format = "(c) Custom format: . . . . . .{0:cYYYY-MM-DD (MMMM)}"
        p_expected = "(c) Custom format: . . . . . .2013-01-26 (January)"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(d) Short date: . . . . . . . {0:d}"
        p_expected = "(d) Short date: . . . . . . . 1/26/2013"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(D) Long date:. . . . . . . . {0:D}"
        p_expected = "(D) Long date:. . . . . . . . Saturday, January 26, 2013"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    Dim p_time As Date
    p_time = TimeValue("8:28:11 PM")
    
    If p_outcome.AssertSuccessful Then
        p_format = "(T) Long time:. . . . . . . . {0:T}"
        p_expected = "(T) Long time:. . . . . . . . 8:28:11 PM"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(f) Full date/short time: . . {0:f}"
        p_expected = "(f) Full date/short time: . . Saturday, January 26, 2013 8:28 PM"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(F) Full date/long time:. . . {0:F}"
        p_expected = "(F) Full date/long time:. . . Saturday, January 26, 2013 8:28:11 PM"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(s) Sortable: . . . . . . . . {0:s}"
        p_expected = "(s) Sortable: . . . . . . . . 2013-01-26T20:28:11"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        ' specify alignment (/padding) and to use escape sequences:
        
        p_format = "\q{0}, {1}!\x20\n'{2,10:C2}'\n'{2,-10:C2}'"
        p_expected = """hello, world! " & VBA.vbLf & "'   100.00$'" & VBA.vbLf & "'100.00$   '"
        p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, "hello", "world", 100)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If
  
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestStringFormat = p_outcome
    
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

''' <summary>   Unit test. Asserts creating a formatted string using simpel replacement. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStringFormatReplace() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestStringFormatReplace"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("aaa", _
            StringExtensions.StringFormat("a{0}{1}", "a", "a"), _
            "Format should build the expected string")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestStringFormatReplace = p_outcome
    
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

''' <summary>   Unit test. Asserts finding an item in a string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStringContains() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestStringContains"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_value As String
    Dim p_candidate As String
    p_value = "the string contains"
    p_candidate = "contains"
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue( _
            StringExtensions.StringContains(p_value, p_candidate), _
            "The string '" & p_value & "' should contain '" & p_candidate & "'.")
            
    If p_outcome.AssertSuccessful Then

        p_value = "the string contains"
        p_candidate = "contained"
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse( _
                StringExtensions.StringContains(p_value, p_candidate), _
                "The string '" & p_value & "' should not contain '" & p_candidate & "'.")

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestStringContains = p_outcome
    
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

''' <summary>   Unit test. Asserts finding items in a string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStringContainsAny() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestStringContainsAny"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_value As String
    Dim p_candidate1 As String, p_candidate2 As String
    p_value = "the string contains"
    p_candidate1 = "the"
    p_candidate2 = "contains"
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue( _
            StringExtensions.StringContainsAny(p_value, True, p_candidate1, p_candidate2), _
            "The string '" & p_value & "' should contain '" & p_candidate1 & "' or '" & p_candidate2 & "'.")
            
    If p_outcome.AssertSuccessful Then

        p_value = "the string contains"
        p_candidate1 = "the"
        p_candidate2 = "contained"
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue( _
                StringExtensions.StringContainsAny(p_value, True, p_candidate1, p_candidate2), _
                "The string '" & p_value & "' should contain '" & p_candidate1 & "' or '" & p_candidate2 & "'.")

    End If
    
    If p_outcome.AssertSuccessful Then

        p_value = "the string contains"
        p_candidate1 = "not"
        p_candidate2 = "contained"
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse( _
                StringExtensions.StringContainsAny(p_value, True, p_candidate1, p_candidate2), _
                "The string '" & p_value & "' should not contain '" & p_candidate1 & "' or '" & p_candidate2 & "'.")

    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestStringContainsAny = p_outcome
    
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

''' <summary>   Unit test. Asserts sub-string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestSubstring() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "TestSubstring"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("oo", StringExtensions.Substring("foobar", 1, 2), _
            "Should get the expected part of the string")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestSubstring = p_outcome
    
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

''' <summary>   Unit test. Asserts convertinh values to binary. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance where
'''             <see cref="Assert.AssertSuccessful"/> is True if the test passed. </returns>
Public Function TestToBinary() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestToBinary"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_value As Long
    Dim p_expected As String
    Dim p_actual As String
    
    p_value = 5
    p_expected = "101"
    p_actual = cc_isr_Core.StringExtensions.ToBinary(p_value)
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
            p_actual, "decimal value '" & CStr(p_value) & "' should convert to as expected.")
    
    If p_outcome.AssertSuccessful Then
        p_value = 16
        p_expected = "10000"
        p_actual = cc_isr_Core.StringExtensions.ToBinary(p_value)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "decimal value '" & CStr(p_value) & "' should convert to as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_value = 5
        p_expected = "00000101"
        p_actual = cc_isr_Core.StringExtensions.ToBinary(p_value, 8)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "decimal value '" & CStr(p_value) & "' should convert to as expected.")
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestToBinary = p_outcome
    
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

''' <summary>   Unit test. Asserts trim left. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTrimLeft() As cc_isr_Test_Fx.Assert
    
    Const p_procedureName As String = "TestTrimLeft"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("bar", _
        StringExtensions.TrimLeft("oobar", "o"), "String should be left trimmed.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestTrimLeft = p_outcome
    
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

''' <summary>   Unit test. Asserts trim right. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTrimRight() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestTrimRight"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("f", _
        StringExtensions.TrimRight("foo", "o"), "String should be right-trimmed.")

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestTrimRight = p_outcome
    
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

''' <summary>   Unit test. Trim end should pass. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTrimEnd() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestTrimEnd"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_trim As String: p_trim = VBA.vbCrLf
    Dim p_expected As String: p_expected = "expected"
    Dim p_text As String: p_text = p_expected & p_trim
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(Len(p_trim) + VBA.Len(p_expected), VBA.Len(p_text), _
        "The length of the appended string should match the sum of the two strings.")

    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqualString(p_expected, _
            cc_isr_Core.StringExtensions.TrimEnd(p_text, p_trim), VBA.VbCompareMethod.vbTextCompare, _
            "The trimmed string should equal the expected string.")
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestTrimEnd = p_outcome
    
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

''' <summary>   Unit test. Trim start should pass. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTrimStart() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestTrimStart"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_trim As String: p_trim = VBA.vbCrLf
    Dim p_expected As String: p_expected = "expected"
    Dim p_text As String: p_text = p_trim & p_expected
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(Len(p_trim) + VBA.Len(p_expected), VBA.Len(p_text), _
        "The length of the appended string should match the sum of the two strings.")

    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqualString(p_expected, _
            cc_isr_Core.StringExtensions.TrimStart(p_text, p_trim), VBA.VbCompareMethod.vbTextCompare, _
            "The trimmed string should equal the expected string.")
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestTrimStart = p_outcome
    
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

''' + + + + + + + + + + + + + + + + + + + + + + + + + + + +
''' Parse numbers
''' + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Unit test. Asserts text should parse to Boolean. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTextShouldParseToBoolean() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestTextShouldParseToBoolean"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_text As String
    Dim p_expectedValue As Boolean
    Dim p_expectedOutcome As Boolean
    Dim p_actualOutcome As Boolean
    Dim p_actualValue As Boolean
    Dim p_actualDetails As String
    
    p_text = "1"
    p_expectedValue = True
    p_expectedOutcome = True
    p_actualOutcome = StringExtensions.TryParseBoolean(p_text, p_actualValue, p_actualDetails)
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, _
            "parse of '" & p_text & "' to boolean should " & _
            IIf(p_expectedOutcome, "succeed", "fail") & "; details: '" & p_actualDetails & "'.")

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_text = "fortytwo"
        ' p_expectedValue value remains uncanged
        p_expectedOutcome = False
        p_actualOutcome = StringExtensions.TryParseBoolean(p_text, p_actualValue, p_actualDetails)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, _
            "parse of '" & p_text & "' to boolean should " & _
            IIf(p_expectedOutcome, "succeed", "fail") & "; details: '" & p_actualDetails & "'.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_text = "&HA"
        p_expectedValue = True
        p_expectedOutcome = True
        p_actualOutcome = StringExtensions.TryParseBoolean(p_text, p_actualValue, p_actualDetails)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, _
            "parse of '" & p_text & "' to boolean should succeed; " & p_actualDetails & ".")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_text = "0"
        p_expectedValue = False
        p_expectedOutcome = True
        p_actualOutcome = StringExtensions.TryParseBoolean(p_text, p_actualValue, p_actualDetails)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, _
            "parse of '" & p_text & "' to boolean should " & _
            IIf(p_expectedOutcome, "succeed", "fail") & "; details: '" & p_actualDetails & "'.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_text = "off"
        p_expectedValue = False
        p_expectedOutcome = True
        p_actualOutcome = StringExtensions.TryParseBoolean(p_text, p_actualValue, p_actualDetails)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, _
            "parse of '" & p_text & "' to boolean should " & _
            IIf(p_expectedOutcome, "succeed", "fail") & "; details: '" & p_actualDetails & "'.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_text = "false"
        p_expectedValue = False
        p_expectedOutcome = True
        p_actualOutcome = StringExtensions.TryParseBoolean(p_text, p_actualValue, p_actualDetails)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, _
            "parse of '" & p_text & "' to boolean should " & _
            IIf(p_expectedOutcome, "succeed", "fail") & "; details: '" & p_actualDetails & "'.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_text = "on"
        p_expectedValue = True
        p_expectedOutcome = True
        p_actualOutcome = StringExtensions.TryParseBoolean(p_text, p_actualValue, p_actualDetails)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, _
            "parse of '" & p_text & "' to boolean should " & _
            IIf(p_expectedOutcome, "succeed", "fail") & "; details: '" & p_actualDetails & "'.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_text = "true"
        p_expectedValue = True
        p_expectedOutcome = True
        p_actualOutcome = StringExtensions.TryParseBoolean(p_text, p_actualValue, p_actualDetails)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, _
            "parse of '" & p_text & "' to boolean should " & _
            IIf(p_expectedOutcome, "succeed", "fail") & "; details: '" & p_actualDetails & "'.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    ' test how parsing affects error handling
    On Error Resume Next
    
    If p_outcome.AssertSuccessful Then
        p_text = "true"
        p_expectedValue = True
        p_expectedOutcome = True
        p_actualOutcome = StringExtensions.TryParseBoolean(p_text, p_actualValue, p_actualDetails)
    
        ' cause an error and see if it is trapped.
        Dim a_deviceByZeroErrorNumber As Integer
        a_deviceByZeroErrorNumber = 11
        Dim a_notANumber As Double
        a_notANumber = 1& / 0&
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_deviceByZeroErrorNumber, Err.Number, _
            "Error number should match divide by zero error.")
   
    End If
    
    On Error GoTo 0
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestTextShouldParseToBoolean = p_outcome
    
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

''' <summary>   Unit test. Asserts text should parse to Double. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTextShouldParseToDouble() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestTextShouldParseToDouble"
    
    ' Trap errors to the error handler.
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_text As String
    Dim p_expectedValue As Double
    Dim p_expectedOutcome As Boolean
    Dim p_actualOutcome As Boolean
    Dim p_actualValue As Double
    Dim p_actualDetails As String
    
    p_text = "42.42"
    p_expectedValue = 42.42
    p_expectedOutcome = True
    p_actualOutcome = StringExtensions.TryParseDouble(p_text, p_actualValue, p_actualDetails)
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, p_actualDetails)

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_text = "fortytwo"
        ' p_expectedValue value remains unchanged
        p_expectedOutcome = False
        p_actualOutcome = StringExtensions.TryParseDouble(p_text, p_actualValue, p_actualDetails)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, p_actualDetails)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_text = "&HA"
        p_expectedValue = 10
        p_expectedOutcome = True
        p_actualOutcome = StringExtensions.TryParseDouble(p_text, p_actualValue, p_actualDetails)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, p_actualDetails)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    ' test how parsing affects error handling
    On Error Resume Next
    
    If p_outcome.AssertSuccessful Then
        p_text = "&HA"
        p_expectedValue = 10
        p_expectedOutcome = True
        p_actualOutcome = StringExtensions.TryParseDouble(p_text, p_actualValue, p_actualDetails)
    
        ' cause an error and see if it is trapped.
        Dim a_deviceByZeroErrorNumber As Integer
        a_deviceByZeroErrorNumber = 11
        Dim a_notANumber As Double
        a_notANumber = 1& / 0&
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_deviceByZeroErrorNumber, Err.Number, _
            "Error number should match divide by zero error.")
   
    End If
    
    On Error GoTo 0
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestTextShouldParseToDouble = p_outcome

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

''' <summary>   Unit test. Asserts text should parse to Long. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTextShouldParseToLong() As cc_isr_Test_Fx.Assert

    Const p_procedureName As String = "TestTextShouldParseToLong"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_text As String
    Dim p_expectedValue As Long
    Dim p_expectedOutcome As Boolean
    Dim p_actualOutcome As Boolean
    Dim p_actualValue As Long
    Dim p_actualDetails As String
    
    p_text = "42"
    p_expectedValue = 42
    p_expectedOutcome = True
    p_actualOutcome = StringExtensions.TryParseLong(p_text, p_actualValue, p_actualDetails)
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, p_actualDetails)

    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_text = "fortytwo"
        ' p_expectedValue value remains uncanged
        p_expectedOutcome = False
        p_actualOutcome = StringExtensions.TryParseLong(p_text, p_actualValue, p_actualDetails)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, p_actualDetails)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_text = "&HA"
        p_expectedValue = 10
        p_expectedOutcome = True
        p_actualOutcome = StringExtensions.TryParseLong(p_text, p_actualValue, p_actualDetails)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedOutcome, p_actualOutcome, p_actualDetails)
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "parse of '" & p_text & "' should yield the expected value.")
    End If
    
    ' test how parsing affects error handling
    On Error Resume Next
    
    If p_outcome.AssertSuccessful Then
        p_text = "&HA"
        p_expectedValue = 10
        p_expectedOutcome = True
        p_actualOutcome = StringExtensions.TryParseLong(p_text, p_actualValue, p_actualDetails)
    
        ' cause an error and see if it is trapped.
        Dim a_deviceByZeroErrorNumber As Long
        a_deviceByZeroErrorNumber = 11
        Dim a_notANumber As Double
        a_notANumber = 1& / 0&
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(a_deviceByZeroErrorNumber, Err.Number, _
            "Error number should match divide by zero error.")
   
    End If
    
    On Error GoTo 0
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print "Test " & Format(This.TestNumber, "00") & " " & p_outcome.BuildReport(p_procedureName) & _
        " Elapsed time: " & VBA.Format$(This.TestStopper.ElapsedMilliseconds, "0.0") & " ms."
    
    Set TestTextShouldParseToLong = p_outcome

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




