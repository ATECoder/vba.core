Attribute VB_Name = "AssertTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Test assertion tests.  </summary>
''' <remarks>   Dependencies: Assert.cls.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Private Type this_
    Name As String
    TestNumber As Integer
    BeforeAllAssert As Assert
    BeforeEachAssert As Assert
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
    BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestAssertingInconclusiveShouldReportInconclusive
        Case 2
            Set p_outcome = TestAssertingFailShouldReportFailure
        Case 3
            Set p_outcome = TestAssertingPassShouldReportPass
        Case 4
            Set p_outcome = TestNothingShouldBeAsserted
        Case 5
            Set p_outcome = TestNothingShouldNotBeAsserted
        Case 6
            Set p_outcome = TestNullShouldBeAsserted
        Case 7
            Set p_outcome = TestNullShouldNotBeAsserted
        Case 8
            Set p_outcome = TestAssertingSamenessShouldReportSameness
        Case 9
            Set p_outcome = TestAssertingNonSamenessShouldReportNonSameness
        Case 10
            Set p_outcome = TestStringEqualityShouldWork
        Case 11
            Set p_outcome = TestShouldAssertCloseDoubleValues
        Case 12
            Set p_outcome = TestShouldAssertCloseSingleValues
        Case Else
    End Select
    Set RunTest = p_outcome
    AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 12
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 12
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

    This.Name = "AssetTests"
    
    Set This.ErrTracer = New ErrTracer
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState

    ' Prime all tests

    This.TestNumber = 0
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful And Not This.ErrTracer Is Nothing Then _
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

    This.TestNumber = This.TestNumber + 1

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

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful And Not This.ErrTracer Is Nothing Then _
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

    ' cleanup after each test.
    If This.BeforeEachAssert.AssertSuccessful Then
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before Each' assert.
    Set This.BeforeEachAssert = Nothing

    ' report any leftover errors.
    If Not This.ErrTracer Is Nothing Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
    Else
        Set p_outcome = Assert.Inconclusive("Errors reported cleaning up test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & p_outcome.AssertMessage)
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
    If Not This.ErrTracer Is Nothing Then _
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


''' <summary>   Unit test. Asserting <see cref="Assert.Inconclusive"/> should report Inconclusive. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertInconclusive"/> True. </returns>
Public Function TestAssertingInconclusiveShouldReportInconclusive() As Assert

    Dim p_assert As Assert

    Dim p_outcome As Assert
    
    Set p_assert = Assert.Inconclusive("Asserting Inconclusive to test inconclusive outcome.")
    
    Set p_outcome = Assert.IsTrue(p_assert.AssertInconclusive, "Asserting inconclusive should report AssertInconclusive as True.")
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsFalse(p_assert.AssertSuccessful, "Asserting inconclusive should report AssertSuccessful as False.")
    
    End If
    
    Debug.Print p_outcome.BuildReport("TestAssertingInconclusiveShouldReportInconclusive")
    
    Set TestAssertingInconclusiveShouldReportInconclusive = p_outcome
    
End Function

''' <summary>   Unit test. Asserting <see cref="Assert.Fail"/> should report failure. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestAssertingFailShouldReportFailure() As Assert

    Dim p_assert As Assert
    Dim p_outcome As Assert
    
    Set p_assert = Assert.Fail("Asserting Fail to test failure outcome.")
    
    Set p_outcome = Assert.IsFalse(p_assert.AssertSuccessful, "Asserting failure should report AssertSuccessful as false.")
    
    Debug.Print p_outcome.BuildReport("TestAssertingFailShouldReportFailure")
    
    Set TestAssertingFailShouldReportFailure = p_outcome
    
End Function

''' <summary>   Unit test. Asserting <see cref="Assert.Pass"/> should report pass. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestAssertingPassShouldReportPass() As Assert

    Dim p_assert As Assert
    Dim p_outcome As Assert
    
    Set p_assert = Assert.Pass("Asserting Pass to test Pass outcome.")
    
    Set p_outcome = Assert.IsTrue(p_assert.AssertSuccessful, "Asserting Pass should report AssertSuccessful as True.")
    
    Debug.Print p_outcome.BuildReport("TestAssertingPassShouldReportPass")
    
    Set TestAssertingPassShouldReportPass = p_outcome
    
End Function

''' <summary>   Unit test. Asserting nothing should assert nothing. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestNothingShouldBeAsserted() As Assert

    Dim p_object As Object
    Set p_object = Nothing
    
    Dim p_outcome As Assert
    
    Set p_outcome = Assert.IsNothing(p_object, "Object should be noting.")
    
    Debug.Print p_outcome.BuildReport("TestNothingShouldBeAsserted")
    
    Set TestNothingShouldBeAsserted = p_outcome
    
End Function

''' <summary>   Unit test. Asserting not nothing should not assert nothing. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestNothingShouldNotBeAsserted() As Assert

    Dim p_object As Object
    Set p_object = Assert
    
    Dim p_outcome As Assert
    
    Set p_outcome = Assert.IsNotNothing(p_object, "Object should be not be noting.")
    
    Debug.Print p_outcome.BuildReport("TestNothingShouldNotBeAsserted")
    
    Set TestNothingShouldNotBeAsserted = p_outcome
    
End Function

''' <summary>   Unit test. Asserting Null should assert Null. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestNullShouldBeAsserted() As Assert

    Dim p_object As Object
    Dim p_value As Integer
    Dim p_variant As Variant
    
    Dim p_outcome As Assert
    
    Set p_outcome = Assert.IsNull(p_value, "Integer value should be Null (Not IsObject()).")
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNull(p_variant, "Unset Variant should be Null (Not IsObject()).")
    
    End If
        
    p_variant = CInt(0)
     
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNull(p_variant, "Variant set to integer should be Null (Not IsObject()).")
    
    End If
     
    p_variant = "a"
     
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNull(p_variant, "Variant set to a string should be Null (Not IsObject()).")
    
    End If
     
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNull(Empty, "'Empty' should be Null (Not IsObject()).")
    
    End If
        
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNull(Null, "'Null' should be Null (Not IsObject()).")
    
    End If
    
    Debug.Print p_outcome.BuildReport("TestNullShouldBeAsserted")
    
    Set TestNullShouldBeAsserted = p_outcome
    
End Function

''' <summary>   Unit test. Asserting not Null should not assert Null. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestNullShouldNotBeAsserted() As Assert

    Dim p_object As Object
    Dim p_variant As Variant
    
    Dim p_outcome As Assert
    
    Set p_outcome = Assert.IsNotNull(p_object, "Object should be not be Null (IsObject()).")
    
    Debug.Print "TestNullShouldNotBeAsserted " & _
        IIf(p_outcome.AssertSuccessful, "passed.", "failed: " & p_outcome.AssertMessage)
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsTrue(VBA.IsObject(Nothing), "IsObject(Nothing) should be true.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsFalse(VBA.IsNull(Nothing), "IsNull(Nothing) should be false.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsTrue(VBA.IsObject(Nothing), "VBA.IsObject(Nothing) should be true.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNotNull(Nothing, "'Nothing' should not be Null (IsObject()).")
    
    End If
    
    Set p_variant = p_object
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNotNull(Nothing, "Variant set to an object should not be Null (IsObject()).")
    
    End If
    
    Debug.Print p_outcome.BuildReport("TestNullShouldNotBeAsserted")
    
    Set TestNullShouldNotBeAsserted = p_outcome
    
End Function

''' <summary>   Unit test. Asserting <see cref="Assert.Same"/> should report success if
'''             objects are the same and failure if objects are not the same. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestAssertingSamenessShouldReportSameness() As Assert

    Dim p_outcome As Assert
    
    Dim p_object1 As Object
    Dim p_object2 As Variant
    
    ' set the object to the testing sheet.
    Set p_object1 = cc_isr_Test_Fx.UnitTestSheet
    Set p_object2 = p_object1
    
    Set p_outcome = Assert.AreSame(p_object1, p_object2, "The objects should be the same.")
    Set p_outcome = Assert.IsTrue(p_outcome.AssertSuccessful, _
            "Asserting sameness on the same objects should report AssertSuccessful as True.")
    
    If p_outcome.AssertSuccessful Then
    
        Set p_object2 = Nothing
        Set p_outcome = Assert.AreSame(p_object1, p_object2, "The objects should not be the same.")
        Set p_outcome = Assert.IsFalse(p_outcome.AssertSuccessful, _
                "Asserting sameness on different objects should report AssertSuccessful as False.")
    
    End If
    
    Debug.Print p_outcome.BuildReport("TestAssertingSamenessShouldReportSameness")
    
    Set TestAssertingSamenessShouldReportSameness = p_outcome
    
End Function

''' <summary>   Unit test. Asserting <see cref="Assert.NotSame"/> should report success if
'''             objects are not the same and failure if objects are the same. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestAssertingNonSamenessShouldReportNonSameness() As Assert

    Dim p_outcome As Assert
    
    Dim p_object1 As Object
    Dim p_object2 As Variant
    
    ' set the object tothe testing sheet.
    Set p_object1 = cc_isr_Test_Fx.UnitTestSheet
    Set p_object2 = p_object1
    
    Set p_outcome = Assert.AreNotSame(p_object1, p_object2, "The objects should be the same.")
    Set p_outcome = Assert.IsFalse(p_outcome.AssertSuccessful, _
            "Asserting non sameness on the same objects should report AssertSuccessful as False.")
    
    If p_outcome.AssertSuccessful Then
    
        Set p_object2 = Nothing
        Set p_outcome = Assert.AreNotSame(p_object1, p_object2, "The objects should not be the same.")
        Set p_outcome = Assert.IsTrue(p_outcome.AssertSuccessful, _
                "Asserting non sameness on different objects should report AssertSuccessful as True.")
    
    End If
   
    Debug.Print p_outcome.BuildReport("TestAssertingNonSamenessShouldReportNonSameness")
    
    Set TestAssertingNonSamenessShouldReportNonSameness = p_outcome
    
End Function

''' <summary>   Unit test. String equality should work. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestStringEqualityShouldWork() As Assert

    Dim p_outcome As Assert
    
    Dim p_expected As String
    Dim p_actual As String
    
    p_expected = "ALL CAPS"
    p_actual = "ALL CAPS"
    Set p_outcome = Assert.AreEqualString(p_expected, p_actual, VBA.VbCompareMethod.vbBinaryCompare, _
        "The two strings should equal using binary compare.")
        
    If p_outcome.AssertSuccessful Then
    
        p_expected = "ALL CAPS"
        p_actual = "all caps"
        Set p_outcome = Assert.AreNotEqualString(p_expected, p_actual, VBA.VbCompareMethod.vbBinaryCompare, _
            "The two string should not equal using binary compare.")
    
    End If
   
    If p_outcome.AssertSuccessful Then
    
        p_expected = "ALL CAPS"
        p_actual = "all caps"
        Set p_outcome = Assert.AreEqualString(p_expected, p_actual, VBA.VbCompareMethod.vbTextCompare, _
            "The two string should equal using text compare.")
    
    End If
   
    Debug.Print p_outcome.BuildReport("TestStringEqualityShouldWork")
    
    Set TestStringEqualityShouldWork = p_outcome
    
End Function

''' <summary>   Unit test. Asserts if double values are close or not. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestShouldAssertCloseDoubleValues() As Assert

    Const p_procedureName As String = "TestShouldAssertCloseDoubleValues"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As Assert
    
    Dim p_expectedValue As Double
    Dim p_actualValue As Double
    Dim p_epsilon As Double
    
    Set p_outcome = Assert.Pass("entered " & p_procedureName)
    
    If p_outcome.AssertSuccessful Then
    
        p_expectedValue = 10.1
        p_epsilon = 0.05
        ' this fails if not reducing the difference
        p_actualValue = p_expectedValue + 0.999 * p_epsilon
        Set p_outcome = Assert.AreCloseDouble(p_expectedValue, p_actualValue, p_epsilon, _
            "Values should be within Epsilon of each other.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_expectedValue = 10.1
        p_epsilon = 0.05
        p_actualValue = p_expectedValue + 1.0001 * p_epsilon
        Set p_outcome = Assert.AreNotCloseDouble(p_expectedValue, p_actualValue, p_epsilon, _
            "Values should not be within Epsilon of each other.")
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful And Not This.ErrTracer Is Nothing Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestShouldAssertCloseDoubleValues")
    
    Set TestShouldAssertCloseDoubleValues = p_outcome
    
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

''' <summary>   Unit test. Asserts if Single values are close or not. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestShouldAssertCloseSingleValues() As Assert

    Const p_procedureName As String = "TestShouldAssertCloseSingleValues"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As Assert
    
    Dim p_expectedValue As Single
    Dim p_actualValue As Single
    Dim p_epsilon As Single
    
    Set p_outcome = Assert.Pass("entered " & p_procedureName)
    
    If p_outcome.AssertSuccessful Then
    
        p_expectedValue = 10.1
        p_epsilon = 0.05
        ' this fails if not reducing the difference
        p_actualValue = p_expectedValue + 0.999 * p_epsilon
        Set p_outcome = Assert.AreCloseSingle(p_expectedValue, p_actualValue, p_epsilon, _
            "Values should be within Epsilon of each other.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_expectedValue = 10.1
        p_epsilon = 0.05
        p_actualValue = p_expectedValue + 1.0001 * p_epsilon
        Set p_outcome = Assert.AreNotCloseSingle(p_expectedValue, p_actualValue, p_epsilon, _
            "Values should not be within Epsilon of each other.")
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful And Not This.ErrTracer Is Nothing Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestShouldAssertCloseSingleValues")
    
    Set TestShouldAssertCloseSingleValues = p_outcome
    
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




