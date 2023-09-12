Attribute VB_Name = "CoreExtensionTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Core extension methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Name As String
    TestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
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
    'BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestWaitShouldEqualOrExceedDuration
        Case 2
            Set p_outcome = TestNowResolution
        Case 3
            Set p_outcome = TestDefaultValues
        Case 4
            Set p_outcome = TestParameterArrayPropagated
        Case Else
    End Select
    Set RunTest = p_outcome
    'AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    'BeforeAll
    RunTest 2
    'AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    This.Name = "CoreExtensionTests"
    'BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 4
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
    'AfterAll
    Debug.Print "Ran " & VBA.CStr(This.RunCount) & " out of " & VBA.CStr(This.TestCount) & " tests."
    Debug.Print "Passed: " & VBA.CStr(This.PassedCount) & "; Failed: " & VBA.CStr(This.FailedCount) & _
                "; Inconclusive: " & VBA.CStr(This.InconclusiveCount) & "."
End Sub


''' <summary>   Unit test. Asserts that a wait time should be longer or equal to the expected duration. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestWaitShouldEqualOrExceedDuration() As cc_isr_Test_Fx.Assert

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
    
    Debug.Print p_outcome.BuildReport("TestWaitShouldEqualOrExceedDuration")

    Set TestWaitShouldEqualOrExceedDuration = p_outcome
    
End Function

''' <summary>   Unit test. Asserts that the resolution of VBA.Now() should be longer or equal
'''             to the expected resolution but smaller than double of that resoltion. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestNowResolution() As cc_isr_Test_Fx.Assert

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
    
    Debug.Print p_outcome.BuildReport("TestNowResolution")
    
    Set TestNowResolution = p_outcome
    
End Function

''' <summary>   Unit test. Asserts that default values are as expected. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestDefaultValues() As cc_isr_Test_Fx.Assert

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
    
    Debug.Print p_outcome.BuildReport("TestDefaultValues")

    Set TestDefaultValues = p_outcome

End Function

''' <summary>   Unit test. Asserts that the paramter array propagated through nested methods
''' without errors. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestParameterArrayPropagated() As cc_isr_Test_Fx.Assert
    
    On Error Resume Next
    
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
    
    On Error GoTo 0
    
    Debug.Print p_outcome.BuildReport("TestDefaultValues")
    
    Set TestParameterArrayPropagated = p_outcome

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

