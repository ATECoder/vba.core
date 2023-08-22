Attribute VB_Name = "CoreExtensionTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Core extension methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Asserts that a wait time should be longer or equal to the expected duration. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestWaitShouldEqualOrExceedDuration() As Assert

    Dim p_outcome As Assert
    
    Dim p_expectedDuration As Double
    p_expectedDuration = 0.1
    Dim p_actualDuration As Double: p_actualDuration = cc_isr_Core_IO.CoreExtensions.Wait(p_expectedDuration)
    Set p_outcome = Assert.IsTrue(p_expectedDuration <= p_actualDuration, _
        "Wait time " & CStr(p_actualDuration) & " should be equal ot longer than the specified duration of " & CStr(p_expectedDuration) & " .")
    
End Function

''' <summary>   Unit test. Asserts that default values are as expected. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestDefaultValues() As Assert

    Dim p_outcome As Assert
    
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
Public Function TestParameterArrayPropagated() As Assert
    
    On Error Resume Next
    
    Dim p_outcome As Assert
    
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

