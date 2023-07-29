Attribute VB_Name = "CoreExtensionTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Core extension methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Asserts that default values are as expected. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestDefaultValues() As Assert

    Dim outcome As Assert
    
    Set outcome = Assert.AreEqual(False, cc_isr_Core.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbBoolean), _
        "The default value of VBA.VbVarType.vbBoolean should equal.")
    
    If outcome.AssertSuccessful Then _
        Set outcome = Assert.AreEqual(0, cc_isr_Core.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbByte), _
            "The default value of VBA.VbVarType.vbByte should equal.")
    
    If outcome.AssertSuccessful Then _
        Set outcome = Assert.AreEqual(Empty, cc_isr_Core.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbArray), _
            "The default value of VBA.VbVarType.vbArray should equal.")
    
    If outcome.AssertSuccessful Then _
        Set outcome = Assert.AreEqual(Null, cc_isr_Core.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbNull), _
            "The default value of VBA.VbVarType.vbNull should equal.")
    
    If outcome.AssertSuccessful Then _
        Set outcome = Assert.IsNull(cc_isr_Core.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbNull), _
            "The default value of VBA.VbVarType.vbNull should be Null.")
    
    If outcome.AssertSuccessful Then _
        Set outcome = Assert.IsTrue(cc_isr_Core.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbObject) Is Nothing, _
            "The default value of VBA.VbVarType.vbObject should be nothing.")
    
    If outcome.AssertSuccessful Then _
        Set outcome = Assert.AreEqual(0, cc_isr_Core.CoreExtensions.GetDefaultValue(VBA.VbVarType.vbLongLong), _
            "The default value of VBA.VbVarType.vbLongLong should equal.")
    
    If outcome.AssertSuccessful Then
        Debug.Print "TestDefaultValues passed"
    Else
        Debug.Print "Test failed: " & outcome.AssertMessage
    End If

    Set TestDefaultValues = outcome

End Function

''' <summary>   Unit test. Asserts that the paramter array propagated through nested methods
''' without errors. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestParameterArrayPropagated() As Assert
    
    On Error Resume Next
    
    Dim outcome As Assert
    
    Dim p_dummyVariant As Variant
    p_dummyVariant = "a"
    Dim p_dummyArray() As Variant
    Dim p_unboxedTokens As Variant
    
    Dim p_tokens As Variant
    p_tokens = Method1("a", "b", "c")
    
    Set outcome = Assert.AreEqual(VBA.Err.Number, 0, _
        "The parameter array should pass without errors")
        
    If outcome.AssertSuccessful Then
    
        Set outcome = Assert.AreEqual(TypeName(p_dummyArray), TypeName(p_tokens), _
        "The nested parameter array type should match the expected type")
    
    End If
        
    If outcome.AssertSuccessful Then
    
        Set outcome = Assert.AreEqual(TypeName(p_dummyArray), TypeName(p_tokens(0)), _
        "The first element of the nested parameter array type should match the expected type")
    
    End If
        
    If outcome.AssertSuccessful Then
    
        p_unboxedTokens = CoreExtensions.UnboxParameterArray(p_tokens)
        
        Set outcome = Assert.AreEqual(TypeName(p_dummyArray), TypeName(p_unboxedTokens), _
        "The unboxed parameter array type should match the expected type")
        
    End If
    
    If outcome.AssertSuccessful Then
    
        Set outcome = Assert.AreEqual(TypeName(p_dummyVariant), TypeName(p_unboxedTokens(0)), _
        "The first element of the nested parameter array type should match the expected type")
    
    End If
    
    On Error GoTo 0
    
    Set TestParameterArrayPropagated = outcome

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

