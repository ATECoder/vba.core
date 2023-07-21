Attribute VB_Name = "CoreExtensionTests"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' CoreExtensionsTests.bas
'
' Dependencies:
'
' Assert.cls
' CoreExtensions.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Asserts that the paramter array was passed to the string of routines
''' without errors. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestParameterArrayShoudPassWithoutErrors() As Assert
    
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
    
    Set TestParameterArrayShoudPassWithoutErrors = outcome

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

