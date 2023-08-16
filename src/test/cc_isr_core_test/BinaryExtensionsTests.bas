Attribute VB_Name = "BinaryExtensionsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Binary extension methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Binary bits should invert. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestBitsShouldInvert() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_initialValue As String
    Dim p_expectedValue As String
    Dim p_actualValue As String
    
    p_initialValue = "101"
    p_expectedValue = "010"
    p_actualValue = cc_isr_core.BinaryExtensions.InvertBits(p_initialValue)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "Actual value should equal the inverted bits of '" & p_initialValue & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = "0"
        p_expectedValue = "1"
        p_actualValue = cc_isr_core.BinaryExtensions.InvertBits(p_initialValue)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the inverted bits of '" & p_initialValue & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = "00"
        p_expectedValue = "11"
        p_actualValue = cc_isr_core.BinaryExtensions.InvertBits(p_initialValue)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the inverted bits of '" & p_initialValue & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = "11"
        p_expectedValue = "00"
        p_actualValue = cc_isr_core.BinaryExtensions.InvertBits(p_initialValue)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the inverted bits of '" & p_initialValue & "'.")
    
    End If
    
    Debug.Print p_outcome.BuildReport("TestBitsShouldInvert")
        
    Set TestBitsShouldInvert = p_outcome
     
End Function

''' <summary>   Unit test. Binary values should add. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestBinaryValuesShouldAdd() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_initialValue1 As String
    Dim p_initialValue2 As String
    Dim p_expectedValue As String
    Dim p_actualValue As String
    
    p_initialValue1 = "1"
    p_initialValue2 = "0"
    p_expectedValue = "1"
    p_actualValue = cc_isr_core.BinaryExtensions.AddBinary(p_initialValue1, p_initialValue2)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "Actual value should equal the sum of '" & p_initialValue1 & "' and '" & p_initialValue2 & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue1 = "10"
        p_initialValue2 = "0"
        p_expectedValue = "10"
        p_actualValue = cc_isr_core.BinaryExtensions.AddBinary(p_initialValue1, p_initialValue2)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the sum of '" & p_initialValue1 & "' and '" & p_initialValue2 & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue1 = "10"
        p_initialValue2 = "10"
        p_expectedValue = "100"
        p_actualValue = cc_isr_core.BinaryExtensions.AddBinary(p_initialValue1, p_initialValue2)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the sum of '" & p_initialValue1 & "' and '" & p_initialValue2 & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue1 = "11"
        p_initialValue2 = "10"
        p_expectedValue = "101"
        p_actualValue = cc_isr_core.BinaryExtensions.AddBinary(p_initialValue1, p_initialValue2)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the sum of '" & p_initialValue1 & "' and '" & p_initialValue2 & "'.")
    
    End If
    
    Debug.Print p_outcome.BuildReport("TestBinaryValuesShouldAdd")
        
    Set TestBinaryValuesShouldAdd = p_outcome
     
End Function

''' <summary>   Unit test. Long decimal value should convert to binary. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestLongValueShouldConvertToBinary() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_initialValue As Integer
    Dim p_expectedValue As String
    Dim p_actualValue As String
    
    p_initialValue = 0
    p_expectedValue = "0000"
    p_actualValue = cc_isr_core.BinaryExtensions.LongToBinary(p_initialValue, 4)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 1
        p_expectedValue = "0001"
        p_actualValue = cc_isr_core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 5
        p_expectedValue = "0101"
        p_actualValue = cc_isr_core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -1
        p_expectedValue = 1111
        p_actualValue = cc_isr_core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -2
        p_expectedValue = 1110
        p_actualValue = cc_isr_core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -3
        p_expectedValue = 1101
        p_actualValue = cc_isr_core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    Debug.Print p_outcome.BuildReport("TestLongValueShouldConvertToBinary")
        
    Set TestLongValueShouldConvertToBinary = p_outcome
     
End Function

''' <summary>   Unit test. Fractional decimal value should convert to binary. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestFractionalValueShouldConvertToBinary() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_initialValue As Double
    Dim p_expectedValue As String
    Dim p_actualValue As String
    
    p_initialValue = 0
    p_expectedValue = "0000"
    p_actualValue = cc_isr_core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 0.5
        p_expectedValue = "1000"
        p_actualValue = cc_isr_core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 0.25
        p_expectedValue = "0100"
        p_actualValue = cc_isr_core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 0.75
        p_expectedValue = 1100
        p_actualValue = cc_isr_core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        On Error Resume Next
        p_initialValue = -0.75
        p_actualValue = cc_isr_core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(UserDefinedErrors.InvalidArgumentError.Code, _
                Err.Number, _
                "Attempting to convert a negative fractional should raise the Invalid Argument Error code..")
        On Error GoTo 0
    
    End If
    
    Debug.Print p_outcome.BuildReport("TestFractionalValueShouldConvertToBinary")

    Set TestFractionalValueShouldConvertToBinary = p_outcome
     
End Function

''' <summary>   Unit test. Double decimal value should convert to binary. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestDoubleValueShouldConvertToBinary() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_initialValue As Double
    Dim p_expectedValue As String
    Dim p_actualValue As String
    
    p_initialValue = 0
    p_expectedValue = "0000.0000"
    p_actualValue = cc_isr_core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
            "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 1.5
        p_expectedValue = "0001.1000"
        p_actualValue = cc_isr_core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 1.25
        p_expectedValue = "0001.0100"
        p_actualValue = cc_isr_core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -1#
        p_expectedValue = "1111.0000"
        p_actualValue = cc_isr_core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -1.5
        p_expectedValue = "1110.1000"
        p_actualValue = cc_isr_core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 12.5
        p_expectedValue = "01100.10000"
        p_actualValue = cc_isr_core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 12.25
        p_expectedValue = "01100.01000"
        p_actualValue = cc_isr_core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 12.75
        p_expectedValue = "01100.11000"
        p_actualValue = cc_isr_core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -12.5
        p_expectedValue = "10011.10000"
        p_actualValue = cc_isr_core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -12.25
        p_expectedValue = "10011.11000"
        p_actualValue = cc_isr_core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then

        p_initialValue = -12.75
        p_expectedValue = "10011.01000"
        p_actualValue = cc_isr_core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")

    End If
    
    Debug.Print p_outcome.BuildReport("TestLongValueShouldConvertToBinary")

    Set TestDoubleValueShouldConvertToBinary = p_outcome

End Function


