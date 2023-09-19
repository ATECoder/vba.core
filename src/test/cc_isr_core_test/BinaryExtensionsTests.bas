Attribute VB_Name = "BinaryExtensionsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Binary extension methods. </summary>
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
    'AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    'BeforeAll
    RunTest 1
    'AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    This.Name = "BinaryExtensionTests"
    'BeforeAll
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
    'AfterAll
    Debug.Print "Ran " & VBA.CStr(This.RunCount) & " out of " & VBA.CStr(This.TestCount) & " tests."
    Debug.Print "Passed: " & VBA.CStr(This.PassedCount) & "; Failed: " & VBA.CStr(This.FailedCount) & _
                "; Inconclusive: " & VBA.CStr(This.InconclusiveCount) & "."
End Sub


''' <summary>   Unit test. Binary bits should invert. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestBitsShouldInvert() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_initialValue As String
    Dim p_expectedValue As String
    Dim p_actualValue As String
    
    p_initialValue = "101"
    p_expectedValue = "010"
    p_actualValue = cc_isr_Core.BinaryExtensions.InvertBits(p_initialValue)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
            "Actual value should equal the inverted bits of '" & p_initialValue & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = "0"
        p_expectedValue = "1"
        p_actualValue = cc_isr_Core.BinaryExtensions.InvertBits(p_initialValue)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the inverted bits of '" & p_initialValue & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = "00"
        p_expectedValue = "11"
        p_actualValue = cc_isr_Core.BinaryExtensions.InvertBits(p_initialValue)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the inverted bits of '" & p_initialValue & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = "11"
        p_expectedValue = "00"
        p_actualValue = cc_isr_Core.BinaryExtensions.InvertBits(p_initialValue)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
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
    p_actualValue = cc_isr_Core.BinaryExtensions.AddBinary(p_initialValue1, p_initialValue2)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
            "Actual value should equal the sum of '" & p_initialValue1 & "' and '" & p_initialValue2 & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue1 = "10"
        p_initialValue2 = "0"
        p_expectedValue = "10"
        p_actualValue = cc_isr_Core.BinaryExtensions.AddBinary(p_initialValue1, p_initialValue2)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the sum of '" & p_initialValue1 & "' and '" & p_initialValue2 & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue1 = "10"
        p_initialValue2 = "10"
        p_expectedValue = "100"
        p_actualValue = cc_isr_Core.BinaryExtensions.AddBinary(p_initialValue1, p_initialValue2)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should equal the sum of '" & p_initialValue1 & "' and '" & p_initialValue2 & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue1 = "11"
        p_initialValue2 = "10"
        p_expectedValue = "101"
        p_actualValue = cc_isr_Core.BinaryExtensions.AddBinary(p_initialValue1, p_initialValue2)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
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
    p_actualValue = cc_isr_Core.BinaryExtensions.LongToBinary(p_initialValue, 4)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
            "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 1
        p_expectedValue = "0001"
        p_actualValue = cc_isr_Core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 5
        p_expectedValue = "0101"
        p_actualValue = cc_isr_Core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -1
        p_expectedValue = 1111
        p_actualValue = cc_isr_Core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -2
        p_expectedValue = 1110
        p_actualValue = cc_isr_Core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -3
        p_expectedValue = 1101
        p_actualValue = cc_isr_Core.BinaryExtensions.LongToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
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
    p_actualValue = cc_isr_Core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
            "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 0.5
        p_expectedValue = "1000"
        p_actualValue = cc_isr_Core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 0.25
        p_expectedValue = "0100"
        p_actualValue = cc_isr_Core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 0.75
        p_expectedValue = 1100
        p_actualValue = cc_isr_Core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        On Error Resume Next
        p_initialValue = -0.75
        p_actualValue = cc_isr_Core.BinaryExtensions.FractionalToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(UserDefinedErrors.InvalidArgumentError.Number, _
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
    p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
            "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 1.5
        p_expectedValue = "0001.1000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 1.25
        p_expectedValue = "0001.0100"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If

    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -1#
        p_expectedValue = "1111.0000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -1.5
        p_expectedValue = "1110.1000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 4)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 12.5
        p_expectedValue = "01100.10000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 12.25
        p_expectedValue = "01100.01000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = 12.75
        p_expectedValue = "01100.11000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -12.5
        p_expectedValue = "10011.10000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        p_initialValue = -12.25
        p_expectedValue = "10011.11000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")
    
    End If
    
    If p_outcome.AssertSuccessful Then

        p_initialValue = -12.75
        p_expectedValue = "10011.01000"
        p_actualValue = cc_isr_Core.BinaryExtensions.DoubleToBinary(p_initialValue, 5)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expectedValue, p_actualValue, _
                "Actual value should be the binary value of '" & VBA.CStr(p_initialValue) & "'.")

    End If
    
    Debug.Print p_outcome.BuildReport("TestLongValueShouldConvertToBinary")

    Set TestDoubleValueShouldConvertToBinary = p_outcome

End Function


