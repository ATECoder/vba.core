Attribute VB_Name = "BinaryExtensionsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Binary extension methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Binary bits should invert. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestBitsShouldInvert() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_initialValue As String
    Dim p_expectedValue As String
    Dim p_actualValue As String
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("a", StringExtensions.CharAt("foobar", 5), _
            "Should get the expected character from the string")

    Debug.Print "TestBitsShouldInvert " & _
        IIf(p_outcome.AssertSuccessful, "passed.", "failed: " & p_outcome.AssertMessage)
End Function



