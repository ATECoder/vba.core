Attribute VB_Name = "StringExtensionsTests"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' StringExtensionsTests.bas
'
' Dependencies:
'
' Assert.cls
' StringExtensions.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Asserts character at an index position. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCharAt() As cc_isr_Test_Fx.Assert

    Set TestCharAt = cc_isr_Test_Fx.Assert.AreEqual("a", StringExtensions.CharAt("foobar", 5), _
            "Should get the expected character from the string")

End Function

''' <summary>   Unit test. Asserts end width. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestEndsWith() As cc_isr_Test_Fx.Assert

    Set TestEndsWith = cc_isr_Test_Fx.Assert.IsTrue(StringExtensions.EndsWith("foobar", "bar"), _
            "String should end with the expected value")
    
End Function

''' <summary>   Unit test. Asserts inserting. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestInsert() As cc_isr_Test_Fx.Assert
    
    Dim p_original As String: p_original = "1234"
    Dim p_added As String: p_added = "99"
    
    Dim p_expected As String
    Dim p_position As Long
    Dim p_suffix As String
    
    p_position = 0: p_expected = "991234": p_suffix = "-th"
    Set TestInsert = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
            StringExtensions.Insert(p_original, p_added, p_position), _
            "Added string '" & p_added & "' should be inserted into '" & _
            p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
            
    If TestInsert.AssertSuccessful Then
        p_position = 1: p_expected = "991234": p_suffix = "-st"
        Set TestInsert = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If TestInsert.AssertSuccessful Then
        p_position = 2: p_expected = "199234": p_suffix = "-nd"
        Set TestInsert = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If TestInsert.AssertSuccessful Then
        p_position = 3: p_expected = "129934": p_suffix = "-rd"
        Set TestInsert = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If TestInsert.AssertSuccessful Then
        p_position = 4: p_expected = "123994": p_suffix = "-th"
        Set TestInsert = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If TestInsert.AssertSuccessful Then
        p_position = 100: p_expected = "123499": p_suffix = "-th (after the last)"
        Set TestInsert = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
End Function



''' <summary>   Unit test. Asserts delimited string element should pop. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestPop() As cc_isr_Test_Fx.Assert
    
    Dim p_delimitedString As String: p_delimitedString = "a,b,c"
    
    Set TestPop = cc_isr_Test_Fx.Assert.AreEqual("a", _
            StringExtensions.Pop(p_delimitedString, ","), _
            "First element in " & p_delimitedString & " should pop")
            
    If TestPop.AssertSuccessful Then
    
        Set TestPop = cc_isr_Test_Fx.Assert.AreEqual("b", _
            StringExtensions.Pop(p_delimitedString, ","), _
            "Second element in " & p_delimitedString & " should pop")
    
    End If
    
    If TestPop.AssertSuccessful Then
    
        Set TestPop = cc_isr_Test_Fx.Assert.AreEqual("c", _
            StringExtensions.Pop(p_delimitedString, ","), _
            "Third element in " & p_delimitedString & " should pop")
    End If
    
    If TestPop.AssertSuccessful Then
    
        Set TestPop = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, _
            StringExtensions.Pop(p_delimitedString, ","), _
            "No element in " & p_delimitedString & " should pop")
            
    End If
    
End Function

''' <summary>   Unit test. Asserts creating a repeated string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestRepeat() As cc_isr_Test_Fx.Assert
    
    Set TestRepeat = cc_isr_Test_Fx.Assert.AreEqual("aaa", StringExtensions.Repeat("a", 3), _
            "Should constract a string with repreated strings")
            
End Function

''' <summary>   Unit test. Asserts start with. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStartsWith() As cc_isr_Test_Fx.Assert

    Set TestStartsWith = cc_isr_Test_Fx.Assert.IsTrue(StringExtensions.StartsWith("foobar", "foo"), _
            "String should start with the expected value.")
            
End Function

''' <summary>   Unit test. Asserts creating a formatted string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStringFormat() As cc_isr_Test_Fx.Assert

    Set TestStringFormat = cc_isr_Test_Fx.Assert.AreEqual("aaa", _
            StringExtensions.StringFormat("a{0}{1}", "a", "a"), _
            "Format should build the expected string")

End Function

''' <summary>   Unit test. Asserts sub-string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestSubstring() As cc_isr_Test_Fx.Assert
    
    Set TestSubstring = cc_isr_Test_Fx.Assert.AreEqual("oo", StringExtensions.Substring("foobar", 1, 2), _
            "Should get the expected part of the string")

End Function

''' <summary>   Unit test. Asserts trim left. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTrimLeft() As cc_isr_Test_Fx.Assert
    
    Set TestTrimLeft = cc_isr_Test_Fx.Assert.AreEqual("bar", _
        StringExtensions.TrimLeft("oobar", "o"), "String should be left trimmed.")

End Function

''' <summary>   Unit test. Asserts trim right. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTrimRight() As cc_isr_Test_Fx.Assert

    Set TestTrimRight = cc_isr_Test_Fx.Assert.AreEqual("f", _
        StringExtensions.TrimRight("foo", "o"), "String should be right-trimmed.")

End Function


