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

    Set TestCharAt = cc_isr_Test_Fx.Assert.areEqual("a", StringExtensions.CharAt("foobar", 5), _
            "Should get the expected character from the string")

End Function

''' <summary>   Unit test. Asserts end width. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestEndsWith() As cc_isr_Test_Fx.Assert

    Set TestEndsWith = cc_isr_Test_Fx.Assert.IsTrue(StringExtensions.EndsWith("foobar", "bar"), _
            "String should end with the expected value")
    
End Function

''' <summary>   Unit test. Asserts escape seqquences existence and values. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestEscapeSequences() As cc_isr_Test_Fx.Assert

    Dim p_escapes As Collection
    Set p_escapes = cc_isr_Core.StringExtensions.EscapeSequences
    Set TestEscapeSequences = cc_isr_Test_Fx.Assert.IsNotNull(p_escapes, _
            "Escape sequences should be created")
            
    If Not TestEscapeSequences.AssertSuccessful Then Exit Function

    Set TestEscapeSequences = cc_isr_Test_Fx.Assert.areEqual(8, p_escapes.count, _
            "Number of escape sequences should match")
            
    If Not TestEscapeSequences.AssertSuccessful Then Exit Function
    
    Dim p_escape As EscapeSequence
    Dim p_item As EscapeSequence
    For Each p_escape In p_escapes
    
        Set p_item = p_escapes(p_escape.value)
        
        Set TestEscapeSequences = cc_isr_Test_Fx.Assert.areEqual(p_escape.value, p_item.value, _
                "For each escape value must match collection item value")
                
        If Not TestEscapeSequences.AssertSuccessful Then Exit For
    
        Set TestEscapeSequences = cc_isr_Test_Fx.Assert.areEqual(p_escape.Literal, p_item.Literal, _
                "For each escape replacement value must match collection item replacement value")
                
        If Not TestEscapeSequences.AssertSuccessful Then Exit For
    
    Next
    
End Function

''' <summary>   Unit test. Asserts inserting. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestInsertRepelaceEscapeSequences() As cc_isr_Test_Fx.Assert

    Dim p_original As String: p_original = "\t1234\r\n"
    Dim p_expected As String: p_expected = VBA.vbTab & "1234" & VBA.Chr$(13) & VBA.Chr$(10)
    Dim p_actual As String: p_actual = cc_isr_Core.StringExtensions.ReplaceEscapeSequences(p_original)

    Dim areEqual As Boolean: areEqual = cc_isr_Core.StringExtensions.AreEqualDebug(p_expected, p_actual)
    
    Set TestInsertRepelaceEscapeSequences = cc_isr_Test_Fx.Assert.areEqual(p_expected, p_actual, _
            "Literal values should be inserted in place of the escaped sequences.")
            
    If TestInsertRepelaceEscapeSequences.AssertSuccessful Then
    
        p_expected = p_original
        p_original = p_actual
        p_actual = cc_isr_Core.StringExtensions.InsertEscapeSequences(p_original)
        Set TestInsertRepelaceEscapeSequences = cc_isr_Test_Fx.Assert.areEqual(p_expected, p_actual, _
                "Escape sequences should be inserted in place of the literal characters.")
    End If

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
    Set TestInsert = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            StringExtensions.Insert(p_original, p_added, p_position), _
            "Added string '" & p_added & "' should be inserted into '" & _
            p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
            
    If TestInsert.AssertSuccessful Then
        p_position = 1: p_expected = "991234": p_suffix = "-st"
        Set TestInsert = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If TestInsert.AssertSuccessful Then
        p_position = 2: p_expected = "199234": p_suffix = "-nd"
        Set TestInsert = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If TestInsert.AssertSuccessful Then
        p_position = 3: p_expected = "129934": p_suffix = "-rd"
        Set TestInsert = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If TestInsert.AssertSuccessful Then
        p_position = 4: p_expected = "123994": p_suffix = "-th"
        Set TestInsert = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If TestInsert.AssertSuccessful Then
        p_position = 100: p_expected = "123499": p_suffix = "-th (after the last)"
        Set TestInsert = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
End Function



''' <summary>   Unit test. Asserts delimited string element should pop. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestPop() As cc_isr_Test_Fx.Assert
    
    Dim p_delimitedString As String: p_delimitedString = "a,b,c"
    
    Set TestPop = cc_isr_Test_Fx.Assert.areEqual("a", _
            StringExtensions.Pop(p_delimitedString, ","), _
            "First element in " & p_delimitedString & " should pop")
            
    If TestPop.AssertSuccessful Then
    
        Set TestPop = cc_isr_Test_Fx.Assert.areEqual("b", _
            StringExtensions.Pop(p_delimitedString, ","), _
            "Second element in " & p_delimitedString & " should pop")
    
    End If
    
    If TestPop.AssertSuccessful Then
    
        Set TestPop = cc_isr_Test_Fx.Assert.areEqual("c", _
            StringExtensions.Pop(p_delimitedString, ","), _
            "Third element in " & p_delimitedString & " should pop")
    End If
    
    If TestPop.AssertSuccessful Then
    
        Set TestPop = cc_isr_Test_Fx.Assert.areEqual(VBA.vbNullString, _
            StringExtensions.Pop(p_delimitedString, ","), _
            "No element in " & p_delimitedString & " should pop")
            
    End If
    
End Function

''' <summary>   Unit test. Asserts creating a repeated string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestRepeat() As cc_isr_Test_Fx.Assert
    
    Set TestRepeat = cc_isr_Test_Fx.Assert.areEqual("aaa", StringExtensions.Repeat("a", 3), _
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

    Dim p_format As String
    Dim p_expected As String
    Dim p_actual As String
    
    p_format = "a{0}{1}"
    p_expected = "abc"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, "b", "c")
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")
            
    If Not TestStringFormat.AssertSuccessful Then Exit Function
    
    p_format = "(C) Currency: {0:C}\n"
    p_expected = "(C) Currency: -123.45$" & VBA.vbLf
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function
    
    p_format = "(C) Currency: {0:C}"
    p_expected = "(C) Currency: -123.00$"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function
    
    p_format = "(D) Decimal:. . . . . . . . . {0:D}"
    p_expected = "(D) Decimal:. . . . . . . . . -123"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function
    
    p_format = "(E) Scientific: . . . . . . . {0:E}"
    p_expected = "(E) Scientific: . . . . . . . -1.23450E2"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    p_format = "(F) Fixed point:. . . . . . . {0:F}"
    p_expected = "(F) Fixed point:. . . . . . . -123.45"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    p_format = "(F) Fixed point:. . . . . . . {0:F1}"
    p_expected = "(F) Fixed point:. . . . . . . -123.5"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    p_format = "(N) Number: . . . . . . . . . {0:N}"
    p_expected = "(N) Number: . . . . . . . . . -123"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    p_format = "(P) Percent:. . . . . . . . . {0:P}"
    p_expected = "(P) Percent:. . . . . . . . . -12,345%"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    p_format = "(R) Round-trip: . . . . . . . {0:R}"
    p_expected = "(R) Round-trip: . . . . . . . -123.45"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, -123.45)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    p_format = "(X) Hexadecimal:. . . . . . . {0:X}"
    p_expected = "(X) Hexadecimal:. . . . . . . 0xFF85"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, CInt(-123))
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    p_format = "(X) Hexadecimal:. . . . . . . {0:x}"
    p_expected = "(X) Hexadecimal:. . . . . . . 0xff85"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, CInt(-123))
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    Dim p_date As Date: p_date = DateValue("January 26, 2013") + TimeValue("8:28:11 PM")

    p_format = "(c) Custom format: . . . . . .{0:cYYYY-MM-DD (MMMM)}"
    p_expected = "(c) Custom format: . . . . . .2013-01-26 (January)"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    p_format = "(d) Short date: . . . . . . . {0:d}"
    p_expected = "(d) Short date: . . . . . . . 1/26/2013"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    p_format = "(D) Long date:. . . . . . . . {0:D}"
    p_expected = "(D) Long date:. . . . . . . . Saturday, January 26, 2013"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    Dim p_time As Date
    p_time = TimeValue("8:28:11 PM")
    
    p_format = "(T) Long time:. . . . . . . . {0:T}"
    p_expected = "(T) Long time:. . . . . . . . 8:28:11 PM"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    p_format = "(f) Full date/short time: . . {0:f}"
    p_expected = "(f) Full date/short time: . . Saturday, January 26, 2013 8:28 PM"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    p_format = "(F) Full date/long time:. . . {0:F}"
    p_expected = "(F) Full date/long time:. . . Saturday, January 26, 2013 8:28:11 PM"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

    p_format = "(s) Sortable: . . . . . . . . {0:s}"
    p_expected = "(s) Sortable: . . . . . . . . 2013-01-26T20:28:11"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, p_date)
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function
    
    ' specify alignment (/padding) and to use escape sequences:
    
    p_format = "\q{0}, {1}!\x20\n'{2,10:C2}'\n'{2,-10:C2}'"
    p_expected = """hello, world! " & VBA.vbLf & "'   100.00$'" & VBA.vbLf & "'100.00$   '"
    p_actual = cc_isr_Core.StringExtensions.StringFormat(p_format, "hello", "world", 100)
    
    Set TestStringFormat = cc_isr_Test_Fx.Assert.areEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")

    If Not TestStringFormat.AssertSuccessful Then Exit Function

End Function

''' <summary>   Unit test. Asserts creating a formatted string using simpel replacement. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStringFormatReplace() As cc_isr_Test_Fx.Assert

    Set TestStringFormatReplace = cc_isr_Test_Fx.Assert.areEqual("aaa", _
            StringExtensions.StringFormat("a{0}{1}", "a", "a"), _
            "Format should build the expected string")

End Function

''' <summary>   Unit test. Asserts finding an item in a string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStringContains() As cc_isr_Test_Fx.Assert

    Dim p_value As String
    Dim p_candidate As String
    p_value = "the string contains"
    p_candidate = "contains"
    Set TestStringContains = cc_isr_Test_Fx.Assert.IsTrue( _
            StringExtensions.StringContains(p_value, p_candidate), _
            "The string '" & p_value & "' should contain '" & p_candidate & "'.")
            
    If Not TestStringContains.AssertSuccessful Then Exit Function
    
    p_value = "the string contains"
    p_candidate = "contained"
    Set TestStringContains = cc_isr_Test_Fx.Assert.IsFalse( _
            StringExtensions.StringContains(p_value, p_candidate), _
            "The string '" & p_value & "' should not contain '" & p_candidate & "'.")

End Function

''' <summary>   Unit test. Asserts finding items in a string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStringContainsAny() As cc_isr_Test_Fx.Assert

    Dim p_value As String
    Dim p_candidate1 As String, p_candidate2 As String
    p_value = "the string contains"
    p_candidate1 = "the"
    p_candidate2 = "contains"
    Set TestStringContainsAny = cc_isr_Test_Fx.Assert.IsTrue( _
            StringExtensions.StringContainsAny(p_value, True, p_candidate1, p_candidate2), _
            "The string '" & p_value & "' should contain '" & p_candidate1 & "' or '" & p_candidate2 & "'.")
            
    If Not TestStringContainsAny.AssertSuccessful Then Exit Function
    
    p_value = "the string contains"
    p_candidate1 = "the"
    p_candidate2 = "contained"
    Set TestStringContainsAny = cc_isr_Test_Fx.Assert.IsTrue( _
            StringExtensions.StringContainsAny(p_value, True, p_candidate1, p_candidate2), _
            "The string '" & p_value & "' should contain '" & p_candidate1 & "' or '" & p_candidate2 & "'.")
    
    p_value = "the string contains"
    p_candidate1 = "not"
    p_candidate2 = "contained"
    Set TestStringContainsAny = cc_isr_Test_Fx.Assert.IsFalse( _
            StringExtensions.StringContainsAny(p_value, True, p_candidate1, p_candidate2), _
            "The string '" & p_value & "' should not contain '" & p_candidate1 & "' or '" & p_candidate2 & "'.")

End Function



''' <summary>   Unit test. Asserts sub-string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestSubstring() As cc_isr_Test_Fx.Assert
    
    Set TestSubstring = cc_isr_Test_Fx.Assert.areEqual("oo", StringExtensions.Substring("foobar", 1, 2), _
            "Should get the expected part of the string")

End Function

''' <summary>   Unit test. Asserts trim left. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTrimLeft() As cc_isr_Test_Fx.Assert
    
    Set TestTrimLeft = cc_isr_Test_Fx.Assert.areEqual("bar", _
        StringExtensions.TrimLeft("oobar", "o"), "String should be left trimmed.")

End Function

''' <summary>   Unit test. Asserts trim right. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTrimRight() As cc_isr_Test_Fx.Assert

    Set TestTrimRight = cc_isr_Test_Fx.Assert.areEqual("f", _
        StringExtensions.TrimRight("foo", "o"), "String should be right-trimmed.")

End Function


