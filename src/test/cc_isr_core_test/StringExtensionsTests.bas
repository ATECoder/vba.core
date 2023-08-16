Attribute VB_Name = "StringExtensionsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. String extension methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Asserts character at an index position. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCharAt() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("a", StringExtensions.CharAt("foobar", 5), _
            "Should get the expected character from the string")

    Debug.Print p_outcome.BuildReport("TestCharAt")
    
    Set TestCharAt = p_outcome

End Function

''' <summary>   Unit test. Asserts end width. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestEndsWith() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(StringExtensions.EndsWith("foobar", "bar"), _
            "String should end with the expected value")
    
    Debug.Print p_outcome.BuildReport("TestEndsWith")
    
    Set TestEndsWith = p_outcome

End Function

''' <summary>   Unit test. Asserts escape sequences existence and values. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestEscapeSequences() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_escapes As Collection
    Set p_escapes = cc_isr_core.StringExtensions.EscapeSequences
    Set p_outcome = cc_isr_Test_Fx.Assert.IsNotNull(p_escapes, _
            "Escape sequences should be created")
            
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(8, p_escapes.count, _
                "Number of escape sequences should match")
                
    End If

    If p_outcome.AssertSuccessful Then
    
        Dim p_escape As EscapeSequence
        Dim p_item As EscapeSequence
        For Each p_escape In p_escapes
        
            Set p_item = p_escapes(p_escape.value)
            
            Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_escape.value, p_item.value, _
                    "For each escape value must match collection item value")
                    
            If Not p_outcome.AssertSuccessful Then Exit For
        
            Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_escape.Literal, p_item.Literal, _
                    "For each escape replacement value must match collection item replacement value")
                    
            If Not p_outcome.AssertSuccessful Then Exit For
        
        Next
        
    End If
    
    Debug.Print p_outcome.BuildReport("TestEscapeSequences")
    
    Set TestEscapeSequences = p_outcome

End Function

''' <summary>   Unit test. Asserts inserting. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestInsertRepelaceEscapeSequences() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_original As String: p_original = "\t1234\r\n"
    Dim p_expected As String: p_expected = VBA.vbTab & "1234" & VBA.Chr$(13) & VBA.Chr$(10)
    Dim p_actual As String: p_actual = cc_isr_core.StringExtensions.ReplaceEscapeSequences(p_original)

    Dim p_areEqual As Boolean: p_areEqual = cc_isr_core.StringExtensions.AreEqualDebug(p_expected, p_actual)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, p_actual, _
            "Literal values should be inserted in place of the escaped sequences.")
            
    If p_outcome.AssertSuccessful Then
    
        p_expected = p_original
        p_original = p_actual
        p_actual = cc_isr_core.StringExtensions.InsertEscapeSequences(p_original)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, p_actual, _
                "Escape sequences should be inserted in place of the literal characters.")
    End If

    Debug.Print p_outcome.BuildReport("TestInsertRepelaceEscapeSequences")
    
    Set TestInsertRepelaceEscapeSequences = p_outcome

End Function

''' <summary>   Unit test. Asserts inserting. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestInsert() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_original As String: p_original = "1234"
    Dim p_added As String: p_added = "99"
    
    Dim p_expected As String
    Dim p_position As Long
    Dim p_suffix As String
    
    p_position = 0: p_expected = "991234": p_suffix = "-th"
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
            StringExtensions.Insert(p_original, p_added, p_position), _
            "Added string '" & p_added & "' should be inserted into '" & _
            p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
            
    If p_outcome.AssertSuccessful Then
        p_position = 1: p_expected = "991234": p_suffix = "-st"
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_position = 2: p_expected = "199234": p_suffix = "-nd"
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_position = 3: p_expected = "129934": p_suffix = "-rd"
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_position = 4: p_expected = "123994": p_suffix = "-th"
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_position = 100: p_expected = "123499": p_suffix = "-th (after the last)"
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                StringExtensions.Insert(p_original, p_added, p_position), _
                "Added string '" & p_added & "' should be inserted into '" & _
                p_original & "' at the " & CStr(p_position) & p_suffix & " position.")
    End If
    
    Debug.Print p_outcome.BuildReport("TestInsert")
    
    Set TestInsert = p_outcome

End Function

''' <summary>   Unit test. Asserts delimited string element should pop. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestPop() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_delimitedString As String: p_delimitedString = "a,b,c"
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("a", _
            StringExtensions.Pop(p_delimitedString, ","), _
            "First element in " & p_delimitedString & " should pop")
            
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("b", _
            StringExtensions.Pop(p_delimitedString, ","), _
            "Second element in " & p_delimitedString & " should pop")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("c", _
            StringExtensions.Pop(p_delimitedString, ","), _
            "Third element in " & p_delimitedString & " should pop")
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(VBA.vbNullString, _
            StringExtensions.Pop(p_delimitedString, ","), _
            "No element in " & p_delimitedString & " should pop")
            
    End If
    
    Debug.Print p_outcome.BuildReport("TestPop")
    
    Set TestPop = p_outcome

End Function

''' <summary>   Unit test. Asserts creating a repeated string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestRepeat() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("aaa", StringExtensions.Repeat("a", 3), _
            "Should constract a string with repreated strings")
            
    Debug.Print p_outcome.BuildReport("TestRepeat")
    
    Set TestRepeat = p_outcome

End Function

''' <summary>   Unit test. Asserts start with. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStartsWith() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(StringExtensions.StartsWith("foobar", "foo"), _
            "String should start with the expected value.")
            
    Debug.Print p_outcome.BuildReport("TestStartsWith")
    
    Set TestStartsWith = p_outcome

End Function

''' <summary>   Unit test. Asserts parsing format string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestFormatStringParser() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_format As String
    Dim p_expected As String
    Dim p_actual As String
    Dim p_dateTime As Date
    Dim p_stringValue As String
    
    Dim p_formatGroup As String, p_precisionSpecifier As Integer
    Dim p_formatSpecifier As String, p_precisionString As String
    Dim p_itemIndex As Integer, p_success As Boolean, p_message As String

    Dim p_expectedFormatGroup As String
    Dim p_expectedFormatSpecifier As String, p_expectedPrecisionString As String
    Dim p_expectedItemIndex As Integer

    p_format = "{0:F11}"
    p_expectedItemIndex = 0
    p_expectedFormatGroup = "F11"
    p_expectedPrecisionString = "11"
    p_expectedFormatSpecifier = "F"
    
    p_success = cc_isr_core.StringExtensions.ParseFormatSpecification(p_format, p_itemIndex, p_formatGroup, _
                                         p_precisionString, p_formatSpecifier, _
                                         p_message)
        
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, "'" & p_format & "' should parse: " & p_message)
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedItemIndex, p_itemIndex, _
        "'" & p_format & "' item index should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFormatGroup, p_formatGroup, _
            "'" & p_format & "' format group should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedPrecisionString, p_precisionString, _
            "'" & p_format & "' precision string should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFormatSpecifier, p_formatSpecifier, _
            "'" & p_format & "' format specifier should equal.")
    End If
    
    p_format = "{1:F1}"
    p_expectedItemIndex = 1
    p_expectedFormatGroup = "F1"
    p_expectedPrecisionString = "1"
    p_expectedFormatSpecifier = "F"
    
    p_success = cc_isr_core.StringExtensions.ParseFormatSpecification(p_format, p_itemIndex, p_formatGroup, _
                                         p_precisionString, p_formatSpecifier, _
                                         p_message)
        
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, "'" & p_format & "' should parse: " & p_message)
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedItemIndex, p_itemIndex, _
        "'" & p_format & "' item index should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFormatGroup, p_formatGroup, _
            "'" & p_format & "' format group should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedPrecisionString, p_precisionString, _
            "'" & p_format & "' precision string should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFormatSpecifier, p_formatSpecifier, _
            "'" & p_format & "' format specifier should equal.")
    End If
    
    p_format = "{2:MMMM dd, yyyy}"
    p_expectedItemIndex = 2
    p_expectedFormatGroup = "MMMM dd, yyyy"
    p_expectedPrecisionString = ""
    p_expectedFormatSpecifier = "MMMM dd, yyyy"
    
    p_success = cc_isr_core.StringExtensions.ParseFormatSpecification(p_format, p_itemIndex, p_formatGroup, _
                                         p_precisionString, p_formatSpecifier, _
                                         p_message)
        
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_success, "'" & p_format & "' should parse: " & p_message)
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedItemIndex, p_itemIndex, _
        "'" & p_format & "' item index should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFormatGroup, p_formatGroup, _
            "'" & p_format & "' format group should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedPrecisionString, p_precisionString, _
            "'" & p_format & "' precision string should equal.")
    End If
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expectedFormatSpecifier, p_formatSpecifier, _
            "'" & p_format & "' format specifier should equal.")
    End If
    
    Debug.Print p_outcome.BuildReport("TestFormatStringParser")
    
    Set TestFormatStringParser = p_outcome

End Function


''' <summary>   Unit test. Asserts creating formatted date strings. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance where
'''             <see cref="Assert.AssertSuccessful"/> is True if the test passed. </returns>
Public Function TestDateStringFormat() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_format As String
    Dim p_expected As String
    Dim p_actual As String
    Dim p_dateTime As Date
    Dim p_stringValue As String
    
    p_stringValue = "12:00:00 AM"
    p_dateTime = CDate(p_stringValue)
    p_format = "{0:MMMM dd, yyyy}"
    p_expected = "December 30, 1899"
    p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, p_dateTime)
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
            p_actual, "'" & p_format & "' should format CDate('" & p_stringValue & "' as expected.")
    
    p_stringValue = "12:00:00 AM"
    p_format = "{0:MMMM dd, yyyy}"
    p_expected = "December 30, 1899"
    p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, p_stringValue)
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
            p_actual, "'" & p_format & "' should format '" & p_stringValue & "' as expected.")
    
    Debug.Print p_outcome.BuildReport("TestDateStringFormat")
    
    Set TestDateStringFormat = p_outcome
    

End Function

''' <summary>   Unit test. Asserts creating a formatted string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance where
'''             <see cref="Assert.AssertSuccessful"/> is True if the test passed. </returns>
Public Function TestStringFormat() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_format As String
    Dim p_expected As String
    Dim p_actual As String
    
    p_format = "a{0}{1}"
    p_expected = "abc"
    p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, "b", "c")
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
            p_actual, "'" & p_format & "' should format the expected value.")
            
    If p_outcome.AssertSuccessful Then
        p_format = "(B) Binary: {0:B}"
        p_expected = "(B) Binary: 10000101"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(B) Binary: {0:B16}"
        p_expected = "(B) Binary: 1111111110000101"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If
            
    If p_outcome.AssertSuccessful Then
        p_format = "(C) Currency: {0:C}\n"
        p_expected = "(C) Currency: -123.45$" & VBA.vbLf
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(C) Currency: {0:C}"
        p_expected = "(C) Currency: -123.00$"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_format = "(D) Decimal:. . . . . . . . . {0:D}"
        p_expected = "(D) Decimal:. . . . . . . . . -123"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If
    

    If p_outcome.AssertSuccessful Then
        p_format = "(E) Scientific: . . . . . . . {0:E}"
        p_expected = "(E) Scientific: . . . . . . . -1.23450E2"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_format = "(F) Fixed point:. . . . . . . {0:F}"
        p_expected = "(F) Fixed point:. . . . . . . -123.45"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_format = "(F) Fixed point:. . . . . . . {0:F1}"
        p_expected = "(F) Fixed point:. . . . . . . -123.5"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(G) General:. . . . . . . {0:G}"
        p_expected = "(G) General:. . . . . . . -1.23450E2"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(G) General:. . . . . . . {0:G4}"
        p_expected = "(G) General:. . . . . . . -123.5"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(N) Number: . . . . . . . . . {0:N}"
        p_expected = "(N) Number: . . . . . . . . . -123"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(P) Percent:. . . . . . . . . {0:P}"
        p_expected = "(P) Percent:. . . . . . . . . -12,345%"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(R) Round-trip: . . . . . . . {0:R}"
        p_expected = "(R) Round-trip: . . . . . . . -123.45"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, -123.45)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(X) Hexadecimal:. . . . . . . {0:X}"
        p_expected = "(X) Hexadecimal:. . . . . . . FF85"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, CInt(-123))
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(X) Hexadecimal:. . . . . . . {0:x}"
        p_expected = "(X) Hexadecimal:. . . . . . . ff85"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, CInt(-123))
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    Dim p_date As Date: p_date = DateValue("January 26, 2013") + TimeValue("8:28:11 PM")

    If p_outcome.AssertSuccessful Then
        p_format = "(c) Custom format: . . . . . .{0:cYYYY-MM-DD (MMMM)}"
        p_expected = "(c) Custom format: . . . . . .2013-01-26 (January)"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(d) Short date: . . . . . . . {0:d}"
        p_expected = "(d) Short date: . . . . . . . 1/26/2013"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(D) Long date:. . . . . . . . {0:D}"
        p_expected = "(D) Long date:. . . . . . . . Saturday, January 26, 2013"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    Dim p_time As Date
    p_time = TimeValue("8:28:11 PM")
    
    If p_outcome.AssertSuccessful Then
        p_format = "(T) Long time:. . . . . . . . {0:T}"
        p_expected = "(T) Long time:. . . . . . . . 8:28:11 PM"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(f) Full date/short time: . . {0:f}"
        p_expected = "(f) Full date/short time: . . Saturday, January 26, 2013 8:28 PM"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(F) Full date/long time:. . . {0:F}"
        p_expected = "(F) Full date/long time:. . . Saturday, January 26, 2013 8:28:11 PM"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        p_format = "(s) Sortable: . . . . . . . . {0:s}"
        p_expected = "(s) Sortable: . . . . . . . . 2013-01-26T20:28:11"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, p_date)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If

    If p_outcome.AssertSuccessful Then
        ' specify alignment (/padding) and to use escape sequences:
        
        p_format = "\q{0}, {1}!\x20\n'{2,10:C2}'\n'{2,-10:C2}'"
        p_expected = """hello, world! " & VBA.vbLf & "'   100.00$'" & VBA.vbLf & "'100.00$   '"
        p_actual = cc_isr_core.StringExtensions.StringFormat(p_format, "hello", "world", 100)
        
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "'" & p_format & "' should format the expected value.")
    End If
  
    Debug.Print p_outcome.BuildReport("TestStringFormat")
  
    Set TestStringFormat = p_outcome

End Function

''' <summary>   Unit test. Asserts creating a formatted string using simpel replacement. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStringFormatReplace() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("aaa", _
            StringExtensions.StringFormat("a{0}{1}", "a", "a"), _
            "Format should build the expected string")

    Debug.Print p_outcome.BuildReport("TestStringFormatReplace")
    
    Set TestStringFormatReplace = p_outcome
    
End Function

''' <summary>   Unit test. Asserts finding an item in a string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStringContains() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_value As String
    Dim p_candidate As String
    p_value = "the string contains"
    p_candidate = "contains"
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue( _
            StringExtensions.StringContains(p_value, p_candidate), _
            "The string '" & p_value & "' should contain '" & p_candidate & "'.")
            
    If p_outcome.AssertSuccessful Then

        p_value = "the string contains"
        p_candidate = "contained"
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse( _
                StringExtensions.StringContains(p_value, p_candidate), _
                "The string '" & p_value & "' should not contain '" & p_candidate & "'.")

    End If
    
    Debug.Print p_outcome.BuildReport("TestStringContains")
    
    Set TestStringContains = p_outcome

End Function

''' <summary>   Unit test. Asserts finding items in a string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestStringContainsAny() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_value As String
    Dim p_candidate1 As String, p_candidate2 As String
    p_value = "the string contains"
    p_candidate1 = "the"
    p_candidate2 = "contains"
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue( _
            StringExtensions.StringContainsAny(p_value, True, p_candidate1, p_candidate2), _
            "The string '" & p_value & "' should contain '" & p_candidate1 & "' or '" & p_candidate2 & "'.")
            
    If p_outcome.AssertSuccessful Then

        p_value = "the string contains"
        p_candidate1 = "the"
        p_candidate2 = "contained"
        Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue( _
                StringExtensions.StringContainsAny(p_value, True, p_candidate1, p_candidate2), _
                "The string '" & p_value & "' should contain '" & p_candidate1 & "' or '" & p_candidate2 & "'.")

    End If
    
    If p_outcome.AssertSuccessful Then

        p_value = "the string contains"
        p_candidate1 = "not"
        p_candidate2 = "contained"
        Set p_outcome = cc_isr_Test_Fx.Assert.IsFalse( _
                StringExtensions.StringContainsAny(p_value, True, p_candidate1, p_candidate2), _
                "The string '" & p_value & "' should not contain '" & p_candidate1 & "' or '" & p_candidate2 & "'.")

    End If
    
    Debug.Print p_outcome.BuildReport("TestStringContainsAny")
    
    Set TestStringContainsAny = p_outcome
    
End Function

''' <summary>   Unit test. Asserts sub-string. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestSubstring() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("oo", StringExtensions.Substring("foobar", 1, 2), _
            "Should get the expected part of the string")

    Debug.Print p_outcome.BuildReport("TestSubstring")
    
    Set TestSubstring = p_outcome

End Function

''' <summary>   Unit test. Asserts convertinh values to binary. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance where
'''             <see cref="Assert.AssertSuccessful"/> is True if the test passed. </returns>
Public Function TestToBinary() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_value As Long
    Dim p_expected As String
    Dim p_actual As String
    
    p_value = 5
    p_expected = "101"
    p_actual = cc_isr_core.StringExtensions.ToBinary(p_value)
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
            p_actual, "decimal value '" & CStr(p_value) & "' should convert to as expected.")
    
    If p_outcome.AssertSuccessful Then
        p_value = 16
        p_expected = "10000"
        p_actual = cc_isr_core.StringExtensions.ToBinary(p_value)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "decimal value '" & CStr(p_value) & "' should convert to as expected.")
    End If
    
    If p_outcome.AssertSuccessful Then
        p_value = 5
        p_expected = "00000101"
        p_actual = cc_isr_core.StringExtensions.ToBinary(p_value, 8)
        Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_expected, _
                p_actual, "decimal value '" & CStr(p_value) & "' should convert to as expected.")
    End If
    
    Debug.Print p_outcome.BuildReport("TestToBinary")
    
    Set TestToBinary = p_outcome
    

End Function

''' <summary>   Unit test. Asserts trim left. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTrimLeft() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("bar", _
        StringExtensions.TrimLeft("oobar", "o"), "String should be left trimmed.")

    Debug.Print p_outcome.BuildReport("TestTrimLeft")
    Set TestTrimLeft = p_outcome

End Function

''' <summary>   Unit test. Asserts trim right. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestTrimRight() As cc_isr_Test_Fx.Assert

    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual("f", _
        StringExtensions.TrimRight("foo", "o"), "String should be right-trimmed.")

    Debug.Print p_outcome.BuildReport("TestTrimRight")
    Set TestTrimRight = p_outcome

End Function


