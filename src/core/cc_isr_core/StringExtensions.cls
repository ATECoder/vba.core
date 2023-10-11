VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   String extensions. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Type this_
    PaddingCharacter As String
    EscapeSequences As Collection
End Type

Private This As this_

Option Explicit

''' <summary>   Report if the values are not equal. </summary>
''' <param name="a_expected">   [String] the expected value. </param>
''' <param name="a_actual">     [String] The actual value. </param>
''' <returns>   [Boolean]. </returns>
Public Function AreEqualDebug(ByVal a_expected As String, ByVal a_actual As String)

    Dim AreEqual As Boolean: AreEqual = True
    AreEqual = 1 = VBA.InStr(1, a_expected, a_actual)
    If Not AreEqual Then
    
        Dim p_expectedLength As Integer: p_expectedLength = VBA.Len(a_expected)
        Dim p_actualLength As Integer: p_actualLength = VBA.Len(a_actual)
        If p_actualLength <> p_expectedLength Then
        
            Debug.Print "Unequal length. Expected = " & VBA.CStr$(p_expectedLength) & _
                        " actual = " & VBA.CStr$(p_actualLength) & "."
        
        End If
        
        Dim i As Integer
        For i = 1 To IIf(p_expectedLength < p_actualLength, p_expectedLength, p_actualLength)
            If VBA.Mid$(a_expected, i, 1) <> VBA.Mid$(a_actual, i, 1) Then
                AreEqual = False
                Debug.Print "@" & i & "; Expected=" & VBA.Asc(VBA.Mid$(a_expected, i, 1)) & _
                            "; Actual=" & VBA.Asc(VBA.Mid$(a_actual, i, 1))
                Exit For
            End If
        Next i

    End If
    
    AreEqualDebug = AreEqual

End Function

''' <summary>   Gets the character at the given index from the given string. </summary>
''' <param name="a_text">    [String] The text to get the character from. </param>
''' <param name="a_index">   [Integer] The index of the character to get. </param>
''' <returns>   [String]. <returns>
Public Function CharAt(ByVal a_text As String, ByVal a_index As Integer) As String
    
    CharAt = VBA.Mid$(a_text, a_index, 1)

End Function

''' <summary>   Checks whether the given text ends with the given sequence. </summary>
''' <param name="a_text">            [String] The text to check for the sequence. </param>
''' <param name="a_endText">         [String] The text to be located at the end. </param>
''' <param name="a_compareMethod">   [Optional, VbCompareMethod, vbBinaryCompare] The compare method. </param>
''' <returns>   A Boolean. <returns>
Public Function EndsWith(ByVal a_text As String, ByVal a_endText As String, _
                           Optional ByVal a_compareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
    
    EndsWith = 0 = VBA.StrComp(VBA.Right$(a_text, VBA.Len(a_endText)), a_endText, a_compareMethod)

End Function

''' <summary>   Creates a new escape sequence. </summary>
''' <remarks>   Escape sequences are character combinations consisting of a backslash (\) followed by a letter
''' or by a combination of digits. Escape sequences are used to represent a newline character, single quotation mark,
''' or certain other characters in a character constant. In C#, an escape sequence is regarded as a single character
''' and is therefore valid as a character constant.
''' </remarks>
''' <param name="a_value">     [String] The value., e.g., '/r'. </param>
''' <param name="a_literal">   [String] The Literal value, e.g., <see cref="VBA.VbCr"/>. </param>
''' <returns>   [<see cref=""EscapeSequence/>]. </returns>
Public Function EscapeSequence(ByVal a_value As String, ByVal a_literal As String) As EscapeSequence
    Dim p_escapeSequence As New EscapeSequence
    p_escapeSequence.Value = a_value
    p_escapeSequence.Literal = a_literal
    Set EscapeSequence = p_escapeSequence
End Function

''' <summary>   Returns an array of <see cref="EscapeSequence"/> types. </summary>
''' <value>   An Array of <see cref="EscapeSequence"/> types. </summary>
Public Function EscapeSequences() As Collection

    Static p_escapes As Collection
    
    If p_escapes Is Nothing Then
        Dim p_value As String
        Set p_escapes = New Collection
        p_value = "\a": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(7)), p_value  ' alert, bell
        p_value = "\b": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(8)), p_value  ' backspace
        p_value = "\t": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(9)), p_value  ' tab
        p_value = "\n": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(10)), p_value ' new line
        p_value = "\v": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(11)), p_value ' vertical tab
        p_value = "\f": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(12)), p_value ' From Feed
        p_value = "\r": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(13)), p_value ' CR
        p_value = "\q": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(34)), p_value ' quote
    End If
        
    Set EscapeSequences = p_escapes
    
End Function

''' <summary>   Inserts escape sequences in place of their literal representation. </summary>
''' <param name="a_original">   [String] The original string. </param>
''' <returns>   [String]. <returns>
Function InsertEscapeSequences(ByVal a_original As String) As String

    Dim p_escape As EscapeSequence
    Dim p_escapes As Collection: Set p_escapes = Me.EscapeSequences
    
    For Each p_escape In p_escapes
        
        If Me.StringContains(a_original, p_escape.Literal) Then _
            a_original = Replace(a_original, p_escape.Literal, p_escape.Value)

    Next
    
    InsertEscapeSequences = a_original
    
End Function

''' <summary>   Replace white space with empty spaces. </summary>
''' <remarks>   <see href="https://stackoverflow.com/questions/42306358/removing-all-spaces-in-string"/>
''' </remarks>
''' <param name="a_original">   [String] The original string. </param>
''' <returns>   [String]. <returns>
Public Function RemoveWhiteSpace(a_original As String) As String
    With New RegExp
        .Pattern = "\s"
        .MultiLine = True
        .Global = True
        RemoveWhiteSpace = .Replace(a_original, vbNullString)
    End With
End Function

''' <summary>   Replace escape sequences with their literal representation. </summary>
''' <param name="a_original">   [String] The original string. </param>
''' <returns>   [String]. <returns>
Function ReplaceEscapeSequences(ByVal a_original As String) As String

    If Me.StringContains(a_original, "\") Then
    
        Dim p_escape As EscapeSequence
        Dim p_escapes As Collection: Set p_escapes = Me.EscapeSequences
        
        For Each p_escape In p_escapes
            
            If Me.StringContains(a_original, p_escape.Value) Then _
                a_original = Replace(a_original, p_escape.Value, p_escape.Literal)
    
            If Not Me.StringContains(a_original, "\") Then Exit For
        
        Next
    
    End If
    
    Dim char As Long
    Dim p_regex As VBScript_RegExp_55.RegExp
    Dim p_matches As VBScript_RegExp_55.MatchCollection
    Dim p_match As VBScript_RegExp_55.Match
    Dim p_submatch As String
    
    If Me.StringContains(a_original, "\") Then
    
        ' replace octal escape sequence
    
        Set p_regex = New VBScript_RegExp_55.RegExp
        p_regex.Pattern = "\\(\d{3})"
        p_regex.IgnoreCase = True
        p_regex.Global = True
        Set p_matches = p_regex.Execute(a_original)
    
        If p_matches.Count <> 0 Then
            For Each p_match In p_matches
            
                p_submatch = p_match.SubMatches(0)
                
                ' p_submatch contains the octal number representing the ASCII code we're after:
                
                p_submatch = "&O" & p_submatch ' prepend octal prefix
                char = CLng(p_submatch)
                a_original = Replace(a_original, p_match.Value, VBA.Chr$(char))
            Next
        End If
    
    End If

    If Me.StringContains(a_original, "\") Then
    
        ' replace "ASCII (hex)" escape sequence
        
        Set p_regex = New VBScript_RegExp_55.RegExp
        p_regex.Pattern = "\\x(\w{2})"
        p_regex.IgnoreCase = True
        p_regex.Global = True
        Set p_matches = p_regex.Execute(a_original)
        
        If p_matches.Count <> 0 Then
            For Each p_match In p_matches
                p_submatch = p_match.SubMatches(0)
                
                ' p_submatch contains the hex value representing the ASCII code we're after:
                
                p_submatch = "&H" & p_submatch ' prepend hex prefix
                char = CLng(p_submatch)
                a_original = Replace(a_original, p_match.Value, VBA.Chr$(char))
            Next
        End If
    
    End If

    ReplaceEscapeSequences = a_original
   

End Function

''' <summary>   Inserts or adds a string to an existing string at the specified position. </summary>
''' <param name="a_original">   [String] The original string. </param>
''' <param name="a_added">      [String] The string to insert or add. </param>
''' <param name="a_position">   [Long] The location for inserting the added string. </param>
''' <returns>   [String]. <returns>
Function Insert(ByVal a_original As String, a_added As String, a_position As Long) As String

    If a_position < 1 Then a_position = 1
    If VBA.Len(a_original) < a_position Then a_position = VBA.Len(a_original) + 1

    Insert = VBA.Mid$(a_original, 1, a_position - 1) _
                      & a_added _
                      & VBA.Mid$(a_original, a_position, VBA.Len(a_original) - a_position + 1)

End Function

''' <summary>   Returns true if the value is null or empty. </summary>
''' <param name="a_text">   [String] The string to check. </param>
''' <returns>   true if the value is null or empty. </returns>
Public Function IsNullOrEmpty(ByVal a_text As String) As Boolean
    IsNullOrEmpty = (a_text = vbNullString)
End Function

''' <summary>   Pops an element from a delimited string. </summary>
''' <remarks>   Use this function to extract a substring from a string up to but not including the delimiter.
''' </remarks>
''' <param name="a_delimitedString">   [String] A series of tokens where a remainder string is returned. </param>
''' <param name="a_delimiter">         [String] Specifies the delimiter to use for returning the token. </param>
''' <returns>   [String]. The first sub string ending but not including the the delimiter or the
''' original string if the delimiter was not located. </returns>
Public Function Pop(ByRef a_delimitedString As String, ByVal a_delimiter As String) As String

  ' find the delimiter in the string
  
  Dim p_position As Long: p_position = InStr(1, a_delimitedString, a_delimiter, VbCompareMethod.vbTextCompare)
  
  ' check if we have a delimiter in the string
  If p_position > 0 Then
  
    ' if we have a delimiter, return the string all the
    ' way up to but not including the delimiter
    Pop = VBA.Left$(a_delimitedString, p_position - 1)
    
    ' return the string past the delimiter
    a_delimitedString = VBA.Mid$(a_delimitedString, p_position + VBA.Len(a_delimiter))
    
  Else
    ' if the string does not include the delimiter, we
    ' must return the original string
    Pop = a_delimitedString
    
    ' and empty the returned string
    a_delimitedString = Empty
    
  End If
  
End Function

''' <summary>   Repeats the given string the given amount of times. </summary>
''' <param name="a_text">    [String] The text to repeat. </param>
''' <param name="a_count">   [Integer] The amount of times to repeat the given string. </param>
''' <returns>   [String]. <returns>
Public Function Repeat(ByVal a_text As String, ByVal a_count As Long) As String
    
    Repeat = vbNullString
    Dim i As Long
    For i = 1 To a_count
        Repeat = Repeat & a_text
    Next

End Function


''' <summary>   Checks whether a candidate is located in the given text. </summary>
''' <param name="a_text">            [String] The text to check for containment. </param>
''' <param name="a_candidate">       [String] The text to locate within the string. </param>
''' <param name="a_compareMethod">   [Optional, VbCompareMethod, vbBinaryCompare] The compare method. </param>
''' <returns>   A Boolean <returns>
Public Function StringContains(ByVal a_text As String, ByVal a_candidate As String, _
                               Optional ByVal a_compareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
    
    StringContains = 0 < InStr(1, a_text, a_candidate, a_compareMethod)

End Function

''' <summary>   Checks whether any values are located in the given text. </summary>
''' <param name="a_text">            [String] The text to check for containment. </param>
''' <param name="a_a_caseSensitive">   [Boolean] True if case sensitive. </param>
''' <param name="a_values">          [Variant] The text to locate within the string. </param>
Public Function StringContainsAny(ByVal string_source As String, ByVal a_caseSensitive As Boolean, _
    ParamArray find_values() As Variant) As Boolean

    Dim i As Integer, p_found As Boolean
    
    Dim p_compareMethod  As VbCompareMethod
    p_compareMethod = IIf(a_caseSensitive, VbCompareMethod.vbBinaryCompare, VbCompareMethod.vbTextCompare)
    
    For i = LBound(find_values) To UBound(find_values)
        p_found = 0 <> InStr(1, string_source, find_values(i), p_compareMethod)
        If p_found Then Exit For
    Next
    
    StringContainsAny = p_found
    
End Function


''' <summary>   Replaces braced numerical value with tokens similar to VB.NET String.Format
''' by without formatting suffixes. </summary>
''' <param name="a_format">   [String] The format string 'A{0} b{1}'. </param>
''' <param name="a_values">   [Variant, ParamArray] The parameter values matching the
'''                           <ParamRef name="a_format"/>. </param>
Public Function StringFormatReplace(ByVal a_format As String, ParamArray a_values() As Variant) As String
    
    Dim i As Long
    
    If VBA.vbNullString = a_format Then
        
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.NullArgumentError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".StringFormatReplace", _
                    "Argument name: 'a_format'."
    
    Else
    
        ' unbox the parameter array in case it came from another
        ' function which argument too is a parameter array.
    
        Dim p_tokens As Variant
        p_tokens = CoreExtensions.UnboxParameterArray(a_values)
        
        For i = LBound(p_tokens) To UBound(p_tokens)
            a_format = Replace$(a_format, "{" & i & "}", p_tokens(i))
        Next
        StringFormatReplace = a_format
    
    End If

End Function

''' <summary>   Parses a format specification into elements. </summary>
''' <param name="a_format">            [String] A format specification such as '{0:F11}' or '{1:MMMM, dd yy}'. </param>
''' <param name="a_itemIndex">         [Integer] The item index, e.g., 0 or 1 for the above formats, respectively. </param>
''' <param name="a_formatGroup">       [String] The format part of the format string,
'''                                    e.g., 'F11' or 'MMMM, dd yy' for the above formats, respectively. </param>
''' <param name="a_precisionString">   [String] The numeric precision string, e.g., 11 or Empty
'''                                    for the above formats, respectively. </param>
''' <param name="a_formatSpecifier">   [String] The format specifier part of the format group,
'''                                    e.g., 'F' or 'MMMM, dd yy' for the above formats, respectively. </param>
''' <param name="a_message">           [String] A success or failure message. </param>
''' <returns>   [Boolean] true if parsing succeeded. </returns>
Public Function ParseFormatSpecification(ByVal a_format As String, _
        ByRef a_itemIndex As Integer, ByRef a_formatGroup As String, _
        ByRef a_precisionString As String, ByRef a_formatSpecifier As String, _
        ByRef a_message As String) As Boolean

    Dim p_index As String
    Dim p_format As String: p_format = a_format

    If VBA.vbNullString = p_format Then
        ParseFormatSpecification = False
        a_message = "Format string is null."
    Else
        ' strip the braces.
        p_format = Me.TrimLeft(p_format, "{")
        p_format = Me.TrimRight(p_format, "}")
    End If

    If VBA.vbNullString = p_format Then
    
        ParseFormatSpecification = False
        a_message = "Format string is null after removal of its braces."
    
    ElseIf Not Me.StringContains(p_format, ":") Then
        p_index = p_format
        If VBA.IsNumeric(p_index) Then
            a_itemIndex = CInt(p_index)
            ParseFormatSpecification = True
        Else
            a_message = "Item index in format string '" & p_format & "' must be numeric."
            ParseFormatSpecification = False
        End If
    
    Else
        p_index = VBA.Left$(p_format, VBA.InStr(1, p_format, ":") - 1)
        If VBA.IsNumeric(p_index) Then
            a_itemIndex = CInt(p_index)
            
            a_formatGroup = VBA.Mid$(p_format, VBA.InStr(1, p_format, ":") + 1, (VBA.LenB(p_format) / 2) - 2)
            
            ParseFormatSpecification = Me.ParseFormatGroup(a_formatGroup, a_precisionString, a_formatSpecifier, a_message)
        
        Else
            
            a_message = "Item index in format string '" & p_format & "' must be numeric."
            ParseFormatSpecification = False
        
        End If
    
    End If
    
End Function

''' <summary>   Parses a format group into elements. </summary>
''' <param name="a_formatGroup">       [String] The format part of the format string,
'''                                    e.g., 'F11' or 'MMMM, dd yy' for the above formats, respectively. </param>
''' <param name="a_precisionString">   [String] The numeric precision string, e.g., 11 or Empty
'''                                    for the above formats, respectively. </param>
''' <param name="a_formatSpecifier">   [String] The format specifier part of the format group,
'''                                    e.g., 'F' or 'MMMM, dd yy' for the above formats, respectively. </param>
''' <param name="a_message">           [String] A success or failure message. </param>
''' <returns>   [Boolean] true if parsing succeeded. </returns>
Public Function ParseFormatGroup(ByVal a_formatGroup As String, _
        ByRef a_precisionString As String, ByRef a_formatSpecifier As String, _
        ByRef a_message As String) As Boolean

    If VBA.vbNullString = a_formatGroup Then
        ParseFormatGroup = False
        a_message = "Format group is null."
    Else
        a_formatSpecifier = a_formatGroup
        a_precisionString = VBA.vbNullString
        Do While VBA.IsNumeric(VBA.Right$(a_formatSpecifier, 1))
            a_precisionString = VBA.Right$(a_formatSpecifier, 1) & a_precisionString
            a_formatSpecifier = VBA.Left$(a_formatSpecifier, VBA.Len(a_formatSpecifier) - 1)
        Loop
        ParseFormatGroup = True
    End If
    
End Function


''' <summary>   Implements the .NET String.Format() on a value. </summary>
''' <param name="a_format">   [String] The format string, e.g., 'F11' or 'MMMM, dd yy'. </param>
''' <param name="a_value">    [Variant] The value to format. </param>
Public Function StringFormatValue(ByVal a_format As String, ByVal a_value As Variant) As String

    Const p_procedureName As String = "StringFormatDate"

    If VbVarType.vbDate = VarType(a_value) Then
    
        StringFormatValue = Me.StringFormatDate(a_format, a_value)
        
    ElseIf VBA.IsNumeric(a_value) Then
    
        StringFormatValue = Me.StringFormatNumber(a_format, a_value)
    
    Else
    
        Dim p_precisionSpecifier As Integer
        Dim p_formatSpecifier As String, p_precisionString As String
        Dim p_success As Boolean, p_message As String
    
        p_success = Me.ParseFormatGroup(a_format, p_precisionString, p_formatSpecifier, p_message)
        
        If p_success Then
        
            ' handle the case introduced by the Stack Overflow article of a custom date format that starts with a c.
            If VBA.Len(p_formatSpecifier) > 1 Then
                If Me.StartsWith(p_formatSpecifier, "c", vbTextCompare) Then
                    p_formatSpecifier = VBA.Right$(p_formatSpecifier, VBA.LenB(p_formatSpecifier) / 2 - 1)
                End If
            End If
            StringFormatValue = VBA.Format$(a_value, p_formatSpecifier)
        Else
        
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.FormatError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " " & p_message
    
        End If
    
    End If

End Function

''' <summary>   Implements the .NET String.Format() on a numeric value. </summary>
''' <param name="a_format">          [String] The format string, e.g., 'F11'. </param>
''' <param name="a_value">           [Variant] The value to format. </param>
''' <param name="a_denoteLiteral">   [Boolean] True to prefix 0b for binary literal or
'''                                  0x for hex literal. </param>
Public Function StringFormatNumber(ByVal a_format As String, _
        ByVal a_value As Variant, Optional ByVal a_denoteLiteral As Boolean = False) As String

    Const p_procedureName As String = "StringFormatDate"

    Dim p_precisionSpecifier As Integer, a_valueFormat As String, p_stringValue As String
    Dim p_formatSpecifier As String, p_precisionString As String
    Dim p_success As Boolean, p_message As String

    p_success = Me.ParseFormatGroup(a_format, p_precisionString, p_formatSpecifier, p_message)
    
    If Not p_success Then
    
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.FormatError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " " & p_message
    
    End If

    If p_precisionString <> vbNullString And Not IsNumeric(p_precisionString) Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.FormatError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
            " Precision string '" & p_precisionString & "' is not numeric."

    If p_precisionString = vbNullString Then p_precisionString = "0"

    p_stringValue = VBA.CStr(a_value)

    Select Case p_formatSpecifier


        ' Binary format
        
        Case "B", "b"
            
            ' Precision specifier determines number of digits in returned string.

            p_precisionSpecifier = CInt(p_precisionString)

            StringFormatNumber = IIf(a_denoteLiteral, "0b", VBA.vbNullString) & _
                                  Me.ToBinary(a_value, p_precisionSpecifier)

        ' CURRENCY format, formats string as currency.
        
        Case "C", "c"
        
            ' Precision specifier determines number of decimal digits.
            ' This implementation ignores regional settings
            ' (hard-coded group separator, decimal separator and currency sign).

            If p_precisionString = "0" Then p_precisionString = "2"

            p_precisionSpecifier = CInt(p_precisionString)
            a_valueFormat = "#,##0.$"

            ' if a non-zero precision is specified...
            
            a_valueFormat = Replace$(a_valueFormat, ".", "." & String$(p_precisionSpecifier, VBA.Chr$(48)))

            StringFormatNumber = VBA.Format$(p_stringValue, a_valueFormat)

        ' DECIMAL format, formats string as integer number.
        
        Case "D", "d"
            
            ' Precision specifier determines number of digits in returned string.

            p_precisionSpecifier = CInt(p_precisionString)
            a_valueFormat = "0"
            a_valueFormat = Right$(String$(p_precisionSpecifier, "0") & a_valueFormat, _
                IIf(p_precisionSpecifier = 0, VBA.Len(a_valueFormat), p_precisionSpecifier))

            StringFormatNumber = VBA.Format$(p_stringValue, a_valueFormat)

        ' EXPONENTIAL NOTATION format (aka "Scientific Notation")

        Case "E", "e"
        
            ' Precision specifier determines number of decimals in returned string.
            ' This implementation ignores regional settings'
            ' (hard-coded decimal separator).

            p_precisionSpecifier = CInt(p_precisionString)
            a_valueFormat = "0.00000#" & p_formatSpecifier & "-#" 'defaults to 6 decimals

            If LenB(a_format) > 2 And p_precisionSpecifier > 0 Then
                
                ' if a non-zero precision is specified...
                
                a_valueFormat = "0." & String$(p_precisionSpecifier - 1, VBA.Chr$(48)) & "#" & p_formatSpecifier & "-#"

            ElseIf VBA.LenB(a_format) > 2 And p_precisionSpecifier = 0 Then
                
                cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidFormatStringError, _
                        ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                        " Format group '" & a_format & "' length " & CStr(VBA.LenB(a_format)) & _
                        " is greater than 2 and the precision specifier is 0."
            
            End If
            
            StringFormatNumber = VBA.Format$(p_stringValue, a_valueFormat)

        ' FIXED-POINT format

        Case "F", "f"
            
            ' Precision specifier determines number of decimals in returned string.
            ' This implementation ignores regional settings'
            ' (hard-coded decimal separator).

            p_precisionSpecifier = CInt(p_precisionString)
            a_valueFormat = "0"
            If LenB(a_format) > 2 And p_precisionSpecifier > 0 Then
                
                ' if a non-zero precision is specified...
                
                a_valueFormat = (a_valueFormat & ".") & String$(p_precisionSpecifier, VBA.Chr$(48))
            
            Else
                
                ' no precision specified - default to 2 decimals:
                
                a_valueFormat = "0.00"
                
            End If

            StringFormatNumber = VBA.Format$(p_stringValue, a_valueFormat)

        ' GENERAL format (recursive)
        
        Case "G", "g"
            
            ' returns the shortest of either FIXED-POINT or SCIENTIFIC formats in case of a Double.
            ' returns DECIMAL format in case of a Integer or Long.

            Dim p_eNotation As String, p_ePower As Integer, p_specifier As String
            
            p_precisionSpecifier = IIf(CInt(p_precisionString) > 0, CInt(p_precisionString), _
                IIf(Me.StringContains(p_stringValue, "."), VBA.Len(p_stringValue) - InStr(1, p_stringValue, "."), 0))

            ' track character case of formatSpecifier:
            
            p_specifier = IIf(p_formatSpecifier = "G", "D", "d")

            If VbVarType.vbInteger = VBA.VarType(a_value) Or VbVarType.vbLong = VBA.VarType(a_value) Then
                
                ' Integer types: use {0:D} (recursive call):
                
                StringFormatNumber = StringFormatNumber(p_specifier, a_value)

            ElseIf VbVarType.vbDouble = VBA.VarType(a_value) Then
                
                ' Non-integer types: use {0:E}
                
                p_specifier = IIf(p_formatSpecifier = "G", "E", "e")

                ' evaluate the exponential notation value (recursive call):
                
                p_eNotation = StringFormatNumber(p_specifier, p_stringValue)

                ' get the power of eNotation:
                
                p_ePower = VBA.Mid$(p_eNotation, InStr(1, VBA.UCase$(p_eNotation), "E") + 1, _
                                VBA.Len(p_eNotation) - VBA.InStr(1, VBA.UCase$(p_eNotation), "E"))

                If p_ePower > -5 And Abs(p_ePower) < p_precisionSpecifier Then
                    
                    ' use {0:F} when ePower > -5 and abs(ePower) < precisionSpecifier:
                    ' evaluate the floating-point value (recursive call):
                    
                    p_specifier = IIf(p_formatSpecifier = "G", "F", "f")
                    
                    ' dh: set precision specifier per .NET format
                    
                    Dim p_power As Double: p_power = VBA.Log(VBA.Abs(a_value)) / VBA.Log(10#)
                    
                    Dim p_prevision As Integer: p_prevision = p_precisionSpecifier - CInt(p_power + 0.5)
                    
                    StringFormatNumber = StringFormatNumber(p_specifier & _
                         IIf(p_prevision > 0, CStr(p_prevision), vbNullString), a_value)
                Else
                    
                    ' fallback to {0:E} if previous rule didn't apply:
                    
                    StringFormatNumber = p_eNotation
                
                End If

            End If

        ' NUMERIC format, formats string as an integer or decimal number.
        
        Case "N", "n"
            
            ' Precision specifier determines number of decimal digits.
            ' This implementation ignores regional settings'
            ' (hard-coded group and decimal separators).

            p_precisionSpecifier = CInt(p_precisionString)
            If LenB(a_format) > 2 And p_precisionSpecifier > 0 Then
                
                ' if a non-zero precision is specified...
                
                a_valueFormat = "#,##0"
                a_valueFormat = (a_valueFormat & ".") & String$(p_precisionSpecifier, VBA.Chr$(48))

            Else
                ' only the "D" is specified
                
                a_valueFormat = "#,##0"
            End If

            StringFormatNumber = VBA.Format$(p_stringValue, a_valueFormat)

        ' PERCENT format. Formats string as a percentage.
        
        Case "P", "p"
            
            ' Value is multiplied by 100 and displayed with a percent symbol.
            ' Precision specifier determines number of decimal digits.

            a_valueFormat = "#,##0%"
            p_precisionSpecifier = CInt(p_precisionString)
            If LenB(a_format) > 2 And p_precisionSpecifier > 0 Then
                
                'if a non-zero precision is specified...
                
                a_valueFormat = "#,##0"
                a_valueFormat = (a_valueFormat & ".") & String$(p_precisionSpecifier, VBA.Chr$(48))

            Else
                
                ' only the "P" is specified
                a_valueFormat = "#,##0"
            
            End If

            ' Append the percentage sign to the format string:
            
            a_valueFormat = a_valueFormat & "%"

            StringFormatNumber = VBA.Format$(p_stringValue, a_valueFormat)

        ' ROUND-TRIP format (a string that can round-trip to an identical number)
        
        Case "R", "r"
            
            ' example: ?StringFormat("{0:R}", 0.0000000001141596325677345362656)
            '          ...returns "0.000000000114159632567735"

            ' convert value to a Double (chop off overflow digits):
            
            p_stringValue = CDbl(p_stringValue)

            StringFormatNumber = VBA.Format$(p_stringValue, a_valueFormat)
        
        ' HEX format. Formats a string as a Hexadecimal value.
        
        Case "X", "x"
            
            ' Precision specifier determines number of total digits.
            ' Returned string is prefixed with "&x" to specify Hex.

            p_stringValue = Hex$(a_value)
            p_precisionSpecifier = CInt(p_precisionString)

            If LenB(p_precisionString) > 0 Then 'precision here stands for left padding
                p_stringValue = VBA.Right$(VBA.String$(p_precisionSpecifier, "0") & p_stringValue, _
                           IIf(p_precisionSpecifier = 0, VBA.Len(p_stringValue), p_precisionSpecifier))
            End If

            ' add C# hex specifier, apply specified casing:
            ' (VB6 hex specifier would cause Format() to reverse the formatting):
            
            p_stringValue = IIf(a_denoteLiteral, "0x", VBA.vbNullString) & _
                             IIf(p_formatSpecifier = "X", VBA.UCase$(p_stringValue), VBA.LCase$(p_stringValue))

            StringFormatNumber = VBA.Format$(p_stringValue, a_valueFormat)

        Case Else

            If IsNumeric(p_formatSpecifier) And Val(a_format) = 0 Then
                
                p_formatSpecifier = a_format
                p_stringValue = VBA.Format$(p_stringValue, a_format)
                StringFormatNumber = VBA.Format$(p_stringValue, a_valueFormat)
            
            Else
                cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidFormatStringError, _
                        ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                        " Format specifier '" & p_formatSpecifier & "' is not numeric" & _
                        " or value of format group '" & Val(a_format) & "' is not zero."
            End If
    End Select
    
End Function

''' <summary>   Implements the .NET String.Format() on a date value. </summary>
''' <param name="a_format">   [String] The format string, e.g., 'MMMM, dd yy' or 'D'. </param>
''' <param name="a_value">    [Variant] The Date value to format. </param>
Public Function StringFormatDate(ByVal a_format As String, ByVal a_value As Date) As String

    Const p_procedureName As String = "StringFormatDate"

    Dim p_precisionSpecifier As Integer, p_valueFormat As String
    Dim p_formatSpecifier As String, p_precisionString As String
    Dim p_success As Boolean, p_message As String

    p_success = Me.ParseFormatGroup(a_format, p_precisionString, p_formatSpecifier, p_message)

    If Not p_success Then
    
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.FormatError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".StringFormatDate", _
                " " & p_message
    
    End If
    
    ' handle the case introduced by the Stack Overflow article of a custom date format that starts with a c.
    If VBA.Len(p_formatSpecifier) > 1 Then
        If Not Me.StartsWith(p_formatSpecifier, "c", vbTextCompare) Then
            p_formatSpecifier = "c" & p_formatSpecifier
        End If
        p_valueFormat = VBA.Right$(p_formatSpecifier, VBA.LenB(p_formatSpecifier) / 2 - 1)
        p_formatSpecifier = "c"
    End If
    
    Select Case p_formatSpecifier

        ' CUSTOM date/time format
        
        Case "c", "C"
            
            ' format already set.

        ' SHORT DATE format
        
        Case "d"
            
            p_valueFormat = "ddddd"


        ' LONG DATE format
        Case "D"
        
            p_valueFormat = "dddddd"


        ' FULL DATE format (short)
        
        Case "f"
            
            p_valueFormat = "dddddd h:mm AM/PM"

        
        ' FULL DATE format (long)
        Case "F"
        
            p_valueFormat = "dddddd ttttt"

        Case "g"
        
            p_valueFormat = "ddddd hh:mm AM/PM"

        Case "G"
        
            p_valueFormat = "ddddd ttttt"


        ' SORTABLE DATETIME format
        
        Case "s"
            
            p_valueFormat = "yyyy-mm-ddThh:mm:ss"


        ' SHORT TIME format
        
        Case "t"
            
            p_valueFormat = "hh:mm AM/PM"


        ' LONG TIME format
        
        Case "T"
            
            p_valueFormat = "ttttt"

        Case Else
            
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.FormatError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                    " Unknown format specifier '" & p_formatSpecifier & "'."
        
    End Select
    
    StringFormatDate = VBA.Format$(a_value, p_valueFormat)

End Function

''' <summary>   Implements the .NET String.Format(). </summary>
''' <remarks>   <see href="https://stackoverflow.com/questions/14534360/implementing-string-format-in-vb6"/></remarks>
''' <param name="a_format">   [String] The format string. </param>
''' <param name="a_values">   [Variant, ParamArray] The parameter p_values matching the
'''                           <ParamRef name="a_format"/>. </param>
Public Function StringFormat(ByVal a_format As String, ParamArray a_values() As Variant) As String

    Const p_procedureName As String = "StringFormat"

    Dim p_returnValue As String
    Dim p_valuesCount As Integer

    If VBA.vbNullString = a_format Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.NullArgumentError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".StringFormat", _
                    "Argument name: 'a_format'."
    
    ' unbox the parameter array in case it came from another
    ' function which argument too is a parameter array.

    Dim p_values As Variant
    p_values = CoreExtensions.UnboxParameterArray(a_values)
    
    ' use SPACE as default padding character
    
    If This.PaddingCharacter = VBA.vbNullString Then This.PaddingCharacter = VBA.Chr$(32)

    ' figure out number of passed p_values:
    
    p_valuesCount = UBound(p_values) + 1

    Dim p_regex As VBScript_RegExp_55.RegExp
    Dim p_matches As VBScript_RegExp_55.MatchCollection
    Dim p_match As VBScript_RegExp_55.Match
    Dim p_matchValue As String

    ' when a_format starts with "@", escapes are not replaced
    ' (string is treated as a literal string with placeholders)
    
    Dim p_useLiteral As Boolean
    
    ' indicates whether HEX specifier "0x" is to be replaced with VB '&H' specified
    Dim p_useVbHexSpecifier As Boolean: p_useVbHexSpecifier = False
    
    ' validate a_format:
    
    Set p_regex = New VBScript_RegExp_55.RegExp
    p_regex.Pattern = "{({{)*(\w+)(,-?\d+)?(:[^}]+)?}(}})*"
    p_regex.IgnoreCase = True
    p_regex.Global = True
    Set p_matches = p_regex.Execute(a_format)

    ' determine if parameter array argument count matches number of unique regex matches:
    
    Dim p_uniqueCount As Integer
    Dim p_commaSeparatedValues As String
    For Each p_match In p_matches
        If Not Me.StringContains(p_commaSeparatedValues, p_match.SubMatches(1)) Then
            p_uniqueCount = p_uniqueCount + 1
            p_commaSeparatedValues = p_commaSeparatedValues & p_match.SubMatches(1) & ","
        End If
    Next

    ' unique indices count must match p_valuesCount:
    
    If p_matches.Count > 0 And p_uniqueCount <> p_valuesCount Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.FormatError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                " Unique count " & VBA.CStr(p_uniqueCount) & " not equal the number of values of " & _
                VBA.CStr(p_valuesCount) & "."
                
    p_useLiteral = Me.StartsWith(a_format, "@")

    ' remove the "@" literal specifier
    
    If p_useLiteral Then a_format = VBA.Right$(a_format, VBA.Len(a_format) - 1)

    If Not p_useLiteral And Me.StringContains(a_format, "\\") Then _
        a_format = Replace(a_format, "\\", VBA.Chr$(27))

    If Me.StringContains(a_format, "\\") Then _
        a_format = Replace(a_format, "\\", VBA.Chr$(27))

    If p_matches.Count = 0 And a_format <> vbNullString And UBound(p_values) = -1 Then
        
        ' only a_format was specified: skip to checking escape sequences:
        
        p_returnValue = a_format
        
    ElseIf UBound(p_values) = -1 And p_matches.Count > 0 Then
       
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.NullArgumentError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Argument name: 'p_values'."
    
    Else
    
        p_returnValue = a_format
    
        ' dissect a_format:
    
        Dim i As Integer, p_stringValue As String, p_valueIndex As Integer ' i: iterator; v: value; p: placeholder
        Dim p_alignmentGroup As String, p_alignmentSpecifier As String
        Dim p_formattedValue As String, p_alignmentPadding As Integer
        Dim p_value As Variant, p_valueFormat As String
        
        p_valueFormat = vbNullString
    
        ' iterate regex matches (each match is a placeholder):
        
        For i = 0 To p_matches.Count - 1
    
            ' get the placeholder specified index:
            
            Set p_match = p_matches(i)
            p_valueIndex = CInt(p_match.SubMatches(1))
    
            ' if specified index (0-based) > p_uniqueCount (1-based), something's wrong:
            
            If p_valueIndex > p_uniqueCount - 1 Then _
                cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.FormatError, _
                        ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                        " Specified index " & VBA.CStr(p_valueIndex) & " exceeds " & VBA.CStr(p_uniqueCount - 1) & "."
            
            p_value = p_values(p_valueIndex)
            p_stringValue = VBA.CStr(p_value)
            
            ' get the alignment specifier if it is specified:
            
            p_alignmentGroup = p_match.SubMatches(2)
            If p_alignmentGroup <> vbNullString Then _
                p_alignmentSpecifier = Right$(p_alignmentGroup, LenB(p_alignmentGroup) / 2 - 1)
    
            ' get the format specifier if it is specified:
            
            p_matchValue = p_match.Value
            If Me.StringContains(p_matchValue, ":") Then
    
                Dim p_formatGroup As String, p_precisionSpecifier As Integer
                Dim p_formatSpecifier As String, p_precisionString As String
                Dim p_itemIndex As Integer, p_success As Boolean, p_message As String
    
                p_success = Me.ParseFormatSpecification(p_matchValue, p_itemIndex, p_formatGroup, _
                                             p_precisionString, p_formatSpecifier, _
                                             p_message)
    
                p_formattedValue = Me.StringFormatValue(p_formatGroup, p_value)
                
            Else
            
                p_formattedValue = VBA.Format$(p_stringValue, p_valueFormat)
            
            End If
    
            ' apply specified alignment specifier:
            
            If p_alignmentSpecifier <> vbNullString Then
    
                p_alignmentPadding = Abs(CInt(p_alignmentSpecifier))
                
                If CInt(p_alignmentSpecifier) < 0 Then
                    
                    'negative: left-justified alignment
                    
                    If p_alignmentPadding - VBA.Len(p_formattedValue) > 0 Then _
                        p_formattedValue = p_formattedValue & _
                            String$(p_alignmentPadding - VBA.Len(p_formattedValue), This.PaddingCharacter)
                Else
                    'positive: right-justified alignment
                    
                    If p_alignmentPadding - VBA.Len(p_formattedValue) > 0 Then _
                        p_formattedValue = String$(p_alignmentPadding - VBA.Len(p_formattedValue), This.PaddingCharacter) & p_formattedValue
                End If
            
            End If
    
            ' Replace C# hex specifier with VB6 hex specifier,
            ' only if hex specifier was introduced in this function:
            
            If Not p_useLiteral And p_useVbHexSpecifier And Me.StringContains(p_formattedValue, "0x") Then _
                p_formattedValue = Replace$(p_formattedValue, "0x", "&H")
    
            ' replace all occurrences of placeholder {i} with their formatted p_values:
            
            p_returnValue = Replace(p_returnValue, p_matchValue, p_formattedValue, Count:=1)
    
    
            ' reset before reiterating:
            
            p_valueFormat = vbNullString
            
        Next
    
    End If

    ' if there's no more backslashes, don't bother checking for the rest:
    
    If Not (p_useLiteral Or Not Me.StringContains(p_returnValue, "\")) Then _
        p_returnValue = Me.ReplaceEscapeSequences(p_returnValue)
    
    If Not p_useLiteral And Me.StringContains(p_returnValue, VBA.Chr$(27)) Then _
        p_returnValue = VBA.Replace(p_returnValue, VBA.Chr$(27), "\")
    
    StringFormat = p_returnValue
        
End Function

''' <summary>   Implements the .NET String.Format(). </summary>
''' <remarks>   <see href="https://stackoverflow.com/questions/14534360/implementing-string-format-in-vb6"/></remarks>
''' <param name="a_format">   [String] The format string. </param>
''' <param name="a_values">   [Variant, ParamArray] The parameter p_values matching the
'''                           <ParamRef name="a_format"/>. </param>
Public Function StringFormatGoTo(ByVal a_format As String, ParamArray a_values() As Variant) As String

    Const p_procedureName As String = "StringFormat"

    Dim p_returnValue As String
    Dim p_valuesCount As Integer

    If VBA.vbNullString = a_format Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.NullArgumentError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".StringFormatGoTo", _
                    "Argument name: 'a_format'."
    
    ' unbox the parameter array in case it came from another
    ' function which argument too is a parameter array.

    Dim p_values As Variant
    p_values = CoreExtensions.UnboxParameterArray(a_values)
    
    ' use SPACE as default padding character
    
    If This.PaddingCharacter = VBA.vbNullString Then This.PaddingCharacter = VBA.Chr$(32)

    ' figure out number of passed p_values:
    
    p_valuesCount = UBound(p_values) + 1

    Dim p_regex As VBScript_RegExp_55.RegExp
    Dim p_matches As VBScript_RegExp_55.MatchCollection
    Dim p_match As VBScript_RegExp_55.Match
    Dim p_matchValue As String
    Dim p_valueFormat As String

    ' when a_format starts with "@", escapes are not replaced
    ' (string is treated as a literal string with placeholders)
    
    Dim p_useLiteral As Boolean
    
    ' indicates whether HEX specifier "0x" is to be escaped or not
    ' dh: not used
    Dim p_useVbHexSpecifier As Boolean
    
    ' validate a_format:
    
    Set p_regex = New VBScript_RegExp_55.RegExp
    p_regex.Pattern = "{({{)*(\w+)(,-?\d+)?(:[^}]+)?}(}})*"
    p_regex.IgnoreCase = True
    p_regex.Global = True
    Set p_matches = p_regex.Execute(a_format)

    ' determine if parameter array argument count matches number of unique regex matches:
    
    Dim p_uniqueCount As Integer
    Dim p_commaSeparatedValues As String
    For Each p_match In p_matches
        If Not Me.StringContains(p_commaSeparatedValues, p_match.SubMatches(1)) Then
            p_uniqueCount = p_uniqueCount + 1
            p_commaSeparatedValues = p_commaSeparatedValues & p_match.SubMatches(1) & ","
        End If
    Next

    ' unique indices count must match p_valuesCount:
    
    If p_matches.Count > 0 And p_uniqueCount <> p_valuesCount Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.FormatError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & _
                " Unique count " & VBA.CStr(p_uniqueCount) & " not equal the number of values of " & _
                VBA.CStr(p_valuesCount) & "."
                
    p_useLiteral = Me.StartsWith(a_format, "@")

    ' remove the "@" literal specifier
    
    If p_useLiteral Then a_format = VBA.Right$(a_format, VBA.Len(a_format) - 1)

    If Not p_useLiteral And Me.StringContains(a_format, "\\") Then _
        a_format = Replace(a_format, "\\", VBA.Chr$(27))

    If Me.StringContains(a_format, "\\") Then _
        a_format = Replace(a_format, "\\", VBA.Chr$(27))

    If p_matches.Count = 0 And a_format <> vbNullString And UBound(p_values) = -1 Then
        
        ' only a_format was specified: skip to checking escape sequences:
        
        p_returnValue = a_format
        GoTo checkEscapes
        
    ElseIf UBound(p_values) = -1 And p_matches.Count > 0 Then
       
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.NullArgumentError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Argument name: 'p_values'."
    
    End If

    p_returnValue = a_format

    'dissect a_format:

    Dim i As Integer, p_stringValue As String, p As Integer 'i: iterator; v: value; p: placeholder
    Dim p_alignmentGroup As String, p_alignmentSpecifier As String
    Dim p_formattedValue As String, p_alignmentPadding As Integer
    Dim p_value As Variant

    ' iterate regex matches (each match is a placeholder):
    
    For i = 0 To p_matches.Count - 1

        ' get the placeholder specified index:
        
        Set p_match = p_matches(i)
        p = CInt(p_match.SubMatches(1))

        ' if specified index (0-based) > p_uniqueCount (1-based), something's wrong:
        
        If p > p_uniqueCount - 1 Then _
            cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.FormatError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Specified index " & VBA.CStr(p) & " exceeds " & VBA.CStr(p_uniqueCount - 1) & "."
        
        p_value = p_values(p)
        p_stringValue = CStr(p_value)
        
        ' get the alignment specifier if it is specified:
        
        p_alignmentGroup = p_match.SubMatches(2)
        If p_alignmentGroup <> vbNullString Then _
            p_alignmentSpecifier = Right$(p_alignmentGroup, LenB(p_alignmentGroup) / 2 - 1)

        ' get the format specifier if it is specified:
        
        p_matchValue = p_match.Value
        If Me.StringContains(p_matchValue, ":") Then

            Dim p_formatGroup As String, p_precisionSpecifier As Integer
            Dim p_formatSpecifier As String, p_precisionString As String
            Dim p_itemIndex As Integer, p_success As Boolean, p_message As String

            ' DH: using a parser to get the format elements
            
            p_success = Me.ParseFormatSpecification(p_matchValue, p_itemIndex, p_formatGroup, _
                                         p_precisionString, p_formatSpecifier, _
                                         p_message)

            ' get the string between ":" and "}":
            ' p_formatGroup = VBA.Mid$(p_matchValue, VBA.InStr(1, p_matchValue, ":") + 1, (VBA.LenB(p_matchValue) / 2) - 2)
            ' p_formatGroup = VBA.Left$(p_formatGroup, VBA.LenB(p_formatGroup) / 2 - 1)
            ' p_precisionString = VBA.Right$(p_formatGroup, VBA.LenB(p_formatGroup) / 2 - 1)
            ' p_formatSpecifier = VBA.Mid$(p_matchValue, VBA.InStr(1, p_matchValue, ":") + 1, 1)

            ' applicable formatting depends on the type of the value (yes, GOTO!!):
            
            ' DH: Fixed a bug handling routing of formatting.
            ' allow a format specifier such as {0:F11} or {0:MMM-dd-yy}
            
            If VbVarType.vbDate = VarType(p_values(p)) Then
            
                ' handle the original case of the custom
                ' date format that starts with a c.
                If VBA.Len(p_formatSpecifier) > 1 Then
                    If Not Me.StartsWith(p_formatSpecifier, "c", vbTextCompare) Then
                        p_formatSpecifier = "c" & p_formatSpecifier
                    End If
                    p_precisionString = VBA.Right$(p_formatSpecifier, VBA.LenB(p_formatSpecifier) / 2 - 1)
                    p_formatSpecifier = "c"
                End If
                GoTo DateTimeFormatSpecifiers
                
            ElseIf VbVarType.vbString = VarType(p_values(p)) And VBA.Len(p_formatSpecifier) > 1 Then
            
                ' if we have a string value and a long format string,
                ' just format it using VBA.
                p_valueFormat = p_formatSpecifier
                GoTo ApplyStringFormat
            
            End If
            
            ' If TypeName(p_values(p)) = "Date" Then GoTo DateTimeFormatSpecifiers
            
            If p_stringValue = vbNullString Then GoTo ApplyStringFormat

NumberFormatSpecifiers:

            If p_precisionString <> vbNullString And Not IsNumeric(p_precisionString) Then _
                cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.FormatError, _
                        ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                        " Precision string '" & p_precisionString & "' is not numeric."

            If p_precisionString = vbNullString Then p_precisionString = "0"

            Select Case p_formatSpecifier


                ' CURRENCY format, formats string as currency.
                
                Case "C", "c"
                
                    ' Precision specifier determines number of decimal digits.
                    ' This implementation ignores regional settings
                    ' (hard-coded group separator, decimal separator and currency sign).

                    If p_precisionString = "0" Then p_precisionString = "2"

                    p_precisionSpecifier = CInt(p_precisionString)
                    p_valueFormat = "#,##0.$"

                    ' if a non-zero precision is specified...
                    
                    p_valueFormat = Replace$(p_valueFormat, ".", "." & String$(p_precisionSpecifier, VBA.Chr$(48)))


                ' DECIMAL format, formats string as integer number.
                
                Case "D", "d"
                    
                    ' Precision specifier determines number of digits in returned string.

                    p_precisionSpecifier = CInt(p_precisionString)
                    p_valueFormat = "0"
                    p_valueFormat = Right$(String$(p_precisionSpecifier, "0") & p_valueFormat, _
                        IIf(p_precisionSpecifier = 0, VBA.Len(p_valueFormat), p_precisionSpecifier))


                ' EXPONENTIAL NOTATION format (aka "Scientific Notation")

                Case "E", "e"
                
                    ' Precision specifier determines number of decimals in returned string.
                    ' This implementation ignores regional settings'
                    ' (hard-coded decimal separator).

                    p_precisionSpecifier = CInt(p_precisionString)
                    p_valueFormat = "0.00000#" & p_formatSpecifier & "-#" 'defaults to 6 decimals

                    If LenB(p_formatGroup) > 2 And p_precisionSpecifier > 0 Then
                        
                        ' if a non-zero precision is specified...
                        
                        p_valueFormat = "0." & String$(p_precisionSpecifier - 1, VBA.Chr$(48)) & "#" & p_formatSpecifier & "-#"

                    ElseIf VBA.LenB(p_formatGroup) > 2 And p_precisionSpecifier = 0 Then
                        
                        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidFormatStringError, _
                                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                                " Format group '" & p_formatGroup & "' length " & VBA.CStr(VBA.LenB(p_formatGroup)) & _
                                " is greater than 2 and the precision specifier is 0."
                    
                    End If

                ' FIXED-POINT format

                Case "F", "f"
                    
                    ' Precision specifier determines number of decimals in returned string.
                    ' This implementation ignores regional settings'
                    ' (hard-coded decimal separator).

                    p_precisionSpecifier = CInt(p_precisionString)
                    p_valueFormat = "0"
                    If LenB(p_formatGroup) > 2 And p_precisionSpecifier > 0 Then
                        
                        ' if a non-zero precision is specified...
                        
                        p_valueFormat = (p_valueFormat & ".") & String$(p_precisionSpecifier, VBA.Chr$(48))
                    
                    Else
                        
                        ' no precision specified - default to 2 decimals:
                        
                        p_valueFormat = "0.00"
                        
                    End If


                ' GENERAL format (recursive)
                
                Case "G", "g"
                    
                    ' returns the shortest of either FIXED-POINT or SCIENTIFIC formats in case of a Double.
                    ' returns DECIMAL format in case of a Integer or Long.

                    Dim p_eNotation As String, p_ePower As Integer, p_specifier As String
                    
                    p_precisionSpecifier = IIf(CInt(p_precisionString) > 0, CInt(p_precisionString), _
                        IIf(Me.StringContains(p_stringValue, "."), VBA.Len(p_stringValue) - InStr(1, p_stringValue, "."), 0))

                    ' track character case of formatSpecifier:
                    
                    p_specifier = IIf(p_formatSpecifier = "G", "D", "d")

                    If TypeName(p_values(p)) = "Integer" Or TypeName(p_values(p)) = "Long" Then
                        
                        ' Integer types: use {0:D} (recursive call):
                        
                        p_formattedValue = StringFormat("{0:" & p_specifier & "}", p_values(p))

                    ElseIf TypeName(p_values(p)) = "Double" Then
                        
                        ' Non-integer types: use {0:E}
                        
                        p_specifier = IIf(p_formatSpecifier = "G", "E", "e")

                        ' evaluate the exponential notation value (recursive call):
                        
                        p_eNotation = StringFormat("{0:" & p_specifier & "}", p_stringValue)

                        ' get the power of eNotation:
                        
                        p_ePower = VBA.Mid$(p_eNotation, InStr(1, VBA.UCase$(p_eNotation), "E-") + 1, _
                                        VBA.Len(p_eNotation) - VBA.InStr(1, VBA.UCase$(p_eNotation), "E-"))

                        If p_ePower > -5 And Abs(p_ePower) < p_precisionSpecifier Then
                            
                            ' use {0:F} when ePower > -5 and abs(ePower) < precisionSpecifier:
                            ' evaluate the floating-point value (recursive call):
                            
                            p_specifier = IIf(p_formatSpecifier = "G", "F", "f")
                            p_formattedValue = StringFormat("{0:" & p_formatSpecifier & _
                                 IIf(p_precisionSpecifier <> 0, p_precisionString, vbNullString) & "}", p_values(p))
                        Else
                            
                            ' fallback to {0:E} if previous rule didn't apply:
                            
                            p_formattedValue = p_eNotation
                        
                        End If

                    End If

                    GoTo AlignFormattedValue 'Skip the "ApplyStringFormat" step, it's applied already.


                ' NUMERIC format, formats string as an integer or decimal number.
                
                Case "N", "n"
                    
                    ' Precision specifier determines number of decimal digits.
                    ' This implementation ignores regional settings'
                    ' (hard-coded group and decimal separators).

                    p_precisionSpecifier = CInt(p_precisionString)
                    If LenB(p_formatGroup) > 2 And p_precisionSpecifier > 0 Then
                        
                        ' if a non-zero precision is specified...
                        
                        p_valueFormat = "#,##0"
                        p_valueFormat = (p_valueFormat & ".") & String$(p_precisionSpecifier, VBA.Chr$(48))

                    Else
                        ' only the "D" is specified
                        
                        p_valueFormat = "#,##0"
                    End If


                ' PERCENT format. Formats string as a percentage.
                
                Case "P", "p"
                    
                    ' Value is multiplied by 100 and displayed with a percent symbol.
                    ' Precision specifier determines number of decimal digits.

                    p_valueFormat = "#,##0%"
                    p_precisionSpecifier = CInt(p_precisionString)
                    If LenB(p_formatGroup) > 2 And p_precisionSpecifier > 0 Then
                        
                        'if a non-zero precision is specified...
                        
                        p_valueFormat = "#,##0"
                        p_valueFormat = (p_valueFormat & ".") & String$(p_precisionSpecifier, VBA.Chr$(48))

                    Else
                        
                        ' only the "P" is specified
                        p_valueFormat = "#,##0"
                    
                    End If

                    ' Append the percentage sign to the format string:
                    
                    p_valueFormat = p_valueFormat & "%"


                ' ROUND-TRIP format (a string that can round-trip to an identical number)
                
                Case "R", "r"
                    
                    ' example: ?StringFormat("{0:R}", 0.0000000001141596325677345362656)
                    '          ...returns "0.000000000114159632567735"

                    ' convert value to a Double (chop off overflow digits):
                    
                    p_stringValue = CDbl(p_stringValue)

                
                ' HEX format. Formats a string as a Hexadecimal value.
                
                Case "X", "x"
                    
                    ' Precision specifier determines number of total digits.
                    ' Returned string is prefixed with "&x" to specify Hex.

                    p_stringValue = Hex$(p_value)
                    p_precisionSpecifier = CInt(p_precisionString)

                    If LenB(p_precisionString) > 0 Then 'precision here stands for left padding
                        p_stringValue = VBA.Right$(VBA.String$(p_precisionSpecifier, "0") & p_stringValue, _
                                   IIf(p_precisionSpecifier = 0, VBA.Len(p_stringValue), p_precisionSpecifier))
                    End If

                    ' add C# hex specifier, apply specified casing:
                    ' (VB6 hex specifier would cause Format() to reverse the formatting):
                    
                    p_stringValue = "0x" & IIf(p_formatSpecifier = "X", VBA.UCase$(p_stringValue), VBA.LCase$(p_stringValue))

                Case Else

                    If IsNumeric(p_formatSpecifier) And Val(p_formatGroup) = 0 Then
                        
                        p_formatSpecifier = p_formatGroup
                        p_stringValue = Format(p_stringValue, p_formatGroup)
                    
                    Else
                        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidFormatStringError, _
                                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                                " Format specifier '" & p_formatSpecifier & "' is not numeric" & _
                                " or value of format group '" & Val(p_formatGroup) & "' is not zero."
                    End If
            End Select

            GoTo ApplyStringFormat


DateTimeFormatSpecifiers:
                
            Select Case p_formatSpecifier


                ' CUSTOM date/time format
                
                Case "c", "C"
                    
                    'let VB Format() parse precision specifier as is:
                    
                    p_valueFormat = p_precisionString


                ' SHORT DATE format
                
                Case "d"
                    
                    p_valueFormat = "ddddd"


                ' LONG DATE format
                Case "D"
                
                    p_valueFormat = "dddddd"


                ' FULL DATE format (short)
                
                Case "f"
                    
                    p_valueFormat = "dddddd h:mm AM/PM"

                
                ' FULL DATE format (long)
                Case "F"
                
                    p_valueFormat = "dddddd ttttt"

                Case "g"
                
                    p_valueFormat = "ddddd hh:mm AM/PM"

                Case "G"
                
                    p_valueFormat = "ddddd ttttt"


                ' SORTABLE DATETIME format
                
                Case "s"
                    
                    p_valueFormat = "yyyy-mm-ddThh:mm:ss"


                ' SHORT TIME format
                
                Case "t"
                    
                    p_valueFormat = "hh:mm AM/PM"


                ' LONG TIME format
                
                Case "T"
                    
                    p_valueFormat = "ttttt"

                Case Else
                    
                    cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.FormatError, _
                            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                            " Unknown format specifier '" & p_formatSpecifier & "'."
                
            End Select
            GoTo ApplyStringFormat

        End If

ApplyStringFormat:
            
        'apply computed format string:
        
        p_formattedValue = Format(p_stringValue, p_valueFormat)

AlignFormattedValue:
        
        'apply specified alignment specifier:
        
        If p_alignmentSpecifier <> vbNullString Then

            p_alignmentPadding = Abs(CInt(p_alignmentSpecifier))
            
            If CInt(p_alignmentSpecifier) < 0 Then
                
                'negative: left-justified alignment
                
                If p_alignmentPadding - VBA.Len(p_formattedValue) > 0 Then _
                    p_formattedValue = p_formattedValue & _
                        String$(p_alignmentPadding - VBA.Len(p_formattedValue), This.PaddingCharacter)
            Else
                'positive: right-justified alignment
                
                If p_alignmentPadding - VBA.Len(p_formattedValue) > 0 Then _
                    p_formattedValue = String$(p_alignmentPadding - VBA.Len(p_formattedValue), This.PaddingCharacter) & p_formattedValue
            End If
        
        End If

        ' Replace C# hex specifier with VB6 hex specifier,
        ' only if hex specifier was introduced in this function:
        
        If Not p_useLiteral And p_useVbHexSpecifier And Me.StringContains(p_formattedValue, "0x") Then _
            p_formattedValue = Replace$(p_formattedValue, "0x", "&H")

        ' replace all occurrences of placeholder {i} with their formatted p_values:
        
        p_returnValue = Replace(p_returnValue, p_matchValue, p_formattedValue, Count:=1)

        ' reset before reiterating:
        
        p_valueFormat = vbNullString
    
    Next

checkEscapes:

    ' if there's no more backslashes, don't bother checking for the rest:
    
    If p_useLiteral Or Not Me.StringContains(p_returnValue, "\") Then GoTo normalExit

    p_returnValue = ReplaceEscapeSequences(p_returnValue)
    

normalExit:

    If Not p_useLiteral And Me.StringContains(p_returnValue, VBA.Chr$(27)) Then _
        p_returnValue = Replace(p_returnValue, VBA.Chr$(27), "\")
    
    StringFormatGoTo = p_returnValue
        
End Function


''' <summary>   Converts a regular string to a Unicode string,
''' e.g., "foo" to a L"foo" string. </summary>
''' <param name="a_text">   [String] The string to convert. </param>
''' <returns>   [String]. <returns>
Public Function StringToWideString(ByVal a_text As String) As String
    
    StringToWideString = StrConv(a_text, vbUnicode)

End Function

''' <summary>   Checks whether the given text starts with the given sequence. </summary>
''' <param name="a_text">            [String] The text to check for the start. </param>
''' <param name="a_startText">       [String] The text to be located at the start. </param>
''' <param name="a_compareMethod">   [Optional, VbCompareMethod, vbBinaryCompare] The compare method. </param>
''' <returns>   A Boolean <returns>
Public Function StartsWith(ByVal a_text As String, ByVal a_startText As String, _
                           Optional ByVal a_compareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
    
    StartsWith = 1 = InStr(1, a_text, a_startText, a_compareMethod)

End Function

''' <summary>   Gets the substring from the given text. </summary>
''' <param name="a_text">         [String] The text to get the substring from. </param>
''' <param name="a_startIndex">   [Integer] The index of the first character of the substring. </param>
''' <param name="a_length">       [Optional, null] The number of characters to take from the original string. </param>
''' <returns>   [String]. <returns>
Public Function Substring(ByVal a_text As String, ByVal a_startIndex As Integer, _
        Optional ByVal a_length As Variant) As String
    
    If a_startIndex > VBA.Len(a_text) Then
        a_startIndex = VBA.Len(a_text)
    End If

    If IsMissing(a_length) Then
        a_length = VBA.Len(a_text) - a_startIndex
    End If
    
    If a_length > VBA.Len(a_text) Then
        a_length = VBA.Len(a_text) - a_startIndex
    End If

    Substring = VBA.Left$(VBA.Right$(a_text, VBA.Len(a_text) - a_startIndex), a_length)
    
End Function


''' <summary>   Convert a decimal value to a binary string. </summary>
''' <remarks>
''' <see href="https://www.developerfusion.com/code/5430/convert-decimal-integer-values-to-binary-string-in-vb6/"/>
''' </remarks>
''' <param name="a_value">          [Long] The value to convert. </param>
''' <param name="a_numberOfBits">   [Integer] the number of bits. 1 if fitting bit count. </param>
''' <returns>   [String]. <returns>
Public Function ToBinary(ByVal a_value As Long, Optional a_numberOfBits As Integer = 1) As String
  
    Dim p_increment As Integer: p_increment = a_numberOfBits
    p_increment = IIf(p_increment < 1, 1, p_increment)
    
    ' make sure there are enough bits to contain the number
    If a_value >= 0 Then
        Do While a_value > ((2 ^ a_numberOfBits) - 1)
            a_numberOfBits = a_numberOfBits + p_increment
        Loop
    Else
        Do While -a_value > (2 ^ (a_numberOfBits - 1))
            a_numberOfBits = a_numberOfBits + p_increment
        Loop
    End If
    
    ToBinary = vbNullString
    
    ' build the string
    Dim p_bit As Long: p_bit = 1
    Dim i As Integer
    For i = 0 To (a_numberOfBits - 1)
        ToBinary = CStr((a_value And p_bit) / p_bit) & ToBinary
        p_bit = p_bit * 2
    Next i
  
End Function

''' <summary>   Trims the given character from the start of the text. </summary>
''' <param name="a_text">   [String] The text to trim. </param>
''' <param name="a_char">   [String] The character to trim. </param>
''' <returns>   [String]. <returns>
Public Function TrimLeft(ByVal a_text As String, ByVal a_char As String) As String
    
    Dim p_textLength As Long
    p_textLength = VBA.Len(a_text)
    
    Dim p_firstCharacter As String
    
    Do While p_textLength > 0
        p_firstCharacter = VBA.Left$(a_text, 1)
        
        If p_firstCharacter <> a_char Then
            Exit Do
        End If
        
        a_text = VBA.Right$(a_text, p_textLength - 1)
        p_textLength = VBA.Len(a_text)
    Loop
    
    TrimLeft = a_text
End Function

''' <summary>   Trims the given character from the end of the given text. </summary>
''' <param name="a_text">   [String] The text to trim. </param>
''' <param name="a_char">   [String] The character to trim. </param>
''' <returns>   [String]. <returns>
Public Function TrimRight(ByVal a_text As String, ByVal a_char As String) As String
    
    Dim p_textLength As Long: p_textLength = VBA.Len(a_text)
    
    Dim p_lastCharacter As String
        
    Do While p_textLength > 0
        p_lastCharacter = VBA.Right$(a_text, 1)
        
        If p_lastCharacter <> a_char Then
            Exit Do
        End If
        
        a_text = VBA.Left$(a_text, p_textLength - 1)
        p_textLength = VBA.Len(a_text)
    Loop
    
    TrimRight = a_text
End Function

''' <summary>   Trims the given characters from the end of the given text. </summary>
''' <param name="a_text">            [String] The text to trim. </param>
''' <param name="a_trim">            [String] The characters to trim. </param>
''' <param name="a_compareMethod">   [Optional, <see cref="VBA.VbCompareMethod"/>,
'''                                  <see cref="VBA.VbCompareMethod.vbTextCompare"/>]
'''                                  The compare method. </param>
''' <returns>   [String]. <returns>
Public Function TrimEnd(ByVal a_text As String, ByVal a_trim As String, _
    Optional ByVal a_compareMethod As VBA.VbCompareMethod = VBA.VbCompareMethod.vbTextCompare) As String

    ' get the location of the trim
    Dim p_trimIndex As Integer
    p_trimIndex = VBA.InStr(1, a_text, a_trim, a_compareMethod)
    If p_trimIndex > 0 Then
        TrimEnd = VBA.Left$(a_text, p_trimIndex - 1)
    Else
        TrimEnd = a_text
    End If
    
End Function

''' <summary>   Trims the given characters from the start of the given text. </summary>
''' <param name="a_text">   [String] The text to trim. </param>
''' <param name="a_trim">   [String] The characters to trim. </param>
''' <returns>   [String]. <returns>
Public Function TrimStart(ByVal a_text As String, ByVal a_trim As String) As String

    If 1 = VBA.InStr(1, a_text, a_trim, vbTextCompare) Then
        TrimStart = VBA.Right$(a_text, Len(a_text) - Len(a_trim))
    Else
        TrimStart = a_text
    End If
    
End Function

''' + + + + + + + + + + + + + + + + + + + + + + + + + + + +
''' Parse values
''' + + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Tries to parse a number to Boolean. </summary>
''' <remnarks>   Note that this parsing converts 0 to false and any number to true
''' whereas in VBA, only -1 is true.</remarks>
''' <param name="a_text">      [String] The text to parse. </param>
''' <param name="a_value">     [Boolean] The parsed value. </param>
''' <param name="a_details">   [Out, String] details why parsing failed. </param>
''' <returns>   [Boolean] True if the text was parsed. <returns>
Public Function TryParseBoolean(ByVal a_text As String, ByRef a_value As Boolean, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseBoolean"
    
    Dim p_parsed As Boolean: p_parsed = False
    Dim p_details As String: p_details = VBA.vbNullString: a_details = VBA.vbNullString
    
    Dim p_trueValues As String: p_trueValues = "on,true"
    Dim p_falseValues As String: p_falseValues = "off,false"
    Dim p_value As Boolean
    
    ' default to leave the value unchangd.
    If VBA.vbNullString = a_text Then
        
        p_parsed = False
        p_details = "Unable to parse an empty string '" & a_text & "' to Boolean."
    
    ElseIf VBA.IsNumeric(a_text) Then
        
        On Error Resume Next
        ' any value other than zero is true.
        p_value = 0 <> VBA.CInt(VBA.Val(a_text))
        If Err.Number = 0 Then
            a_value = p_value
            p_parsed = True
        Else
            p_details = "Error " & VBA.CStr(Err.Number) & ", " & Err.Description & _
                " converting '" & a_text & "' to Boolean using '0 <> VBA.CInt(VBA.Val(" & a_text & "))'."
            p_parsed = False
        End If
        On Error GoTo 0
        
    Else
    
        Dim p_item As Variant
        For Each p_item In VBA.Split(p_trueValues, ",")
            If 0 = VBA.StrComp(VBA.CStr(p_item), a_text, vbTextCompare) Then
                p_parsed = True
                p_value = True
                Exit For
            End If
        Next
        
        If Not p_parsed Then
            For Each p_item In VBA.Split(p_falseValues, ",")
                If 0 = VBA.StrComp(VBA.CStr(p_item), a_text, vbTextCompare) Then
                    p_parsed = True
                    p_value = False
                    Exit For
                End If
            Next
        End If
        
        If p_parsed Then
            a_value = p_value
        Else
            p_details = "Unable to convert '" & a_text & "' to Boolean."
        End If
        
    End If

    If Not p_parsed Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & "; " & p_details

    TryParseBoolean = p_parsed

End Function

''' <summary>   Tries to parse a number to double. </summary>
''' <param name="a_text">      [String] The text to parse. </param>
''' <param name="a_value">     [Double] The parsed value. </param>
''' <param name="a_details">   [Out, String] details why parsing failed. </param>
''' <returns>   [Boolean] True if the text was parsed. <returns>
Public Function TryParseDouble(ByVal a_text As String, ByRef a_value As Double, ByRef a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseDouble"
    
    Dim p_parsed As Boolean: p_parsed = False
    Dim p_details As String: p_details = VBA.vbNullString: a_details = VBA.vbNullString
    Dim p_value As Double
    
    ' default to leave the value unchangd.
    If VBA.vbNullString = a_text Then
        p_details = "Unable to parse an empty string '" & a_text & "' to Double."
    ElseIf Not VBA.IsNumeric(a_text) Then
        p_details = "Unable to parse a non-numeric value '" & a_text & "' to Double."
    Else
        On Error Resume Next
        p_value = VBA.CDbl(VBA.Val(a_text))
        If Err.Number = 0 Then
            a_value = p_value
            p_parsed = True
        Else
            p_details = "Error " & VBA.CStr(Err.Number) & ", " & Err.Description & _
                " converting '" & a_text & "' to Double using VBA.CDbl(VBA.Val(" & a_text & "))."
            p_parsed = False
        End If
        On Error GoTo 0
    End If

    If Not p_parsed Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & "; " & p_details

    TryParseDouble = p_parsed

End Function

''' <summary>   Tries to parse a number to Integer. </summary>
''' <param name="a_text">      [String] The text to parse. </param>
''' <param name="a_value">     [Integer] The parsed value. </param>
''' <param name="a_details">   [String] details why parsing failed. </param>
''' <returns>   [Boolean] True if the text was parsed. <returns>
Public Function TryParseInteger(ByVal a_text As String, ByRef a_value As Integer, ByVal a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseInteger"
    
    Dim p_parsed As Boolean: p_parsed = False
    Dim p_details As String: p_details = VBA.vbNullString: a_details = VBA.vbNullString
    Dim p_value As Integer
    
    ' default to leave the value unchangd.
    If VBA.vbNullString = a_text Then
        p_details = "Unable to parse an empty string '" & a_text & "' to Integer."
    ElseIf Not VBA.IsNumeric(a_text) Then
        p_details = "Unable to parse a non-numeric value '" & a_text & "' to Integer."
    Else
        On Error Resume Next
        p_value = VBA.CInt(VBA.Val(a_text))
        If Err.Number = 0 Then
            a_value = p_value
            p_parsed = True
        Else
            p_details = "Error " & VBA.CStr(Err.Number) & ", " & Err.Description & _
                " converting '" & a_text & "' to Integer using VBA.CInt(VBA.Val(" & a_text & "))."
            p_parsed = False
        End If
        On Error GoTo 0
    End If
    
    If Not p_parsed Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & "; " & p_details

    TryParseInteger = p_parsed

End Function


''' <summary>   Tries to parse a number to Long. </summary>
''' <param name="a_text">      [String] The text to parse. </param>
''' <param name="a_value">     [Long] The parsed value. </param>
''' <param name="a_details">   [String] details why parsing failed. </param>
''' <returns>   [Boolean] True if the text was parsed. <returns>
Public Function TryParseLong(ByVal a_text As String, ByRef a_value As Long, ByVal a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseLong"
    
    Dim p_parsed As Boolean: p_parsed = False
    Dim p_details As String: p_details = VBA.vbNullString: a_details = VBA.vbNullString
    Dim p_value As Long
    
    ' default to leave the value unchangd.
    If VBA.vbNullString = a_text Then
        p_details = "Unable to parse an empty string '" & a_text & "' to Long."
    ElseIf Not VBA.IsNumeric(a_text) Then
        p_details = "Unable to parse a non-numeric value '" & a_text & "' to Long."
    Else
        On Error Resume Next
        p_value = VBA.CInt(VBA.Val(a_text))
        If Err.Number = 0 Then
            a_value = p_value
            p_parsed = True
        Else
            p_details = "Error " & VBA.CStr(Err.Number) & ", " & Err.Description & _
                " converting '" & a_text & "' to Long using VBA.CInt(VBA.Val(" & a_text & "))."
            p_parsed = False
        End If
        On Error GoTo 0
    End If

    If Not p_parsed Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & "; " & p_details

    TryParseLong = p_parsed

End Function

''' <summary>   Tries to parse a number to Single. </summary>
''' <param name="a_text">      [String] The text to parse. </param>
''' <param name="a_value">     [Single] The parsed value. </param>
''' <param name="a_details">   [String] details why parsing failed. </param>
''' <returns>   [Boolean] True if the text was parsed. <returns>
Public Function TryParseSingle(ByVal a_text As String, ByRef a_value As Single, ByVal a_details As String) As Boolean
    
    Const p_procedureName As String = "TryParseSingle"
    
    Dim p_parsed As Boolean: p_parsed = False
    Dim p_details As String: p_details = VBA.vbNullString: a_details = VBA.vbNullString
    Dim p_value As Single
    
    ' default to leave the value unchangd.
    If VBA.vbNullString = a_text Then
        p_details = "Unable to parse an empty string '" & a_text & "' to Single."
    ElseIf Not VBA.IsNumeric(a_text) Then
        p_details = "Unable to parse a non-numeric value '" & a_text & "' to Single."
    Else
        On Error Resume Next
        p_value = VBA.CSng(VBA.Val(a_text))
        If Err.Number = 0 Then
            a_value = p_value
            p_parsed = True
        Else
            p_details = "Error " & VBA.CStr(Err.Number) & ", " & Err.Description & _
                " converting '" & a_text & "' to Single using VBA.CDbl(VBA.Val(" & a_text & "))."
            p_parsed = False
        End If
        On Error GoTo 0
    End If

    If Not p_parsed Then _
        a_details = ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName & "; " & p_details

    TryParseSingle = p_parsed

End Function


