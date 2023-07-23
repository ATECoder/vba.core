VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' StringExtensions.cls
'
' Description: Implements string extension methods.
'
' Dependencies:
'
' None
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Type ThisType
    PaddingCharacter As String
    EscapeSequences As Collection
End Type

Private This As ThisType

Option Explicit

''' <summary>   Gets the character at the given index from the given string. </summary>
''' <remarks>   </remarks>
''' <param name="a_text">    [String] The text to get the character from. </param>
''' <param name="a_index">   [Integer] The index of the character to get. </param>
''' <returns>   [String]. <returns>
Public Function CharAt(ByVal a_text As String, ByVal a_index As Integer) As String
    
    CharAt = Mid(a_text, a_index, 1)

End Function

''' <summary>   Checks whether the given text ends with the given sequence. </summary>
''' <remarks>   </remarks>
''' <param name="a_text">            [String] The text to check for the sequence. </param>
''' <param name="a_endText">         [String] The text to be located at the end. </param>
''' <param name="a_compareMethod">   [Optional, VbCompareMethod, vbBinaryCompare] The compare method. </param>
''' <returns>   A Boolean. <returns>
Public Function EndsWith(ByVal a_text As String, ByVal a_endText As String, _
                           Optional ByVal a_compareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
    
    EndsWith = 0 = StrComp(Right(a_text, Len(a_endText)), a_endText, a_compareMethod)

End Function

''' <summary>   Creates a new escape sequence. </summary>
''' <param name="a_value">         [String] The value., e.g., '/n'. </param>
''' <param name="a_replacement">   [String] The replacement value, e.g., <see cref="VBA.VbNewLine"/>. </param>
''' <returns>   [<see cref=""EscapeSequence/>]. </returns>
Public Function EscapeSequence(ByVal a_value As String, ByVal a_replacement As String) As EscapeSequence
    Dim p_escapeSequence As New EscapeSequence
    p_escapeSequence.Value = a_value
    p_escapeSequence.Replacement = a_replacement
    Set EscapeSequence = p_escapeSequence
End Function

''' <summary>   Returns an array of <see cref="EscapeSequence"/> types. </summary>
''' <value>   An Array of <see cref="EscapeSequence"/> types. </summary>
Public Function EscapeSequences() As Collection

    Static p_escapes As Collection
    
    If p_escapes Is Nothing Then
        Dim p_value As String
        Set p_escapes = New Collection
        p_value = "\n": p_escapes.Add Me.EscapeSequence(p_value, VBA.vbNewLine), p_value ' new line
        p_value = "\q": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(34)), p_value ' quote
        p_value = "\t": p_escapes.Add Me.EscapeSequence(p_value, VBA.vbTab), p_value    ' tab
        p_value = "\a": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(7)), p_value  ' alert, bell
        p_value = "\b": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(8)), p_value  ' backspace
        p_value = "\v": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(11)), p_value ' vertical tab
        p_value = "\f": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(12)), p_value ' From Feed
        p_value = "\r": p_escapes.Add Me.EscapeSequence(p_value, VBA.Chr$(13)), p_value ' CR
    End If
        
    Set EscapeSequences = p_escapes
    
End Function

''' <remarks>   Inserts or adds a string to an existing string at the specified position. </remarks>
''' <param name="a_original">   [String] The original string. </param>
''' <param name="a_added">      [String] The string to insert or add. </param>
''' <param name="a_position">   [Long] The location for inserting the added string. </param>
''' <returns>   [String]. <returns>
Function Insert(ByVal a_original As String, a_added As String, a_position As Long) As String

    If a_position < 1 Then a_position = 1
    If Len(a_original) < a_position Then a_position = Len(a_original) + 1

    Insert = Mid(a_original, 1, a_position - 1) _
                        & a_added _
                        & Mid(a_original, a_position, Len(a_original) - a_position + 1)

End Function

''' <summary>   Returns true if the value is null or empty. </summary>
''' <param name="a_text">   [String] The string to check. </param>
''' <returns>   true if the value is null or empty. </returns>
Public Function IsNullOrEmpty(ByVal a_text As String) As Boolean
    IsNullOrEmpty = (a_text = vbNullString)
End Function

''' <summary>   Pops an element from a delimited string. </summary>
''' <remarks>
''' Use this function to extract a substring from a string up to but not including the delimiter.
''' </remarks>
''' <param name="a_delimitedString">   [String] A series of tokens where a remainder string is returned. </param>
''' <param name="a_delimiter">         [String] Specifies the delimiter to use for returning the token. </param>
''' <returns>   [String]. The first sub string ending but not including the the delimiter or the
''' original string if the delimiter was not located. </returns>
Public Function Pop(ByRef a_delimitedString As String, ByVal a_delimiter As String) As String

  ' find the delimiter in the string
  
  Dim p_position As Long: p_position = InStr(1, a_delimitedString, a_delimiter, vbTextCompare)
  
  ' check if we have a delimiter in the string
  If p_position > 0 Then
  
    ' if we have a delimiter, return the string all the
    ' way up to but not including the delimiter
    Pop = Left$(a_delimitedString, p_position - 1)
    
    ' return the string past the delimiter
    a_delimitedString = Mid$(a_delimitedString, p_position + Len(a_delimiter))
    
  Else
    ' if the string does not include the delimiter, we
    ' must return the original string
    Pop = a_delimitedString
    
    ' and empty the returned string
    a_delimitedString = Empty
    
  End If
  
End Function

''' <summary>   Repeats the given string the given amount of times. </summary>
''' <remarks>   </remarks>
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
''' <remarks>   </remarks>
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
    If a_caseSensitive Then
        For i = LBound(find_values) To UBound(find_values)
            p_found = 0 <> InStr(1, string_source, find_values(i), vbBinaryCompare)
            If p_found Then Exit For
        Next
    Else
        For i = LBound(find_values) To UBound(find_values)
            StringContainsAny = 0 <> InStr(1, LCase$(string_source), LCase$(find_values(i)), vbBinaryCompare)
            If p_found Then Exit For
        Next
    End If
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
        
        UserDefinedErrors.RaiseError cc_isr_Core.UserDefinedErrors.NullArgumentError, _
                    "cc_isr_Core.StringExtensions.StringFormat", _
                    "Argument name: 'a_format'."
    
    Else
    
        ' unbox the paramter array in case it came from another
        ' function which argument too is a paramter array.
    
        Dim p_tokens As Variant
        p_tokens = CoreExtensions.UnboxParameterArray(a_values)
        
        For i = LBound(p_tokens) To UBound(p_tokens)
            a_format = Replace$(a_format, "{" & i & "}", p_tokens(i))
        Next
        StringFormatReplace = a_format
    
    End If

End Function

''' <summary>   Implements the .NET String.Format(). </summary>
''' <remarks>   <see href="https://stackoverflow.com/questions/14534360/implementing-string-format-in-vb6"/></remarks>
''' <param name="a_format">   [String] The format string. </param>
''' <param name="a_p_values">   [Variant, ParamArray] The parameter p_values matching the
'''                           <ParamRef name="a_format"/>. </param>
Public Function StringFormat(a_format As String, ParamArray a_p_values()) As String

    Const p_procedureName As String = "StringFormat"

    Dim p_returnValue As String
    Dim p_valuesCount As Integer

    If VBA.vbNullString = a_format Then _
        UserDefinedErrors.RaiseError cc_isr_Core.UserDefinedErrors.NullArgumentError, _
                    "cc_isr_Core.StringExtensions.StringFormat", _
                    "Argument name: 'a_format'."
    
    ' unbox the paramter array in case it came from another
    ' function which argument too is a paramter array.

    Dim p_values As Variant
    p_values = CoreExtensions.UnboxParameterArray(a_p_values)
    

    ' use SPACE as default padding character
    If This.PaddingCharacter = VBA.vbNullString Then This.PaddingCharacter = Chr$(32)

    'figure out number of passed p_values:
    p_valuesCount = UBound(p_values) + 1

    Dim regex As RegExp
    Dim matches As MatchCollection
    Dim thisMatch As Match
    Dim thisString As String
    Dim thisFormat As String

    ' when a_format starts with "@", escapes are not replaced
    ' (string is treated as a literal string with placeholders)
    
    Dim useLiteral As Boolean
    Dim escapeHex As Boolean 'indicates whether HEX specifier "0x" is to be escaped or not
    
    ' validate a_format:
    
    Set regex = New RegExp
    regex.Pattern = "{({{)*(\w+)(,-?\d+)?(:[^}]+)?}(}})*"
    regex.IgnoreCase = True
    regex.Global = True
    Set matches = regex.Execute(a_format)

    ' determine if p_valuesCount matches number of unique regex matches:
    Dim uniqueCount As Integer
    Dim tmpCSV As String
    For Each thisMatch In matches
        If Not Me.StringContains(tmpCSV, thisMatch.SubMatches(1)) Then
            uniqueCount = uniqueCount + 1
            tmpCSV = tmpCSV & thisMatch.SubMatches(1) & ","
        End If
    Next

    ' unique indices count must match p_valuesCount:
    
    If matches.count > 0 And uniqueCount <> p_valuesCount Then _
        cc_isr_Core.UserDefinedErrors.RaiseError cc_isr_Core.UserDefinedErrors.FormatError, _
                TypeName(Me) & "." & p_procedureName & _
                " Unique count " & CStr(uniqueCount) & " not equal the number of values of " & CStr(p_valuesCount) & "."
                
    useLiteral = Me.StartsWith(a_format, "@")

    'remove the "@" literal specifier
    If useLiteral Then a_format = Right$(a_format, Len(a_format) - 1)

    If Not useLiteral And Me.StringContains(a_format, "\\") Then _
        a_format = Replace(a_format, "\\", Chr$(27))

    If Me.StringContains(a_format, "\\") Then _
        a_format = Replace(a_format, "\\", Chr$(27))

    If matches.count = 0 And a_format <> vbNullString And UBound(p_values) = -1 Then
        
        ' only a_format was specified: skip to checking escape sequences:
        
        p_returnValue = a_format
        GoTo checkEscapes
        
    ElseIf UBound(p_values) = -1 And matches.count > 0 Then
       
        cc_isr_Core.UserDefinedErrors.RaiseError cc_isr_Core.UserDefinedErrors.NullArgumentError, _
            TypeName(Me) & "." & p_procedureName, " Argument name: 'p_values'."
    
    End If

    p_returnValue = a_format

    'dissect a_format:

    Dim i As Integer, v As String, p As String 'i: iterator; v: value; p: placeholder
    Dim alignmentGroup As String, alignmentSpecifier As String
    Dim formattedValue As String, alignmentPadding As Integer

    ' iterate regex matches (each match is a placeholder):
    
    For i = 0 To matches.count - 1

        ' get the placeholder specified index:
        Set thisMatch = matches(i)
        p = thisMatch.SubMatches(1)

        ' if specified index (0-based) > uniqueCount (1-based), something's wrong:
        
        If p > uniqueCount - 1 Then _
            cc_isr_Core.UserDefinedErrors.RaiseError cc_isr_Core.UserDefinedErrors.FormatError, _
                TypeName(Me) & "." & p_procedureName, _
                " Specified index " & CStr(p) & " exceeds " & CStr(uniqueCount - 1) & "."
            
        v = p_values(p)

        'get the alignment specifier if it is specified:
        alignmentGroup = thisMatch.SubMatches(2)
        If alignmentGroup <> vbNullString Then _
            alignmentSpecifier = Right$(alignmentGroup, LenB(alignmentGroup) / 2 - 1)


        'get the format specifier if it is specified:
        thisString = thisMatch.Value
        If Me.StringContains(thisString, ":") Then

            Dim formatGroup As String, precisionSpecifier As Integer
            Dim formatSpecifier As String, precisionString As String

            'get the string between ":" and "}":
            formatGroup = Mid$(thisString, InStr(1, thisString, ":") + 1, (LenB(thisString) / 2) - 2)
            formatGroup = Left$(formatGroup, LenB(formatGroup) / 2 - 1)

            precisionString = Right$(formatGroup, LenB(formatGroup) / 2 - 1)
            formatSpecifier = Mid$(thisString, InStr(1, thisString, ":") + 1, 1)

            'applicable formatting depends on the type of the value (yes, GOTO!!):
            If TypeName(p_values(p)) = "Date" Then GoTo DateTimeFormatSpecifiers
            If v = vbNullString Then GoTo ApplyStringFormat

NumberFormatSpecifiers:
            If precisionString <> vbNullString And Not IsNumeric(precisionString) Then _
                cc_isr_Core.UserDefinedErrors.RaiseError cc_isr_Core.UserDefinedErrors.FormatError, _
                    TypeName(Me) & "." & p_procedureName, _
                    " Preceision string '" & precisionString & "' is not numeric."

            If precisionString = vbNullString Then precisionString = "0"

            Select Case formatSpecifier

                Case "C", "c" 'CURRENCY format, formats string as currency.
                    'Precision specifier determines number of decimal digits.
                    'This implementation ignores regional settings
                    '(hard-coded group separator, decimal separator and currency sign).

                    If precisionString = "0" Then precisionString = "2"

                    precisionSpecifier = CInt(precisionString)
                    thisFormat = "#,##0.$"

                    'if a non-zero precision is specified...
                    thisFormat = Replace$(thisFormat, ".", "." & String$(precisionSpecifier, Chr$(48)))

                Case "D", "d" 'DECIMAL format, formats string as integer number.
                    
                    'Precision specifier determines number of digits in returned string.

                    precisionSpecifier = CInt(precisionString)
                    thisFormat = "0"
                    thisFormat = Right$(String$(precisionSpecifier, "0") & thisFormat, _
                        IIf(precisionSpecifier = 0, Len(thisFormat), precisionSpecifier))


                Case "E", "e" 'EXPONENTIAL NOTATION format (aka "Scientific Notation")
                    'Precision specifier determines number of decimals in returned string.
                    'This implementation ignores regional settings'
                    '(hard-coded decimal separator).


                    precisionSpecifier = CInt(precisionString)
                    thisFormat = "0.00000#" & formatSpecifier & "-#" 'defaults to 6 decimals

                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = "0." & String$(precisionSpecifier - 1, Chr$(48)) & "#" & formatSpecifier & "-#"

                    ElseIf VBA.LenB(formatGroup) > 2 And precisionSpecifier = 0 Then
                        
                        cc_isr_Core.UserDefinedErrors.RaiseError cc_isr_Core.UserDefinedErrors.InvalidFormatStringError, _
                            TypeName(Me) & "." & p_procedureName, _
                            " Format group '" & formatGroup & "' length " & CStr(VBA.LenB(formatGroup)) & _
                            " is greater than 2 and the precision specifier is 0."
                    
                    End If


                Case "F", "f" 'FIXED-POINT format
                    
                    'Precision specifier determines number of decimals in returned string.
                    'This implementation ignores regional settings'
                    '(hard-coded decimal separator).

                    precisionSpecifier = CInt(precisionString)
                    thisFormat = "0"
                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = (thisFormat & ".") & String$(precisionSpecifier, Chr$(48))
                    Else
                        'no precision specified - default to 2 decimals:
                        thisFormat = "0.00"
                    End If


                    Case "G", "g" 'GENERAL format (recursive)
                    'returns the shortest of either FIXED-POINT or SCIENTIFIC formats in case of a Double.
                    'returns DECIMAL format in case of a Integer or Long.

                    Dim eNotation As String, ePower As Integer, specifier As String
                    precisionSpecifier = IIf(CInt(precisionString) > 0, CInt(precisionString), _
                        IIf(Me.StringContains(v, "."), Len(v) - InStr(1, v, "."), 0))

                    'track character case of formatSpecifier:
                    specifier = IIf(formatSpecifier = "G", "D", "d")

                    If TypeName(p_values(p)) = "Integer" Or TypeName(p_values(p)) = "Long" Then
                        'Integer types: use {0:D} (recursive call):
                        formattedValue = StringFormat("{0:" & specifier & "}", p_values(p))

                    ElseIf TypeName(p_values(p)) = "Double" Then
                        'Non-integer types: use {0:E}
                        specifier = IIf(formatSpecifier = "G", "E", "e")

                        'evaluate the exponential notation value (recursive call):
                        eNotation = StringFormat("{0:" & specifier & "}", v)

                        'get the power of eNotation:
                        ePower = Mid$(eNotation, InStr(1, UCase$(eNotation), "E-") + 1, Len(eNotation) - InStr(1, UCase$(eNotation), "E-"))

                        If ePower > -5 And Abs(ePower) < precisionSpecifier Then
                            'use {0:F} when ePower > -5 and abs(ePower) < precisionSpecifier:
                            'evaluate the floating-point value (recursive call):
                             specifier = IIf(formatSpecifier = "G", "F", "f")
                             formattedValue = StringFormat("{0:" & formatSpecifier & _
                                 IIf(precisionSpecifier <> 0, precisionString, vbNullString) & "}", p_values(p))
                        Else
                            'fallback to {0:E} if previous rule didn't apply:
                            formattedValue = eNotation
                        End If

                    End If

                    GoTo AlignFormattedValue 'Skip the "ApplyStringFormat" step, it's applied already.


                Case "N", "n" 'NUMERIC format, formats string as an integer or decimal number.
                    
                    'Precision specifier determines number of decimal digits.
                    'This implementation ignores regional settings'
                    '(hard-coded group and decimal separators).

                    precisionSpecifier = CInt(precisionString)
                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = "#,##0"
                        thisFormat = (thisFormat & ".") & String$(precisionSpecifier, Chr$(48))

                    Else 'only the "D" is specified
                        thisFormat = "#,##0"
                    End If

                Case "P", "p" 'PERCENT format. Formats string as a percentage.
                    
                    'Value is multiplied by 100 and displayed with a percent symbol.
                    'Precision specifier determines number of decimal digits.

                    thisFormat = "#,##0%"
                    precisionSpecifier = CInt(precisionString)
                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = "#,##0"
                        thisFormat = (thisFormat & ".") & String$(precisionSpecifier, Chr$(48))

                    Else 'only the "P" is specified
                        thisFormat = "#,##0"
                    End If

                    'Append the percentage sign to the format string:
                    thisFormat = thisFormat & "%"

                Case "R", "r" 'ROUND-TRIP format (a string that can round-trip to an identical number)
                    
                    'example: ?StringFormat("{0:R}", 0.0000000001141596325677345362656)
                    '         ...returns "0.000000000114159632567735"

                    'convert value to a Double (chop off overflow digits):
                    v = CDbl(v)

                Case "X", "x" 'HEX format. Formats a string as a Hexadecimal value.
                    
                    'Precision specifier determines number of total digits.
                    'Returned string is prefixed with "&H" to specify Hex.

                    v = Hex(v)
                    precisionSpecifier = CInt(precisionString)

                    If LenB(precisionString) > 0 Then 'precision here stands for left padding
                        v = Right$(String$(precisionSpecifier, "0") & v, IIf(precisionSpecifier = 0, Len(v), precisionSpecifier))
                    End If

                    'add C# hex specifier, apply specified casing:
                    '(VB6 hex specifier would cause Format() to reverse the formatting):
                    v = "0x" & IIf(formatSpecifier = "X", UCase$(v), LCase$(v))

                Case Else

                    If IsNumeric(formatSpecifier) And Val(formatGroup) = 0 Then
                        formatSpecifier = formatGroup
                        v = Format(v, formatGroup)
                    Else
                        cc_isr_Core.UserDefinedErrors.RaiseError cc_isr_Core.UserDefinedErrors.InvalidFormatStringError, _
                            TypeName(Me) & "." & p_procedureName, _
                            " Format specifier '" & formatSpecifier & "' is not numeric" & _
                            " or value of format group '" & Val(formatGroup) & "' is not zero."
                    End If
            End Select

            GoTo ApplyStringFormat


DateTimeFormatSpecifiers:
                
            Select Case formatSpecifier

                Case "c", "C" 'CUSTOM date/time format
                    
                    'let VB Format() parse precision specifier as is:
                    thisFormat = precisionString

                Case "d" 'SHORT DATE format
                    thisFormat = "ddddd"

                Case "D" 'LONG DATE format
                    thisFormat = "dddddd"

                Case "f" 'FULL DATE format (short)
                    thisFormat = "dddddd h:mm AM/PM"

                Case "F" 'FULL DATE format (long)
                    thisFormat = "dddddd ttttt"

                Case "g"
                    thisFormat = "ddddd hh:mm AM/PM"

                Case "G"
                    thisFormat = "ddddd ttttt"

                Case "s" 'SORTABLE DATETIME format
                    thisFormat = "yyyy-mm-ddThh:mm:ss"

                Case "t" 'SHORT TIME format
                    thisFormat = "hh:mm AM/PM"

                Case "T" 'LONG TIME format
                    thisFormat = "ttttt"

                Case Else
                    
                    cc_isr_Core.UserDefinedErrors.RaiseError cc_isr_Core.UserDefinedErrors.FormatError, _
                        TypeName(Me) & "." & p_procedureName, _
                        " Unknown format specifier '" & formatSpecifier & "'."
                
            End Select
            GoTo ApplyStringFormat

        End If

ApplyStringFormat:
            
        'apply computed format string:
        
        formattedValue = Format(v, thisFormat)

AlignFormattedValue:
        
        'apply specified alignment specifier:
        
        If alignmentSpecifier <> vbNullString Then

            alignmentPadding = Abs(CInt(alignmentSpecifier))
            If CInt(alignmentSpecifier) < 0 Then
                'negative: left-justified alignment
                
                If alignmentPadding - Len(formattedValue) > 0 Then _
                    formattedValue = formattedValue & _
                        String$(alignmentPadding - Len(formattedValue), This.PaddingCharacter)
            Else
                'positive: right-justified alignment
                
                If alignmentPadding - Len(formattedValue) > 0 Then _
                    formattedValue = String$(alignmentPadding - Len(formattedValue), This.PaddingCharacter) & formattedValue
            End If
        
        End If

        'Replace C# hex specifier with VB6 hex specifier,
        'only if hex specifier was introduced in this function:
        
        If (Not useLiteral And escapeHex) And Me.StringContains(formattedValue, "0x") Then _
            formattedValue = Replace$(formattedValue, "0x", "&H")

        'replace all occurrences of placeholder {i} with their formatted p_values:
        
        p_returnValue = Replace(p_returnValue, thisString, formattedValue, count:=1)

        'reset before reiterating:
        
        thisFormat = vbNullString
    
    Next

checkEscapes:

    ' if there's no more backslashes, don't bother checking for the rest:
    
    If useLiteral Or Not Me.StringContains(p_returnValue, "\") Then GoTo normalExit

    Dim p_escape As EscapeSequence
    Dim p_escapes As Collection: Set p_escapes = Me.EscapeSequences
    
    For Each p_escape In p_escapes
        
        If Me.StringContains(p_returnValue, p_escape.Value) Then _
            p_returnValue = Replace(p_returnValue, p_escape.Value, p_escape.Replacement)

        If Not Me.StringContains(p_returnValue, "\") Then GoTo normalExit
    
    Next

    ' replace "ASCII (oct)" escape sequence
    
    Set regex = New RegExp
    regex.Pattern = "\\(\d{3})"
    regex.IgnoreCase = True
    regex.Global = True
    Set matches = regex.Execute(a_format)

    Dim char As Long
    If matches.count <> 0 Then
        For Each thisMatch In matches
            p = thisMatch.SubMatches(0)
            '"p" contains the octal number representing the ASCII code we're after:
            p = "&O" & p 'prepend octal prefix
            char = CLng(p)
            p_returnValue = Replace(p_returnValue, thisMatch.Value, Chr$(char))
        Next
    End If

    ' if there's no more backslashes, don't bother checking for the rest:
    If Not Me.StringContains(p_returnValue, "\") Then GoTo normalExit

    'replace "ASCII (hex)" escape sequence
    Set regex = New RegExp
    regex.Pattern = "\\x(\w{2})"
    regex.IgnoreCase = True
    regex.Global = True
    Set matches = regex.Execute(a_format)

    If matches.count <> 0 Then
        For Each thisMatch In matches
            p = thisMatch.SubMatches(0)
            '"p" contains the hex value representing the ASCII code we're after:
            p = "&H" & p 'prepend hex prefix
            char = CLng(p)
            p_returnValue = Replace(p_returnValue, thisMatch.Value, Chr$(char))
        Next
    End If

normalExit:
    If Not useLiteral And Me.StringContains(p_returnValue, Chr$(27)) Then _
        p_returnValue = Replace(p_returnValue, Chr$(27), "\")
    StringFormat = p_returnValue
        
End Function


''' <summary>   Converts a regular string to a Unicode string,
''' e.g., "foo" to a L"foo" string.. </summary>
''' <param name="a_text">   [String] The string to convert. </param>
''' <returns>   [String]. <returns>
Public Function StringToWideString(ByVal a_text As String) As String
    
    StringToWideString = StrConv(a_text, vbUnicode)

End Function

''' <summary>   Checks whether the given text starts with the given sequence. </summary>
''' <remarks>   </remarks>
''' <param name="a_text">            [String] The text to check for the start. </param>
''' <param name="a_startText">       [String] The text to be located at the start. </param>
''' <param name="a_compareMethod">   [Optional, VbCompareMethod, vbBinaryCompare] The compare method. </param>
''' <returns>   A Boolean <returns>
Public Function StartsWith(ByVal a_text As String, ByVal a_startText As String, _
                           Optional ByVal a_compareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
    
    StartsWith = 1 = InStr(1, a_text, a_startText, a_compareMethod)

End Function

''' <summary>   Gets the substring from the given text. </summary>
''' <remarks>   </remarks>
''' <param name="a_text">         [String] The text to get the substring from. </param>
''' <param name="a_startIndex">   [Integer] The index of the first character of the substring. </param>
''' <param name="a_length">       [Optional, null] The number of characters to take from the original string. </param>
''' <returns>   [String]. <returns>
Public Function Substring(ByVal a_text As String, ByVal a_startIndex As Integer, _
        Optional ByVal a_length As Variant) As String
    
    If a_startIndex > Len(a_text) Then
        a_startIndex = Len(a_text)
    End If

    If IsMissing(a_length) Then
        a_length = Len(a_text) - a_startIndex
    End If
    
    If a_length > Len(a_text) Then
        a_length = Len(a_text) - a_startIndex
    End If

    Substring = Left(Right(a_text, Len(a_text) - a_startIndex), a_length)
    
End Function

''' <remarks>   Trims the given character from the start of the text. </remarks>
''' <param name="a_text">   [String] The text to trim. </param>
''' <param name="a_char">   [String] The character to trim. </param>
''' <returns>   [String]. <returns>
Public Function TrimLeft(ByVal a_text As String, ByVal a_char As String) As String
    
    Dim p_textLength As Long
    p_textLength = Len(a_text)
    
    Dim p_firstCharacter As String
    
    Do While p_textLength > 0
        p_firstCharacter = Left(a_text, 1)
        
        If p_firstCharacter <> a_char Then
            Exit Do
        End If
        
        a_text = Right(a_text, p_textLength - 1)
        p_textLength = Len(a_text)
    Loop
    
    TrimLeft = a_text
End Function

''' <summary>   Trims the given character from the end of the given text. </summary>
''' <param name="a_text">   [String] The text to trim. </param>
''' <param name="a_char">   [String] The character to trim. </param>
''' <returns>   [String]. <returns>
Public Function TrimRight(ByVal a_text As String, ByVal a_char As String) As String
    
    Dim p_textLength As Long: p_textLength = Len(a_text)
    
    Dim p_lastCharacter As String
        
    Do While p_textLength > 0
        p_lastCharacter = Right(a_text, 1)
        
        If p_lastCharacter <> a_char Then
            Exit Do
        End If
        
        a_text = Left(a_text, p_textLength - 1)
        p_textLength = Len(a_text)
    Loop
    
    TrimRight = a_text
End Function

