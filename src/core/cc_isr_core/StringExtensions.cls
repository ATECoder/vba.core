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
' Dependencies:
'
' None
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

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

''' <summary>   Checks whether the given text starts with the given sequence. </summary>
''' <remarks>   </remarks>
''' <param name="a_text">            [String] The text to check for the sequence. </param>
''' <param name="a_startText">       [String] The text to be located at the start. </param>
''' <param name="a_compareMethod">   [Optional, VbCompareMethod, vbBinaryCompare] The compare method. </param>
''' <returns>   A Boolean <returns>
Public Function StartsWith(ByVal a_text As String, ByVal a_startText As String, _
                           Optional ByVal a_compareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
    
    StartsWith = 1 = InStr(1, a_text, a_startText, a_compareMethod)

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

''' <summary>   Gets the character at the given index from the given string. </summary>
''' <remarks>   </remarks>
''' <param name="a_text">    [String] The text to get the character from. </param>
''' <param name="a_index">   [Integer] The index of the character to get. </param>
''' <returns>   [String]. <returns>
Public Function CharAt(ByVal a_text As String, ByVal a_index As Integer) As String
    
    CharAt = Mid(a_text, a_index, 1)

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

''' <summary>   Converts a regular string to a Unicode string,
''' e.g., "foo" to a L"foo" string.. </summary>
''' <param name="a_text">   [String] The string to convert. </param>
''' <returns>   [String]. <returns>
Public Function StringToWideString(ByVal a_text As String) As String
    
    StringToWideString = StrConv(a_text, vbUnicode)

End Function

''' <summary>   Returns true if the value is null or empty. </summary>
''' <param name="a_text">   [String] The string to check. </param>
''' <returns>   true if the value is null or empty. </returns>
Public Function IsNullOrEmpty(ByVal a_text As String) As Boolean
    IsNullOrEmpty = (a_text = vbNullString)
End Function

''' <summary>   Replaces braced numerical value with tokens similar to VB.NET String.Format
''' by without formatting suffixes. </summary>
''' <param name="a_format">   [String] The format string 'A{0} b{1}'. </param>
''' <param name="a_tokens">   [Variant, ParamArray] The parameter list matching the mask. </param>
Public Function StringFormat(ByVal a_format As String, ParamArray a_tokens() As Variant) As String
    
    Dim i As Long
    
    If VBA.vbNullString = a_format Then
        
        UserDefinedErrors.RaiseError cc_isr_Core.UserDefinedErrors.NullArgumentError, _
                    "cc_isr_Core.StringExtensions.StringFormat", _
                    "Argument name: 'a_format'."
    
    Else
    
        ' unbox the paramter array in case it came from another
        ' function which argument too is a paramter array.
    
        Dim p_tokens As Variant
        p_tokens = CoreExtensions.UnboxParameterArray(a_tokens)
        
        For i = LBound(p_tokens) To UBound(p_tokens)
            a_format = Replace$(a_format, "{" & i & "}", p_tokens(i))
        Next
        StringFormat = a_format
    
    End If

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

