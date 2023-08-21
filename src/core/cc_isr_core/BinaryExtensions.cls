VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinaryExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Binary extensions. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Convert a fractional decimal value to a binary string. </summary>
''' <param name="a_value">      [Long] The value to convert. </param>
''' <param name="a_bitCount">   [Integer] the number of bits. </param>
''' <returns>   [String]. <returns>
Public Function FractionalToBinary(ByVal a_value As Double, a_bitCount As Integer) As String
    
    Const p_procedureName As String = "FractionalToBinary"
    
    If a_value < 0 Then _
        cc_isr_Core_IO.UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.InvalidArgumentError, _
                ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "." & p_procedureName, _
                " Argument: a_value = '" & VBA.CStr(a_value) & "' but must be non-negative."
            
    Dim p_binary As String
    
    Do While a_value <> 0 And VBA.Len(p_binary) < a_bitCount
        a_value = a_value * 2
        p_binary = p_binary & IIf(a_value >= 1, "1", "0")
        If a_value >= 1 Then
            a_value = a_value - 1
        End If
    Loop

    ' Pad with trailing zeros if necessary
    While VBA.Len(p_binary) < a_bitCount
        p_binary = p_binary & "0"
    Wend

    FractionalToBinary = p_binary

End Function

''' <summary>   Converts a double value to a binary string. </summary>
''' <param name="a_value">      [Long] The value to convert. </param>
''' <param name="a_bitCount">   [Integer] the number of bits. </param>
''' <returns>   [String]. <returns>
Function DoubleToBinary(ByVal a_value As Double, ByVal a_bitCount As Integer) As String

    Dim p_fractionalBinary As String
    Dim p_integralBinary As String
    
    If a_value = 0 Then
        
        p_integralBinary = VBA.String$(a_bitCount, "0")
        p_fractionalBinary = VBA.String$(a_bitCount, "0")
    
    Else
        
        Dim p_integralValue As Long
        p_integralValue = Int(a_value)
        
        Dim p_franctionalValue As Double
        p_franctionalValue = a_value - p_integralValue
        
        p_integralBinary = Me.LongToBinary(p_integralValue, a_bitCount)
        p_fractionalBinary = Me.FractionalToBinary(p_franctionalValue, a_bitCount)

    End If

    DoubleToBinary = p_integralBinary & "." & p_fractionalBinary

End Function

''' <summary>   Converts a Long value to a binary string. </summary>
''' <param name="a_value">      [Long] The value to convert. </param>
''' <param name="a_bitCount">   [Integer] the number of bits. </param>
''' <returns>   [String]. <returns>
Public Function LongToBinary(ByVal a_value As Long, ByVal a_bitCount As Integer) As String

    Dim p_binary As String
    Dim p_remainder As Integer
    Dim p_isNegative As Boolean
    
    If a_value = 0 Then
        LongToBinary = String$(a_bitCount, "0")
    Else
        If a_value < 0 Then
            p_isNegative = True
            a_value = -a_value
        End If
        
        Do While a_value > 0
            p_remainder = a_value Mod 2
            a_value = a_value \ 2
            p_binary = VBA.CStr(p_remainder) & p_binary
        Loop
        
        ' add leading zeros
        p_binary = Me.AddLeadingZeros(p_binary, a_bitCount)
        
        If p_isNegative Then
            
            p_binary = Me.InvertBits(p_binary)
            
            ' convert to two's complement.
            p_binary = Me.AddBinary(p_binary, "1")
        
        End If
        
        LongToBinary = p_binary
    End If
    
End Function

''' <summary>   Adds a leading zeros. </summary>
''' <param name="a_value">        [String] The initial value. </param>
''' <param name="a_finalLength">  [Integer] the final length of the outcome string. </param>
''' <returns>   [String]. <returns>
Public Function AddLeadingZeros(ByVal a_value As String, ByVal a_finalLength As Integer) As String

    Dim p_difference As Integer
    p_difference = a_finalLength - VBA.Len(a_value)
    
    If p_difference > 0 Then
        AddLeadingZeros = VBA.String$(p_difference, "0") & a_value
    Else
        AddLeadingZeros = a_value
    End If
    
End Function

''' <summary>   Inverts the bits of a binary string. </summary>
''' <param name="a_value">   [String] The initial value. </param>
''' <returns>   [String]. <returns>
Public Function InvertBits(ByVal a_value As String) As String

    Dim i As Integer
    Dim p_inverted As String
    For i = 1 To Len(a_value)
        p_inverted = p_inverted & IIf(VBA.Mid$(a_value, i, 1) = "0", "1", "0")
    Next i
    InvertBits = p_inverted
    
End Function

''' <summary>   Adds two binary values. </summary>
''' <param name="a_binary1">   [String] A binary string. </param>
''' <param name="a_binary2">   [String] A binary string. </param>
''' <returns>   [String]. <returns>
Public Function AddBinary(ByVal a_binary1 As String, ByVal a_binary2 As String) As String

    Dim p_maxLength As Integer
    p_maxLength = IIf(Len(a_binary1) > Len(a_binary2), Len(a_binary1), Len(a_binary2))
    a_binary1 = Me.AddLeadingZeros(a_binary1, p_maxLength)
    a_binary2 = Me.AddLeadingZeros(a_binary2, p_maxLength)

    Dim p_carry As Integer
    Dim p_result As String
    Dim i As Integer

    p_carry = 0
    For i = p_maxLength To 1 Step -1
    
        Dim p_bit1 As Integer, p_bit2 As Integer, p_sum As Integer
        p_bit1 = CInt(Mid(a_binary1, i, 1))
        p_bit2 = CInt(Mid(a_binary2, i, 1))
        
        p_sum = p_bit1 + p_bit2 + p_carry
        p_result = VBA.CStr(p_sum Mod 2) & p_result
        p_carry = p_sum \ 2
    
    Next i

    If p_carry > 0 Then
        p_result = VBA.CStr(p_carry) & p_result
    End If

    AddBinary = p_result
    
End Function

