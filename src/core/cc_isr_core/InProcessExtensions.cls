VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InProcessExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   In Process extensions. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Function DecimalToBinary(decimalNum As Double, bitCount As Integer) As String
    Dim binaryString As String
    Dim isNegative As Boolean

    If decimalNum = 0 Then
        DecimalToBinary = AddLeadingZeros("0", bitCount)
    Else
        If decimalNum < 0 Then
            isNegative = True
            decimalNum = Abs(decimalNum)
        End If

        Dim intValue As Long
        intValue = Int(decimalNum)

        Do While intValue > 0
            Dim remainder As Integer
            remainder = intValue Mod 2
            intValue = intValue \ 2
            binaryString = CStr(remainder) & binaryString
        Loop

                Dim bit As Integer

        If isNegative Then
            ' Convert the fractional part to binary
            Dim fractionalPart As Double
            fractionalPart = decimalNum - Int(decimalNum)

            binaryString = AddLeadingZeros(binaryString, bitCount) ' Pad with leading zeros to bitCount

            ' Convert the fractional part and add it to the binary string
            fractionalPart = fractionalPart * 2 ' Move fractional part to the integer part
            Do While Len(binaryString) < bitCount And fractionalPart > 0
                bit = Int(fractionalPart)
                binaryString = binaryString & CStr(bit)
                fractionalPart = fractionalPart * 2 - bit ' Update fractional part for the next bit
            Loop

            binaryString = InvertBits(binaryString) ' Flip the bits
            binaryString = AddBinary(binaryString, "1") ' Add 1 for Two's complement
        Else
            ' Pad with leading zeros to bitCount, including fractional part
            Do While Len(binaryString) < bitCount
                decimalNum = decimalNum * 2
                bit = Int(decimalNum)
                binaryString = binaryString & CStr(bit)
                decimalNum = decimalNum - bit ' Update decimal number for the next bit
            Loop
        End If

        DecimalToBinary = binaryString
    End If
End Function

' ... (rest of the code remains unchanged)

Function ConvertFractionalPart(fractionalPart As Double, maxDigits As Integer) As String
    Dim fractionalBinary As String
    Do While fractionalPart <> 0 And Len(fractionalBinary) < maxDigits
        fractionalPart = fractionalPart * 2
        fractionalBinary = fractionalBinary & IIf(fractionalPart >= 1, "1", "0")
        If fractionalPart >= 1 Then
            fractionalPart = fractionalPart - 1
        End If
    Loop

    ' Pad with trailing zeros if necessary
    While Len(fractionalBinary) < maxDigits
        fractionalBinary = fractionalBinary & "0"
    Wend

    ConvertFractionalPart = fractionalBinary
End Function

' ... (rest of the code remains unchanged)

Function AddLeadingZeros(binaryStr As String, length As Integer) As String
    Dim diff As Integer
    diff = length - Len(binaryStr)

    If diff > 0 Then
        AddLeadingZeros = String(diff, "0") & binaryStr
    Else
        AddLeadingZeros = binaryStr
    End If
End Function

Function InvertBits(binaryStr As String) As String
    Dim i As Integer
    Dim invertedStr As String
    For i = 1 To Len(binaryStr)
        invertedStr = invertedStr & IIf(Mid(binaryStr, i, 1) = "0", "1", "0")
    Next i
    InvertBits = invertedStr
End Function

Function AddBinary(binaryStr As String, binaryStr2 As String) As String
    Dim maxLen As Integer
    maxLen = IIf(Len(binaryStr) > Len(binaryStr2), Len(binaryStr), Len(binaryStr2))
    binaryStr = AddLeadingZeros(binaryStr, maxLen)
    binaryStr2 = AddLeadingZeros(binaryStr2, maxLen)

    Dim carry As Integer
    Dim result As String
    Dim i As Integer

    carry = 0
    For i = maxLen To 1 Step -1
        Dim bit1 As Integer, bit2 As Integer
        bit1 = CInt(Mid(binaryStr, i, 1))
        bit2 = CInt(Mid(binaryStr2, i, 1))
        Dim sum As Integer
        sum = bit1 + bit2 + carry
        result = CStr(sum Mod 2) & result
        carry = sum \ 2
    Next i

    If carry > 0 Then
        result = CStr(carry) & result
    End If

    AddBinary = result
End Function

