VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Marshal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Byte and endianess extensions. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const integer0xFF As Integer = 256
Private Const long0xFF As Long = 256

''' <summary>   Converts the given int8 to a big-endian byte string. </summary>
''' <param name="a_value">   [Byte] The byte to convert to bytes. </param>
''' <returns>   [String]. </returns>
Public Function Int8ToBytes(ByVal a_value As Byte) As String
    
    Int8ToBytes = VBA.Chr$(a_value)

End Function

''' <summary>   Converts the given big-endian byte string to an int8. </summary>
''' <param name="a_bytes">   [String] The bytes to convert. </param>
''' <returns>   [Byte]. </returns>
Public Function BytesToInt8(ByVal a_bytes As String) As Byte
    
    BytesToInt8 = VBA.Asc(StringExtensions.CharAt(a_bytes, 1))

End Function

''' <summary>   Converts the given int16 to a big-endian byte string. </summary>
''' <param name="a_value">   [Integer] The value to convert. </param>
''' <returns>   [String]. </returns>
Public Function Int16ToBytes(ByVal a_value As Integer) As String
    
    Dim p_bytes As String * 2
    Dim p_rest As Integer

    ' Cannot use VBA.Mid here!
    
    p_rest = a_value Mod integer0xFF
    Mid(p_bytes, 2) = VBA.Chr$(p_rest)

    a_value = (a_value - p_rest) / integer0xFF
    p_rest = a_value Mod integer0xFF
    Mid(p_bytes, 1) = VBA.Chr$(p_rest)

    Int16ToBytes = p_bytes

End Function

''' <summary>   Converts the given big-endian byte string to an int16. </summary>
''' <param name="a_bytes">   [String] The bytes to convert. </param>
''' <returns>   [Long]. </returns>
Public Function BytesToInt16(ByVal a_bytes As String) As Long
    
    BytesToInt16 = VBA.Asc(StringExtensions.CharAt(a_bytes, 1)) * long0xFF + _
                   VBA.Asc(StringExtensions.CharAt(a_bytes, 2))

End Function

''' <summary>   Converts the given int32 to a big-endian byte string. </summary>
''' <param name="a_value">   [Long] The byte to convert. </param>
''' <returns>   </returns>
Public Function Int32ToBytes(ByVal a_value As Long) As String
    
    Dim p_bytes As String * 4
    Dim p_rest As Long

    ' Cannot use VBA.Mid here!

    p_rest = a_value Mod long0xFF
    Mid(p_bytes, 4) = VBA.Chr$(p_rest)

    a_value = (a_value - p_rest) / long0xFF
    p_rest = a_value Mod long0xFF
    Mid(p_bytes, 3) = VBA.Chr$(p_rest)

    a_value = (a_value - p_rest) / long0xFF
    p_rest = a_value Mod long0xFF
    Mid(p_bytes, 2) = VBA.Chr$(p_rest)

    a_value = (a_value - p_rest) / long0xFF
    p_rest = a_value Mod long0xFF
    Mid(p_bytes, 1) = VBA.Chr$(p_rest)

    Int32ToBytes = p_bytes

End Function

''' <summary>   Converts the given big-endian byte string to an int32. </summary>
''' <param name="a_bytes">   [String] The bytes to convert. </param>
''' <returns>   [Long]. </returns>
Public Function BytesToInt32(ByVal a_bytes As String) As Long

    BytesToInt32 = VBA.Asc(StringExtensions.CharAt(a_bytes, 1)) * long0xFF * long0xFF * long0xFF + _
                   VBA.Asc(StringExtensions.CharAt(a_bytes, 2)) * long0xFF * long0xFF + _
                   VBA.Asc(StringExtensions.CharAt(a_bytes, 3)) * long0xFF + _
                   VBA.Asc(StringExtensions.CharAt(a_bytes, 4))

End Function
