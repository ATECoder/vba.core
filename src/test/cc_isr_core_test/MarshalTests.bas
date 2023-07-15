Attribute VB_Name = "MarshalTests"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' MarshalTests.bas
'
' Dependencies:
'
' Assert.cls
' Marshal.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Tests converting an int8 to a big-endian byte string
''' and back from a big-endian byte string to an int8. </summary>
Public Function TestShouldMarshalInt8() As Assert
    
    Dim p_value As Byte: p_value = 10
   
    Set TestShouldMarshalInt8 = Assert.AreEqual(p_value, Marshal.BytesToInt8(Marshal.Int8ToBytes(p_value)), "marshals int8")

End Function

''' <summary>   Tests converting an int16 to a big-endian byte string
''' and back from a big-endian byte string to an int16. </summary>
Public Function TestShouldMarshalInt16() As Assert
    
    Dim p_value As Long: p_value = 10
    
    Set TestShouldMarshalInt16 = Assert.AreEqual(p_value, Marshal.BytesToInt16(Marshal.Int16ToBytes(p_value)), "marshals int16")

End Function

''' <summary>   Tests converting an int32 to a big-endian byte string
''' and back from a big-endian byte string to an int32. </summary>
Public Function TestShouldMarshalInt32() As Assert
    
    Dim p_value As Long: p_value = 10
    
    Set TestShouldMarshalInt32 = Assert.AreEqual(p_value, Marshal.BytesToInt32(Marshal.Int32ToBytes(p_value)), "marshals int32")

End Function
