Attribute VB_Name = "MarshalTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Marshal methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Tests converting an int8 to a big-endian byte string
''' and back from a big-endian byte string to an int8. </summary>
Public Function TestShouldMarshalInt8() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_value As Byte: p_value = 10
   
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_value, _
                                    Marshal.BytesToInt8(Marshal.Int8ToBytes(p_value)), _
                                    "marshals int8")

    Debug.Print p_outcome.BuildReport("TestShouldMarshalInt8")
    
    Set TestShouldMarshalInt8 = p_outcome

End Function

''' <summary>   Tests converting an int16 to a big-endian byte string
''' and back from a big-endian byte string to an int16. </summary>
Public Function TestShouldMarshalInt16() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_value As Long: p_value = 10
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_value, _
                                    Marshal.BytesToInt16(Marshal.Int16ToBytes(p_value)), _
                                    "marshals int16")

    Debug.Print p_outcome.BuildReport("TestShouldMarshalInt16")
    Set TestShouldMarshalInt16 = p_outcome

End Function

''' <summary>   Tests converting an int32 to a big-endian byte string
''' and back from a big-endian byte string to an int32. </summary>
Public Function TestShouldMarshalInt32() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_value As Long: p_value = 10
    
    Set p_outcome = cc_isr_Test_Fx.Assert.AreEqual(p_value, _
                                 Marshal.BytesToInt32(Marshal.Int32ToBytes(p_value)), _
                                 "marshals int32")

    Debug.Print p_outcome.BuildReport("TestShouldMarshalInt32")

    Set TestShouldMarshalInt32 = p_outcome

End Function
