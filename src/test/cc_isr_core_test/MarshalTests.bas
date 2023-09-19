Attribute VB_Name = "MarshalTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Marshal methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Name As String
    TestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    ErrTracer As IErrTracer
    TestCount As Integer
    RunCount As Integer
    PassedCount As Integer
    FailedCount As Integer
    InconclusiveCount As Integer
End Type

Private This As this_

''' <summary>   Runs the specified test. </summary>
Public Function RunTest(ByVal a_testNumber As Integer) As cc_isr_Test_Fx.Assert
    Dim p_outcome As cc_isr_Test_Fx.Assert
    'BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestShouldMarshalInt8
        Case 2
            Set p_outcome = TestShouldMarshalInt16
        Case 3
            Set p_outcome = TestShouldMarshalInt32
        Case Else
    End Select
    Set RunTest = p_outcome
    'AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    'BeforeAll
    RunTest 1
    'AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    This.Name = "MarshalTests"
    'BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 3
    Dim p_testNumber As Integer
    For p_testNumber = 1 To This.TestCount
        Set p_outcome = RunTest(p_testNumber)
        If Not p_outcome Is Nothing Then
            This.RunCount = This.RunCount + 1
            If p_outcome.AssertInconclusive Then
                This.InconclusiveCount = This.InconclusiveCount + 1
            ElseIf p_outcome.AssertSuccessful Then
                This.PassedCount = This.PassedCount + 1
            Else
                This.FailedCount = This.FailedCount + 1
            End If
        End If
        DoEvents
    Next p_testNumber
    'AfterAll
    Debug.Print "Ran " & VBA.CStr(This.RunCount) & " out of " & VBA.CStr(This.TestCount) & " tests."
    Debug.Print "Passed: " & VBA.CStr(This.PassedCount) & "; Failed: " & VBA.CStr(This.FailedCount) & _
                "; Inconclusive: " & VBA.CStr(This.InconclusiveCount) & "."
End Sub


''' <summary>   Tests converting an int8 to a big-endian byte string
''' and back from a big-endian byte string to an int8. </summary>
Public Function TestShouldMarshalInt8() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_value As Byte: p_value = 10
   
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_value, _
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
    
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_value, _
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
    
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_value, _
                                 Marshal.BytesToInt32(Marshal.Int32ToBytes(p_value)), _
                                 "marshals int32")

    Debug.Print p_outcome.BuildReport("TestShouldMarshalInt32")

    Set TestShouldMarshalInt32 = p_outcome

End Function
