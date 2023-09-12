Attribute VB_Name = "StopWatchTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Stopwatch extension methods. </summary>
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
            Set p_outcome = TestBitsShouldInvert
        Case 2
            Set p_outcome = TestTimeShouldExceedExpectedMs
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
    This.Name = "BinaryExtensionTests"
    'BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 2
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


''' <summary>   Unit test. Asserts that the <see cref="Stopwatch"/>.<see cref="Stopwatch.ElapsedMilliseconds"/>
''' exceeds the thread sleep time. </summary>
''' <returns>   An instance of the <see cref="cc_isr_Test_Fx.Assert"/>   class. </returns>
Public Function TestElapsedTimeShouldExceedExpectedMs() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_stopper As Stopwatch: Set p_stopper = cc_isr_Core_IO.Factory.NewStopWatch
    Dim p_expectedMs As Long
    p_expectedMs = 100
    p_stopper.Sleep p_expectedMs + 50
    p_stopper.StopCounter
    Dim p_actualMs As Long: p_actualMs = p_stopper.ElapsedMilliseconds
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_stopper.ElapsedMilliseconds > p_expectedMs, _
            "elapsed time " & CStr(p_stopper.ElapsedMilliseconds) & _
            " must exceed sleep time " & _
            CStr(p_expectedMs))
            
    Debug.Print p_outcome.BuildReport("TestElapsedTimeShouldExceedExpectedMs")
    
    Set TestElapsedTimeShouldExceedExpectedMs = p_outcome
        
End Function

''' <summary>   Unit test. Asserts that the <see cref="Stopwatch"/>.<see cref="Stopwatch.Wait"/>
''' exceeds the specified interval. </summary>
''' <returns>   An instance of the <see cref="cc_isr_Test_Fx.Assert"/>   class. </returns>
Public Function TestTimeShouldExceedExpectedMs() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    
    Dim p_expectedMs As Long: p_expectedMs = 100
    Dim p_stopper As Stopwatch: Set p_stopper = cc_isr_Core_IO.Factory.NewStopWatch
    Dim p_actualMs As Long: p_actualMs = p_stopper.Wait(p_expectedMs)
    
    Set p_outcome = cc_isr_Test_Fx.Assert.IsTrue(p_actualMs >= p_expectedMs, _
            "elapsed time " & CStr(p_actualMs) & " must exceed " & CStr(p_expectedMs))

    Debug.Print p_outcome.BuildReport("TestTimeShouldExceedExpectedMs")

    Set TestTimeShouldExceedExpectedMs = p_outcome
    
End Function



