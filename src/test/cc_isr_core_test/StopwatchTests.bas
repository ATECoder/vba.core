Attribute VB_Name = "StopWatchTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Stopwatch extension methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

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



