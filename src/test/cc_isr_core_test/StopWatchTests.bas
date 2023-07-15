Attribute VB_Name = "StopWatchTests"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' StopWatchTests.bas
'
' Dependencies:
'
' Assert.cls
' StopWatch.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Asserts that the <see cref="StopWatch"/>.<see cref="StopWatch.ElapedMilliseconds"/>
''' exceeds the thread sleep time. </summary>
''' <returns>   An instance of the <see cref="Assert"/>   class. </returns>
Public Function TestElapsedTimeShouldExceedexpectedMs() As Assert
    
    Dim p_stopper As StopWatch: Set p_stopper = cc_isr_Core.Constructor.CreateStopWatch
    Dim p_expectedMs As Long
    p_expectedMs = 100
    p_stopper.Sleep p_expectedMs + 50
    p_stopper.StopCounter
    Dim p_actualMs As Long: p_actualMs = p_stopper.ElapsedMilliseconds
    Set TestElapsedTimeShouldExceedexpectedMs = Assert.IsTrue(p_stopper.ElapsedMilliseconds > p_expectedMs, _
        "elapsed time " & CStr(p_stopper.ElapsedMilliseconds) & " must exceed sleep time " & CStr(p_expectedMs))
        
End Function

''' <summary>   Unit test. Asserts that the <see cref="StopWatch"/>.<see cref="StopWatch.Wait"/>
''' exceeds the specified interval. </summary>
''' <returns>   An instance of the <see cref="Assert"/>   class. </returns>
Public Function TestTimeShouldExceedexpectedMs() As Assert
    
    Dim p_expectedMs As Long: p_expectedMs = 100
    Dim p_stopper As StopWatch: Set p_stopper = cc_isr_Core.Constructor.CreateStopWatch
    Dim p_actualMs As Long: p_actualMs = p_stopper.Wait(p_expectedMs)
    Set TestTimeShouldExceedexpectedMs = Assert.IsTrue(p_actualMs >= p_expectedMs, _
        "elapsed time " & CStr(p_actualMs) & " must exceed " & CStr(p_expectedMs))

End Function



