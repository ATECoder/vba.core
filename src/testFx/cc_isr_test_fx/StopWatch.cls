VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' StopWatch.cls
'
' Dependencies:
'
' None
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Retrieves the current value of the performance counter, which is a high resolution (<1us) time
''' stamp that can be used for time-interval measurements. </summary>
''' <remarks>   </remarks>
''' <param name="a_counts">   [Currency] A pointer to a variable that receives the current performance-counter
'''                           value, in counts. </param>
''' <returns>   [Byte] If the function succeeds, the return value is nonzero.
''' If the function fails, the return value is zero. To get extended error information, call
''' GetLastError. On systems that run Windows XP or later, the function will always succeed and will thus never return zero.
''' <returns>
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef a_counts As Currency) As Byte

''' <summary>   Retrieves the frequency of the performance counter. The frequency of the performance counter
'' is fixed at system boot and is consistent across all processors. Therefore, the frequency need only be queried
''' upon application initialization, and the result can be cached.</summary>
''' <remarks>   </remarks>
''' <param name="a_frequency">   [Currency] A pointer to a variable that receives the current performance-counter
'''                              frequency, in counts per second. If the installed hardware doesn't support
'''                              a high-resolution performance counter, this parameter can be zero (this will not
'''                              occur on systems that run Windows XP or later). </param>
''' <returns>   [Byte] If the installed hardware supports a high-resolution performance counter, the return
''' value is nonzero. If the function fails, the return value is zero. To get extended error information,
''' call GetLastError. On systems that run Windows XP or later, the function will always succeed and will
''' thus never return zero.
''' <returns>
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef a_frequency As Currency) As Byte
   
''' <summary>   Suspends the thread for the specified duration in milliseconds. </summary>
''' <param name="a_sleepTimeMs">   Sleep time in milliseconds. </param>
Private Declare PtrSafe Sub SleepMilliseconds Lib "kernel32" Alias "Sleep" (ByVal a_sleepTimeMs As Long)
   
Private m_overheadCounts As Currency
Private m_counting As Boolean
Private m_startCounts As Currency
Private m_endCounts As Currency
Private m_frequency As Currency

''' <summary>   Initializes the stop watch </summary>
''' <remarks>   Gets the stopwatch clock frequency. </remarks>
Private Sub Class_Initialize()
    QueryPerformanceFrequency m_frequency
    Restart
End Sub

''' <summary>   Restarts the stopwatch counter. </summary>
''' <remarks>   Sets the stopwatch start and end counts and
''' marks the stopwatch as counting.
''' </remarks>
Public Sub Restart()
    QueryPerformanceCounter m_startCounts
    QueryPerformanceCounter m_endCounts
    m_overheadCounts = m_endCounts - m_startCounts
    m_counting = True
End Sub

''' <summary>   Stops the stopwatch counter. </summary>
''' <remarks>   Fixes the stop watch end time and marks it as not counting.
''' </remarks>
Public Sub StopCounter()
    QueryPerformanceCounter m_endCounts
    m_counting = False
End Sub

''' <summary>   Returns the elapsed counts. </summary>
Property Get ElapsedCounts() As Double
    If m_counting Then
        QueryPerformanceCounter m_endCounts
    End If
    ElapsedCounts = m_endCounts - m_startCounts - m_overheadCounts
End Property

''' <summary>   Returns the elapsed time in seconds. </summary>
''' <value>   [Double] Elapsed time in seconds. </value>
Property Get ElapsedSeconds() As Double
    
    ElapsedSeconds = Me.ElapsedCounts / m_frequency

End Property

''' <summary>   Returns the elapsed time in milliseconds. </summary>
''' <value>   [Double] Elapsed time in milliseconds. </value>
Property Get ElapsedMilliseconds() As Double
    
    ElapsedMilliseconds = 1000# * Me.ElapsedSeconds

End Property

''' <summary>   Suspends the thread for the specified duration in milliseconds. </summary>
''' <param name="a_duration">   Sleep time in milliseconds. </param>
Public Sub Sleep(ByVal a_duration As Long)
    
    SleepMilliseconds a_duration

End Sub

''' <summary>   Loop around <see cref="VBA.DoEvents"> for the specified duration. </summary>
''' <param name="a_duration">   [Long] Wait time in milliseconds. </param>
Public Function Wait(ByVal a_duration As Long) As Long
    
    Me.Restart
    While Me.ElapsedMilliseconds < a_duration
        VBA.DoEvents
    Wend
    Wait = Me.ElapsedMilliseconds

End Function


