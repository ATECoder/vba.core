VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Constructs public not creatable objects for use by referencing workbooks. </summary>
''' <remarks>   Dependencies: EventTimer.cls, StopWatch.cls, StringBuilder.cls. <remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   A public creatable singleton class for creating instances of this workbook
''' classes to share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Instantiates and initializes a <see cref="EventTimer"/> object. </summary>
''' <param name="a_intervalMs">   [Long] Number of milliseconds to raise Tick Event between </param>
''' <param name="a_callback">     [IRunnable=Nothing] Callback object. Must contain a Run() method. </param>
''' <returns> [see cref=""EventTimer"/>. The frequency isn't guaranteed but it will be as close as humanly possible. </returns>
''' <returns>   A new instance of the <see cref="EventTimer"/> object. </returns>
Public Function CreateEventTimer(ByVal a_intervalMs As Long, Optional ByVal a_callback As IRunnable = Nothing) As EventTimer
  Set CreateEventTimer = New EventTimer
  Call CreateEventTimer.ProtInit(a_intervalMs, a_callback)
End Function

''' <summary>   Instantiates a new <see cref="StopWatch"/> object. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="StopWatch"/> object. </returns>
Public Function NewStopWatch() As StopWatch
    Set NewStopWatch = New StopWatch
End Function

''' <summary>   Instantiates a new <see cref="StringBuilder"/> object. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="StringBuilder"/> object. </returns>
Public Function NewStringBuilder() As StringBuilder
    Set NewStringBuilder = New StringBuilder
End Function

''' <summary>   Instantiates a new <see cref="CancelEventArgs"/> object. </summary>
''' <returns>   A new instance of the <see cref="CancelEventArgs"/> object. </returns>
Public Function NewCancelEventArgs() As CancelEventArgs
    Set NewCancelEventArgs = New CancelEventArgs
End Function



