VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Constructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' Constructor.cls
'
' Dependencies:
'
' EventTimer.cls
' StopWatch.cls
' StringBuilder.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   A public creatable singleton class for creating instances of this workbook
''' classes to share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Creates an instance of the <see cref="EventTimer"/> class. </summary>
''' <param name="a_intervalMs">   [Long] Number of milliseconds to raise Tick Event between </param>
''' <param name="a_callback">     [IRunnable=Nothing] Callback object. Must contain a Run() method. </param>
''' <returns> [see cref=""EventTimer"/>. The frequency isn't guaranteed but it will be as close as humanly possible. </returns>
''' <returns>   A new instance of the <see cref="EventTimer"/> class. </returns>
Public Function CreateEventTimer(ByVal a_intervalMs As Long, Optional ByVal a_callback As IRunnable = Nothing) As EventTimer
  Set CreateEventTimer = New EventTimer
  Call CreateEventTimer.ProtInit(a_intervalMs, a_callback)
End Function

''' <summary>   Creates a new instance of the <see cref="StopWatch"/> class to share
''' with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="StopWatch"/> class. </returns>
Public Function CreateStopWatch() As StopWatch
    Set CreateStopWatch = New StopWatch
End Function

''' <summary>   Creates a new instance of the <see cref="StringBuilder"/> class to share
''' with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="StringBuilder"/> class. </returns>
Public Function CreateStringBuilder() As StringBuilder
    Set CreateStringBuilder = New StringBuilder
End Function

''' <summary>   Creates a new instance of the <see cref="CancelEventArgs"/> class to share
''' with external workbooks. </summary>
''' <returns>   A new instance of the <see cref="CancelEventArgs"/> class. </returns>
Public Function CreateCancelEventArgs() As CancelEventArgs
    Set CreateCancelEventArgs = New CancelEventArgs
End Function



