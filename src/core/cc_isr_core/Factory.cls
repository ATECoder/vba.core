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
''' <remarks>   Dependencies: EventTimer.cls, Stopwatch.cls, StringBuilder.cls. <remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   A public creatable singleton class for creating instances of this workbook
''' classes to share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Instantiates a new <see cref="CancelEventArgs"/> object. </summary>
''' <returns>   [<see cref="CancelEventArgs"/>]. </returns>
Public Function NewCancelEventArgs() As CancelEventArgs
    Set NewCancelEventArgs = New CancelEventArgs
End Function

''' <summary>   Instantiates an instance of an <see cref="EventTimer"/> object. </summary>
''' <returns>   [<see cref="EventTimer"/>]. </returns>
Public Function NewEventTimer() As EventTimer
  Set NewEventTimer = New EventTimer
End Function

''' <summary>   Instantiates a new <see cref="Stopwatch"/> object. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   [<see cref="Stopwatch"/>]. </returns>
Public Function NewStopwatch() As StopWatch
    Set NewStopwatch = New StopWatch
End Function

''' <summary>   Instantiates a new <see cref="StringBuilder"/> object. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   [<see cref="StringBuilder"/>]. </returns>
Public Function NewStringBuilder() As StringBuilder
    Set NewStringBuilder = New StringBuilder
End Function

''' <summary>   Instantiates a new <see cref="UserDefinedError"/> object. </summary>
''' <returns>   [<see cref="UserDefinedError"/>]. </returns>
Public Function NewUserDefinedError() As UserDefinedError
    Set NewUserDefinedError = New UserDefinedError
End Function



