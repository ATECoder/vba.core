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
''' <summary>   A public creatable singleton class for creating instances of this workbook
'''             classes to share with external workbooks. </summary>
''' <remarks>   Dependencies: EventTimer.cls, StringBuilder.cls. <remarks>
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

''' <summary>   Instantiates a new <see cref="EventArgs"/> object. </summary>
''' <returns>   [<see cref="EventArgs"/>]. </returns>
Public Function NewEventArgs() As EventArgs
    Set NewEventArgs = New EventArgs
End Function

''' <summary>   Instantiates an instance of an <see cref="EventTimer"/> object. </summary>
''' <returns>   [<see cref="EventTimer"/>]. </returns>
Public Function NewEventTimer() As EventTimer
  Set NewEventTimer = New EventTimer
End Function

''' <summary>   Instantiates a new <see cref="PropertyChangedEventArgs"/> object. </summary>
''' <returns>   [<see cref="PropertyChangedEventArgs"/>]. </returns>
Public Function NewPropertyChangedEventArgs() As cc_isr_Core.PropertyChangedEventArgs
    Set NewPropertyChangedEventArgs = New cc_isr_Core.PropertyChangedEventArgs
End Function

''' <summary>   Instantiates and initializes a new <see cref="PropertyChangedEventArgs"/> object. </summary>
''' <param name="a_propertyName">   [String] the property name. </value>
''' <returns>   [<see cref="PropertyChangedEventArgs"/>]. </returns>
Public Function CreatePropertyChangedEventArgs(ByVal a_propertyName As String) As cc_isr_Core.PropertyChangedEventArgs
    Set CreatePropertyChangedEventArgs = Me.NewPropertyChangedEventArgs.Initialize(a_propertyName)
End Function

''' <summary>   Instantiates a new <see cref="StringBuilder"/> object. </summary>
''' <returns>   [<see cref="StringBuilder"/>]. </returns>
Public Function NewStringBuilder() As StringBuilder
    Set NewStringBuilder = New StringBuilder
End Function

