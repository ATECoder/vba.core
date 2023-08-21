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
''' <remarks>   Dependencies: Stopwatch.cls, UserDefinedError.cls. <remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   A public creatable singleton class for creating instances of this workbook
''' classes to share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Instantiates a new <see cref="Stopwatch"/> object. </summary>
''' <returns>   [<see cref="Stopwatch"/>]. </returns>
Public Function NewStopwatch() As cc_isr_Core_IO.Stopwatch
    Set NewStopwatch = New cc_isr_Core_IO.Stopwatch
End Function

''' <summary>   Instantiates a new <see cref="UserDefinedError"/> object. </summary>
''' <returns>   [<see cref="UserDefinedError"/>]. </returns>
Public Function NewUserDefinedError() As UserDefinedError
    Set NewUserDefinedError = New UserDefinedError
End Function



