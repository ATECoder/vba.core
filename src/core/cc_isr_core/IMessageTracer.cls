VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMessageTracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents an object that can trace messages. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Values that represent and actual and maximum trace levels. </summary>
Public Enum TraceLevelOption
    ''' <summary>   An enum constant representing the off option. </summary>
    Off = 0
    ''' <summary>   An enum constant representing the Error option. </summary>
    Error = 1
    ''' <summary>   An enum constant representing the Warning option. </summary>
    Warning = 2
    ''' <summary>   An enum constant representing the Info option. </summary>
    Info = 3
    ''' <summary>   An enum constant representing the verbose option. </summary>
    Verbose = 4
End Enum

''' <summary>   Traces the specified message at the specified level.  </summary>
''' <param name="a_message">   [String] The message to trace. </param>
''' <param name="a_level">    [<see cref="TraceLevelOption"/>, Info] The level at which the message is traced. </param>
Public Sub TraceMessage(ByVal a_message As String, Optional ByVal a_level As TraceLevelOption = 3): End Sub

''' <summary>   Gets or sets the maximum trace level. Messages are traced at or below this level.  </summary>
Public Property Get TraceLevel() As TraceLevelOption: End Property

Public Property Let TraceLevel(ByVal a_value As TraceLevelOption): End Property

''' <summary>   Traces the current <see cref="Err"/> object.  </summary>
Public Sub TraceError(): End Sub


