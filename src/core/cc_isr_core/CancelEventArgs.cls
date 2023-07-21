VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CancelEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' CancelEventArgs.cls
'
' Dependencies:
'
' None
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Private Type MyType
    CancelRequested As Boolean
End Type

Private This As MyType

''' <summary>   Initializes the event arguments. </summary>
Public Sub Initialize()
    This.CancelRequested = False
End Sub

''' <summary>   Gets the cancel requested value. </summary>
''' <value>   [Boolean] True to request the cancellation of the event. </value>
Public Property Get CancelRequested() As Boolean
    CancelRequested = This.CancelRequested
End Property

''' <summary>   Gets the cancel requested value. </summary>
''' <param name="a_value">   [Boolean] true to request the cancellation of the event. </param>
Public Property Let CancelRequested(ByVal a_value As Boolean)
    This.CancelRequested = a_value
End Property

