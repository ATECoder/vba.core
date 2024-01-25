VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CancelEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Cancel event arguments. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

Private Type this_
    CancelRequested As Boolean
End Type

Private This As this_

''' <summary>   Initializes and returns this object. </summary>
''' <param name="a_cancelRequested">   [Boolean] true to request the cancellation of the event. </param>
''' <returns>   [<see cref="CancelEventArgs"/>]. </returns>
Public Function Initialize(Optional ByVal a_cancelRequested As Boolean = False) As CancelEventArgs
    This.CancelRequested = a_cancelRequested
    Set Initialize = Me
End Function

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

