VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EscapeSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An escape sequence. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Defines an escape sequence of a value and a Literal value. </summary>
Private Type this_
    Value As String
    Literal As String
End Type

Private This As this_

''' <summary>   Initializes and returns this object. </summary>
''' <param name="a_value">     [String] The value, e.g., '/r'. </param>
''' <param name="a_literal">   [String] The Literal value corresponding to the value,
'''                            e.g., <see cref="VBA.VbCr"/>. </param>
''' <returns>   [<see cref="EscapeSequence"/>]. </returns>
Public Function Initialize(ByVal a_value As String, ByVal a_literal As String) As EscapeSequence
    This.Value = a_value
    This.Literal = a_literal
    Set Initialize = Me
End Function

''' <summary>   Gets the escape sequence value. </summary>
''' <value>   [String]. </value>
Public Property Get Value() As String
    Value = This.Value
End Property

''' <summary>   Sets the escape sequence value. </summary>
''' <param name="a_value">   [String] The value, e.g., '/r'. </param>
Public Property Let Value(ByVal a_value As String)
    This.Value = a_value
End Property

''' <summary>   Gets the escape sequence literal representation value. </summary>
''' <value>   [String]. </value>
Public Property Get Literal() As String
    Literal = This.Literal
End Property

''' <summary>   Sets the escape sequence literal representation value. </summary>
''' <param name="a_value">   [String] The Literal value corresponding to the value,
'''                          e.g., <see cref="VBA.VbCr"/>. </param>
Public Property Let Literal(ByVal a_value As String)
    This.Literal = a_value
End Property



