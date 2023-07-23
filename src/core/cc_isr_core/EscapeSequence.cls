VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EscapeSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' EscapeSequence.cls
'
' Description: Defines an escape sequence.

' Dependencies:
'
' None
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Defines an escape sequence of a value and a replacement value. </summary>
Private Type ThisType
    Value As String
    Replacement As String
End Type

Private This As ThisType

''' <summary>   Gets the escape sequence value. </summary>
''' <value>   [String]. </value>
Public Property Get Value() As String
    Value = This.Value
End Property

''' <summary>   Sets the escape sequence value. </summary>
''' <param name="a_value">   [String] The value., e.g., '/n'. </param>
Public Property Let Value(ByVal a_value As String)
    This.Value = a_value
End Property

''' <summary>   Gets the escape sequence replacement value. </summary>
''' <value>   [String]. </value>
Public Property Get Replacement() As String
    Replacement = This.Replacement
End Property

''' <summary>   Sets the escape sequence replacement value. </summary>
''' <param name="a_value">   [String] The replacement value., e.g., <see cref="VBA.VbNewLine"/>. </param>
Public Property Let Replacement(ByVal a_value As String)
    This.Replacement = a_value
End Property



