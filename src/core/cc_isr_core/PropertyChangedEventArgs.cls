VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyChangedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type this_
    PropertyName As String
End Type

Private This As this_

''' <summary>   Initializes and returns the event arguments. </summary>
''' <para name="a_propertyName">   [String] The property name. </param>
''' <value>   [<see cref="PropertyChangedEventArgs"/>]. </returns>
Public Function Initialize(ByVal a_propertyName As String) As PropertyChangedEventArgs
    This.PropertyName = a_propertyName
    Set Initialize = Me
End Function

''' <summary>   Gets the property name. </summary>
''' <value>   [String]. </value>
Public Property Get PropertyName() As String
    PropertyName = This.PropertyName
End Property

''' <summary>   Sets the property name. </summary>
''' <param name="a_value">   [String] the property name. </value>
Public Property Let PropertyName(ByVal a_value As String)
    This.PropertyName = a_value
End Property
