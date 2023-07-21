VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserDefinedError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' UserDefinedError.cls
'
' Dependencies:
'
' None
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type MyType
    Code As Long
    Name As String
    Description As String
End Type

Private This As MyType


''' <summary>   Initializes a user defined error. </summary>
''' <param name="a_errorCode">     [Long] The error code. </param>
''' <param name="a_errorName">     [String] The error name. </param>
''' <param name="a_eescription">   [String] The error description. </param>
Public Sub InitiateProperties(ByVal a_errorCode As Long, ByVal a_errorName As String, ByVal a_description As String)
    This.Name = a_errorName
    This.Code = a_errorCode
    This.Description = a_description
End Sub

''' <summary>   Gets the error code. </summary>
''' <value>   [Long] error code. </value>
Public Property Get Code() As Long
    Code = This.Code
End Property

''' <summary>   Gets the error number base. Defaults to 512. </summary>
''' <value>   The error number base value of Long type. </value>
Public Property Get UserErrorNumberBase() As Long
    UserErrorNumberBase = 512
End Property

''' <summary>   Gets the user defined error number after removing the vbObjectError
''' and error number base. </summary>
''' <value>   The user error number of Long type. </value>
Public Property Get UserErrorNumber() As Long
    UserErrorNumber = Me.Code - (VBA.vbObjectError + Me.UserErrorNumberBase)
End Property

''' <summary>   Gets the error name. </summary>
''' <value>   The error name of String type. </value>
Public Property Get Name() As String
    Name = This.Name
End Property

''' <summary>   The error description. </summary>
''' <value>   The error description of String type. </value>
Public Property Get Description() As String
    Description = This.Description
End Property

''' <summary>   Returns a string representation of the error. </summary>
''' <param name="additionaInfo">   [Optional, empty] additional information to append to the error message. </param>
''' <returns>   A String representation of the error. </returns>
Public Function ToString(Optional ByVal a_additionalInfo As String = vbNullString) As String
    ToString = "User Error " & Me.Name & " (0x" & VBA.LCase$(VBA.Hex$(Me.Code)) & " 0x" & VBA.LCase$(VBA.Hex$(Me.UserErrorNumber)) & "): " & _
                Me.Description & a_additionalInfo
End Function

