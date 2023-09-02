VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserDefinedError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   User-defined error. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Number As Long
    Name As String
    Description As String
    Source As String
    Timestamp As Double
End Type

Private This As this_

''' <summary>   Initializes and return this user defined error. </summary>
''' <param name="a_errorNumber">   [Long] The error Number. </param>
''' <param name="a_errorName">     [String] The error name. </param>
''' <param name="a_description">   [String] The error description. </param>
''' <returns>   [<see cref="UserDefinedError"/>]. </returns>
Public Function Initialize(ByVal a_errorNumber As Long, ByVal a_errorName As String, _
        ByVal a_description As String) As UserDefinedError
    This.Name = a_errorName
    This.Number = a_errorNumber
    This.Description = a_description
    This.Source = VBA.vbNullString
    This.Timestamp = CoreExtensions.DaysNow()
    Set Initialize = Me
End Function

''' <summary>   Clones a user defined error. </summary>
''' <param name="a_error">     [[<see cref="UserDefinedError"/>] The error to clone. </param>
''' <returns>   [<see cref="UserDefinedError"/>]. </returns>
Public Function Clone(ByVal a_error As UserDefinedError) As UserDefinedError
    Dim p_error As UserDefinedError
    Set p_error = New UserDefinedError
    p_error.Initialize a_error.Number, a_error.Name, a_error.Description
    p_error.Source = a_error.Source
    p_error.Timestamp = a_error.Timestamp
    Set Clone = p_error
End Function

''' <summary>   Gets the error Number. </summary>
''' <value>   [Long] error Number. </value>
Public Property Get Number() As Long
    Number = This.Number
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
    UserErrorNumber = Me.Number - (VBA.vbObjectError + Me.UserErrorNumberBase)
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

''' <summary>   Gets the error source. </summary>
''' <value>   The error source of String type. </value>
Public Property Get Source() As String
    Source = This.Source
End Property

''' <summary>   Sets the error source. </summary>
''' <param name="a_value">   [String] The error source. </param>
Public Property Let Source(ByVal a_value As String)
    This.Source = a_value
End Property

''' <summary>   Gets the error Timestamp in days since December 30, 1899. </summary>
''' <value>   [Double]. </value>
Public Property Get Timestamp() As Double
    Timestamp = This.Timestamp
End Property

''' <summary>   Sets the error Timestamp in days since December 30, 1899. </summary>
''' <param name="a_value">   [Double] The error Timestamp. </param>
Public Property Let Timestamp(ByVal a_value As Double)
    This.Timestamp = a_value
End Property


''' <summary>   Returns a string representation of the error. </summary>
''' <param name="additionalInfo">   [Optional, empty] additional information to append to the error message. </param>
''' <returns>   [String] A representation of the error. </returns>
Public Function ToString(Optional ByVal a_additionalInfo As String = vbNullString) As String
    
    ToString = Me.Name & " " & _
               VBA.Str$(Me.Number) & " (0x" & VBA.LCase$(VBA.Hex$(Me.Number)) & _
               " 0x" & VBA.LCase$(VBA.Hex$(Me.UserErrorNumber)) & ")" & _
               Me.Description & a_additionalInfo & _
               " occured in " & _
               IIf(VBA.vbNullString = Me.Source, " unknown", Me.Source) & _
               " at " & CoreExtensions.FormatDays(Me.Timestamp) & "."
End Function

