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

''' <summary>   Clones this user defined error and return a new user defined error. </summary>
''' <returns>   [<see cref="UserDefinedError"/>]. </returns>
Public Function Clone() As UserDefinedError
    Dim p_error As UserDefinedError
    Set p_error = New UserDefinedError
    p_error.Initialize Me.Number, Me.Name, Me.Description
    p_error.Source = Me.Source
    p_error.Timestamp = Me.Timestamp
    Set Clone = p_error
End Function

''' <summary>   Initialize this error from the given error. </summary>
''' <param name="a_error">     [[<see cref="UserDefinedError"/>] The error to initialize from. </param>
''' <returns>   [<see cref="UserDefinedError"/>]. </returns>
Public Function FromUserDefinedError(ByVal a_error As UserDefinedError) As UserDefinedError
    This.Name = a_error.Name
    This.Number = a_error.Number
    This.Description = a_error.Description
    This.Source = a_error.Source
    This.Timestamp = a_error.Timestamp
    Set FromUserDefinedError = Me
End Function

''' <summary>   Initialize a user defined error from a user defined error appended with
''' source, description and timestamp. </summary>
''' <param name="a_error">       [<see cref="UserDefinedError"/>]. </param>
''' <param name="a_source">      [String] The error source. </param>
''' <param name="a_suffix">      [String, Optional, Null] A suffix to append to the message description. </param>
''' <param name="a_timestamp">   [Variant double] The error timestamp. </param>
''' <returns>    [<see cref="UserDefinedError"/>]. </returns>
Public Function FromAppendedUserDefinedError(ByVal a_error As UserDefinedError, _
        Optional ByVal a_source As String, _
        Optional ByVal a_suffix As String, _
        Optional ByVal a_timestamp As Variant) As UserDefinedError

    This.Name = a_error.Name
    This.Number = a_error.Number
    This.Description = a_error.Description & a_suffix
     
    If Not VBA.IsMissing(a_source) Then _
        If VBA.vbNullString <> a_source Then _
            This.Source = a_source
    
    If Not VBA.IsMissing(a_timestamp) Then _
        This.Timestamp = a_timestamp
    
    If 0# = CDbl(This.Timestamp) Then _
        This.Timestamp = CoreExtensions.DaysNow()
    
    Set FromAppendedUserDefinedError = Me

End Function

''' <summary>   Initialize the used defined error from the VBA Error object. </summary>
''' <returns>   [<see cref="UserDefinedError"/>]. </returns>
Public Function FromErrorObject() As UserDefinedError
    This.Name = "Error-" & VBA.CStr(Err.Number)
    This.Number = VBA.Err.Number
    This.Description = VBA.Err.Description
    This.Source = Err.Source
    This.Timestamp = CoreExtensions.DaysNow()
    Set FromErrorObject = Me
End Function

''' <summary>   Initialize a user defined error from a user defined error appended with
''' source, description and timestamp. </summary>
''' <param name="a_error">       [<see cref="UserDefinedError"/>]. </param>
''' <param name="a_source">      [String] The error source. </param>
''' <param name="a_suffix">      [String, Optional, Null] A suffix to append to the message description. </param>
''' <param name="a_timestamp">   [Variant double] The error timestamp. </param>
''' <returns>    [<see cref="UserDefinedError"/>]. </returns>
Public Function FromAppendedErrorObject( _
        Optional ByVal a_source As String, _
        Optional ByVal a_suffix As String, _
        Optional ByVal a_timestamp As Variant) As UserDefinedError

    Me.FromErrorObject
    This.Description = Err.Description & a_suffix
    
    If Not VBA.IsMissing(a_source) Then _
        If VBA.vbNullString <> a_source Then _
            This.Source = a_source
    
    If Not VBA.IsMissing(a_timestamp) Then _
        This.Timestamp = a_timestamp
    
    If 0# = CDbl(This.Timestamp) Then _
        This.Timestamp = CoreExtensions.DaysNow()
    
    Set FromAppendedErrorObject = Me

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

''' <summary>   Gets the iondication that this error came from a <see cref="VBA.Err"/> object. </summary>
''' <value>   [Boolean]. </value>
Public Property Get IsFromErrorObject() As Boolean
    IsFromErrorObject = 0 <= Me.Number And 512 >= Me.Number
End Property

''' <summary>   Gets the user defined error number after removing the vbObjectError
''' and error number base. </summary>
''' <value>   [Long]. </value>
Public Property Get UserErrorNumber() As Long
    If Me.IsFromErrorObject Then
        ' if this error was traced from a VBA error object, leave the error
        ' unchanged.
        UserErrorNumber = Me.Number
    Else
        UserErrorNumber = Me.Number - (VBA.vbObjectError + Me.UserErrorNumberBase)
    End If
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
               " 0x" & VBA.LCase$(VBA.Hex$(Me.UserErrorNumber)) & ") " & _
               Me.Description & a_additionalInfo & _
               " In " & _
               IIf(VBA.vbNullString = Me.Source, " unknown", Me.Source) & _
               " at " & CoreExtensions.FormatDays(Me.Timestamp) & "."
End Function

''' <summary>   Checks if the two user defined errors are equal. </summary>
''' <param name="a_error">           [<See cref="UserDefinedError"/>] The error to check for equality. </param>
''' <param name="a_compareMethod">   [Optional, <See cref="VBA.VbCompareMethod"/>, VBA.VbCompareMethod.vbTextCompare ]
'''                                  the comparison method. </param>
''' <param name="a_timespan">        [Optional, Double, 0.01] The timespan in seconds between the two
'''                                  timestamps. </param>
''' <returns>   [Boolean]. </returns>
Public Function Equals(ByVal a_error As UserDefinedError, _
    Optional ByVal a_timespan As Double = 0.01, _
    Optional ByVal a_compareMethod As VBA.VbCompareMethod = VBA.VbCompareMethod.vbTextCompare) As Boolean
    If IsObject(a_error) Then
        Equals = True And Me.Number = a_error.Number And _
                 (0 = VBA.StrComp(Me.Name, a_error.Name, a_compareMethod)) And _
                 (0 = VBA.StrComp(Me.Description, a_error.Description, a_compareMethod)) And _
                 (0 = VBA.StrComp(Me.Source, a_error.Source, a_compareMethod)) And _
                 ((a_timespan / cc_isr_Core_IO.CoreExtensions.SecondsPerDay) >= _
                        VBA.Abs(Me.Timestamp - a_error.Timestamp))
    Else
        Equals = False
    End If
End Function

''' <summary>   Checks if the two user defined errors are the same. </summary>
''' <remarks>   This method is aimed at comparing with a user defiend error that might be
''' derived from the VBA error object and, as such, does not have the same name and might have a
''' new source but still could be aised in proximity to the user defined error. </remarks>
''' <param name="a_error">           [<See cref="UserDefinedError"/>] The error to check for equality. </param>
''' <param name="a_compareMethod">   [Optional, <See cref="VBA.VbCompareMethod"/>, VBA.VbCompareMethod.vbTextCompare ]
'''                                  the comparison method. </param>
''' <param name="a_timespan">        [Optional, Double, 0.01] The timespan in seconds between the two
'''                                  timestamps. </param>
''' <returns>   [Boolean]. </returns>
Public Function Same(ByVal a_error As UserDefinedError, _
    Optional ByVal a_timespan As Double = 0.01, _
    Optional ByVal a_compareMethod As VBA.VbCompareMethod = VBA.VbCompareMethod.vbTextCompare) As Boolean
    If IsObject(a_error) Then
        Same = True And Me.Number = a_error.Number And _
               (0 = VBA.StrComp(Me.Description, a_error.Description, a_compareMethod)) And _
               ((a_timespan / cc_isr_Core_IO.CoreExtensions.SecondsPerDay) >= _
                   VBA.Abs(Me.Timestamp - a_error.Timestamp))
    Else
        Same = False
    End If
End Function




