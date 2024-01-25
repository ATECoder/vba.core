VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A string builder. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_initialSize As Long = 16

Private Type this_
    Buffer() As String
    CurrentIndex As Long
    MaxIndex As Long
End Type

Private This As this_

''' < summary>   Initializes this class. </summary>
''' <remarks>
''' <see href="https://codereview.stackexchange.com/questions/67596/a-lightning-fast-stringbuilder"/>
''' </remarks>
Private Sub Class_Initialize()

    This.CurrentIndex = 0
    This.MaxIndex = m_initialSize
    ReDim This.Buffer(1 To This.MaxIndex)

End Sub

''' <summary>   Initializes and return this user defined error. </summary>
''' <param name="a_initialCapacity">     [Long] The error code. </param>
''' <param name="a_errorName">     [String] The error name. </param>
''' <param name="a_eescription">   [String] The error description. </param>
''' <returns>   [<see cref="StringBuilder"/>]. </returns>
Public Function Initialize(Optional ByVal a_initialCapacity As Long = m_initialSize) As StringBuilder
    This.CurrentIndex = 0
    This.MaxIndex = a_initialCapacity
    ReDim This.Buffer(1 To This.MaxIndex)
    Set Initialize = Me
End Function


''' < summary>   Appends the given text to this StringBuilder. </summary>
''' <param name="a_text">   [String] the text to append. </param>
Public Sub Append(ByVal a_text As String)

    This.CurrentIndex = This.CurrentIndex + 1

    If This.CurrentIndex > This.MaxIndex Then
        This.MaxIndex = 2 * This.MaxIndex
        ReDim Preserve This.Buffer(1 To This.MaxIndex)
    End If
    This.Buffer(This.CurrentIndex) = a_text

End Sub

''' < summary>   Appends a formatted string to this StringBuilder. </summary>
''' <param name="a_format">   [String] the format string. </param>
''' <param name="a_values">   [Variant, ParamArray] The parameter values matching the
'''                           <ParamRef name="a_format"/>. </param>
Public Sub AppendFormat(ByVal a_format As String, ParamArray a_values() As Variant)

    If IsArray(a_values(LBound(a_values))) Then
        Me.Append StringExtensions.StringFormat(a_format, a_values(LBound(a_values)))
    Else
        Me.Append StringExtensions.StringFormat(a_format, a_values)
    End If

End Sub

''' <summary>   Returns the current length. </summary>
''' <value>   [Long]. </value>
Public Property Get length() As Long
    length = This.CurrentIndex
End Property

''' <summary>   Returns the text in this StringBuilder. </summary>
''' <param name="a_separator">   [Optional, String, Null string] The separator to use in
'''                              joining components. </param>
Public Function ToString(Optional ByVal a_separator As String = vbNullString) As String

    If This.CurrentIndex > 0 Then
        ReDim Preserve This.Buffer(1 To This.CurrentIndex)
        This.MaxIndex = This.CurrentIndex
        ToString = Join(This.Buffer, a_separator)
    End If

End Function

