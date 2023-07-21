VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' StringBuilder.cls
'
' Dependencies:
'
' None
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const initialSize As Integer = 16
Private m_buffer() As String
Private m_currentIndex As Long
Private m_maxIndex As Long

''' < summary>   Initializes this class. </summary>
''' <remarks>
''' <see href="https://codereview.stackexchange.com/questions/67596/a-lightning-fast-stringbuilder"/>
''' </remarks>
Private Sub Class_Initialize()

    m_currentIndex = 0
    m_maxIndex = initialSize
    ReDim m_buffer(1 To m_maxIndex)

End Sub

''' < summary>   Appends the given text to this StringBuilder. </summary>
''' <param name="a_text">   [String] the text to append. </param>
Public Sub Append(ByVal a_text As String)

    m_currentIndex = m_currentIndex + 1

    If m_currentIndex > m_maxIndex Then
        m_maxIndex = 2 * m_maxIndex
        ReDim Preserve m_buffer(1 To m_maxIndex)
    End If
    m_buffer(m_currentIndex) = a_text

End Sub

''' < summary>   Appends a formatted string to this StringBuilder. </summary>
''' <param name="a_format">   [String] the format string. </param>
''' <param name="a_tokens">   [Variant, ParamArray] The parameter list matching the mask. </param>
Public Sub AppendFormat(ByVal a_format As String, ParamArray a_tokens() As Variant)

    If IsArray(a_tokens(LBound(a_tokens))) Then
        Me.Append StringExtensions.StringFormat(a_format, a_tokens(LBound(a_tokens)))
    Else
        Me.Append StringExtensions.StringFormat(a_format, a_tokens)
    End If

End Sub

''' <summary>   Returns the current length. </summary>
''' <value>   [Long]. </value>
Public Property Get Length() As Long
    Length = m_currentIndex
End Property

''' <summary>   Returns the text in this StringBuilder. </summary>
''' <param name="a_separator">   [Optional, String, Null string] The separator to use in
'''                              joining components. </param>
Public Function ToString(Optional ByVal a_separator As String = vbNullString) As String

    If m_currentIndex > 0 Then
        ReDim Preserve m_buffer(1 To m_currentIndex)
        m_maxIndex = m_currentIndex
        ToString = Join(m_buffer, a_separator)
    End If

End Function

