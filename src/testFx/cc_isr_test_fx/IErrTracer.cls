VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrTracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents an object that can trace an error. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Traces an error message.  </summary>
''' <para name="a_message">   [String] the message to trace. </param>
Public Sub TraceError(ByVal a_message As String): End Sub

''' <summary>   Asserts and traces leftover errors from the error queue. </summary>
''' <para name="a_delimiter">   [String] a delimiter for concatenating error messages
'''                             from the error queue. </param>
Public Function AssertLeftoverErrors(Optional ByVal a_delimiter As String = VBA.vbCrLf) As Assert
End Function
