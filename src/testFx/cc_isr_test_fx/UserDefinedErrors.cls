VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserDefinedErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' UserDefinedErrors.cls
'
' Dependencies:
'
' StringBuilder.cls
' WorkbookUtilitites.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Builds a an error message from the VBA <see cref="Err"/>. </summary>
Public Function FormattedErr() As String
    
    Dim p_errNoDetails As String
    p_errNoDetails = " (" & (Err.Number - VBA.vbObjectError) & " / 0x" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")"
    FormattedErr = "Error Number=" & Err.Number & p_errNoDetails & ", Source=" & _
                   Err.Source & ", Description=" & Err.Description

End Function

''' <summary>   Sets the Err object source string to project.module.procedure. </summary>
''' <param name="a_errProcedureName">   [String] Specifies the name of the procedure. </param>
''' <param name="a_errModuleName">      [String] Specifies the module name. </param>
''' <param name="a_errProjectName">     [Optional, String] Specifies the project name; otherwise the project
'''                                     name of the active workbook is used. </param>
Public Sub SetErrSource(ByVal a_errProcedureName As String, ByVal a_errModuleName As String, _
        Optional ByVal a_errProjectName As String = vbNullString)

    ' this procedure must not trap errors because it must
    ' not alter the error object.
      
    ' thus we assume that this code is robust and will
    ' not cause errors.
    
    If vbNullString = a_errProjectName Then
        a_errProjectName = ActiveWorkbook.VBProject.Name
    End If
    
    ' get the current source string
    
    Dim p_errorSource As String: p_errorSource = Err.Source
    
    ' build the error source.
    
    p_errorSource = a_errProjectName & "." & a_errModuleName & "." & a_errProcedureName
  
    ' Update the Err.Source
    
    Err.Source = p_errorSource
 
End Sub

''' <summary>   Builds a standard error message. </summary>
''' <param name="a_displayWarning">         [Optional, Boolean, false] True
'''                                         to display a warning rather than
'''                                         an error message. </param>
''' <param name="a_descriptionDelimiter">   [Optional, String, ': '] Specify
'''                                         the delimiter preceding the description. </param>
''' <returns>   A Standard error message string in the form: <para>
''' Error # (0x#) occurred in <c>Source</c>: Description </para><para>
''' or  </para><para>
''' Warning # (0x#) occurred in <c>Source</c>: Description  </para>
''' </returns>
Public Function BuildStandardErrorMessage(Optional ByVal a_displayWarning As Boolean = False, _
                Optional a_descriptionDelimiter As String = ": ") As String

    Dim p_builder As StringBuilder: Set p_builder = New StringBuilder
  
    ' check if we have an error
    If Err.Number <> 0 Then
    
        p_builder.Append IIf(a_displayWarning, "Warning ", "Error ")
        p_builder.Append Format$(Err.Number)
        
        Dim p_errNumber As Long: p_errNumber = Err.Number - vbObjectError
        
        If Abs(p_errNumber) < &HFFFF& Then
            p_errNumber = p_errNumber - &H200
            p_builder.Append " (+0x"
        Else
            p_errNumber = Err.Number
            p_builder.Append " (0x"
        End If
        p_builder.Append Hex$(p_errNumber)
        p_builder.Append ") "
        p_builder.Append "occurred in "
        p_builder.Append Err.Source
        p_builder.Append a_descriptionDelimiter
        p_builder.Append Err.Description
        
    End If

    BuildStandardErrorMessage = p_builder.ToString

End Function


