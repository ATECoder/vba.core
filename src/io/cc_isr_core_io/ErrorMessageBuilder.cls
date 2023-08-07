VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorMessageBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Builds error messages. </summary>
''' <remarks> Dependencies: StringBuilder.cls.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Builds a an error message from the VBA <see cref="Err"/>. </summary>
''' <returns>   A comma-separated error message string in the form: <para>
''' Error Number=#(#/0x# (Last DLL Error 0x#)), Source={source}, Description={description}.
''' </returns>
Public Function FormattedErr() As String
    
    Dim p_errNoDetails As String
    p_errNoDetails = "(" & (VBA.Err.Number - VBA.vbObjectError) & "/0x" & _
                     VBA.LCase$(VBA.Hex$(VBA.Err.Number))
    If Err.LastDllError <> 0 Then
        p_errNoDetails = p_errNoDetails & " (Last DLL Error 0x"
        p_errNoDetails = p_errNoDetails & VBA.LCase$(VBA.Hex$(Err.LastDllError)) & ")"
    End If
    p_errNoDetails = p_errNoDetails & ")"
    FormattedErr = "Error Number=" & VBA.Err.Number & p_errNoDetails & ", Source=" & _
                   VBA.Err.Source & ", Description=" & VBA.Err.Description

End Function

''' <summary>   Sets the Err object source string to project.module.procedure. </summary>
''' <param name="a_errProcedureName">   [String] Specifies the name of the procedure. </param>
''' <param name="a_errModuleName">      [String] Specifies the module name. </param>
''' <param name="a_errProjectName">     [String] Specifies the project name; otherwise the project
'''                                     name of the active workbook is used. </param>
Public Sub SetErrSource(ByVal a_errProcedureName As String, ByVal a_errModuleName As String, _
        Optional ByVal a_errProjectName As String = VBA.vbNullString)

    ' this procedure must not trap errors because it must
    ' not alter the error object.
      
    ' thus we assume that this code is robust and will
    ' not cause errors.
    
    If VBA.vbNullString = a_errProjectName Then
        a_errProjectName = ActiveWorkbook.VBProject.Name
    End If
    
    ' get the current source string
    
    Dim p_errorSource As String: p_errorSource = VBA.Err.Source
    
    ' build the error source.
    
    p_errorSource = a_errProjectName & "." & a_errModuleName & "." & a_errProcedureName
  
    ' Update the VBA.Err.Source
    
    VBA.Err.Source = p_errorSource
 
End Sub

''' <summary>   Builds a standard error message. </summary>
''' <param name="a_displayWarning">         [Boolean, false] True
'''                                         to display a warning rather than
'''                                         an error message. </param>
''' <param name="a_descriptionDelimiter">   [String, ': '] Specify
'''                                         the delimiter preceding the description. </param>
''' <returns>   A Standard error message string in the form: <para>
''' Error 0x# (Last DLL Error 0x#) occurred in <c>Source</c>: Description </para><para>
''' or  </para><para>
''' Warning 0x#  (Last DLL Error 0x#) occurred in <c>Source</c>: Description  </para>
''' </returns>
Public Function BuildStandardErrorMessage(Optional ByVal a_displayWarning As Boolean = False, _
                Optional a_descriptionDelimiter As String = ": ") As String

    Dim p_builder As String: p_builder = VBA.vbNullString
  
    ' check if we have an error
    If VBA.Err.Number <> 0 Then
    
        p_builder = p_builder & IIf(a_displayWarning, "Warning ", "Error ")
        p_builder = p_builder & Format$(VBA.Err.Number)
        
        Dim p_errNumber As Long: p_errNumber = VBA.Err.Number - vbObjectError
        
        If Abs(p_errNumber) < &HFFFF& Then
            p_errNumber = p_errNumber - &H200
            p_builder = p_builder & " +0x"
        Else
            p_errNumber = VBA.Err.Number
            p_builder = p_builder & " 0x"
        End If
        p_builder = p_builder & VBA.LCase$(VBA.Hex$(p_errNumber))
        If Err.LastDllError <> 0 Then
            p_builder = p_builder & " (Last DLL Error 0x"
            p_builder = p_builder & VBA.LCase$(VBA.Hex$(Err.LastDllError)) & ")"
        End If
        p_builder = p_builder & " occurred in "
        p_builder = p_builder & VBA.Err.Source
        p_builder = p_builder & a_descriptionDelimiter
        p_builder = p_builder & VBA.Err.Description
        
    End If

    BuildStandardErrorMessage = p_builder

End Function


