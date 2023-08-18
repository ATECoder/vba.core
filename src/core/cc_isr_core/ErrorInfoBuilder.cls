VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorInfoBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Constructs workbook level error source and messages. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Builds the Err object source string to project.module.procedure. </summary>
''' <param name="a_errProcedureName">   [String] Specifies the name of the procedure. </param>
''' <param name="a_errModuleName">      [String] Specifies the module name. </param>
''' <param name="a_errProjectName">     [Optional, String] Specifies the project name; otherwise the project
'''                                     name of the active workbook is used. </param>
''' <returns>   A Standard error source string to project.module.procedure.  </returns>
Friend Function BuildErrSource(ByVal a_errProcedureName As String, ByVal a_errModuleName As String, _
        Optional ByVal a_errProjectName As String = vbNullString) As String

    ' this procedure must not trap errors because it must
    ' not alter the error object.
      
    ' thus we assume that this code is robust and will
    ' not cause errors.
    
    If VBA.vbNullString = a_errProjectName Then a_errProjectName = ActiveWorkbook.VBProject.Name
    
    ' get the current source string
    
    Dim p_errorSource As String: p_errorSource = Err.Source
    
    ' build the error source.
    
    p_errorSource = a_errProjectName & "." & a_errModuleName & "." & a_errProcedureName
  
    ' Update the Err.Source
    
    Err.Source = p_errorSource
 
End Function

''' <summary>   Sets the Err object source string to project.module.procedure. </summary>
''' <param name="a_errProcedureName">   [String] Specifies the name of the procedure. </param>
''' <param name="a_errModuleName">      [String] Specifies the module name. </param>
''' <param name="a_errProjectName">     [Optional, String] Specifies the project name; otherwise the project
'''                                     name of the active workbook is used. </param>
Public Sub SetErrSource(ByVal a_errProcedureName As String, ByVal a_errModuleName As String, _
        Optional ByVal a_errProjectName As String = vbNullString)

    Dim p_errorSource As String

    On Error Resume Next
    
    ' Try to build the error source the user defined errors object.
    ' This is done in case we have an issue connecting to the external
    ' workbook and wish to prevent a crash of this workbook.
    p_errorSource = cc_isr_Core.UserDefinedErrors.BuildErrSource(a_errProcedureName, a_errModuleName, a_errProjectName)
    
    On Error GoTo 0
    
    If VBA.vbNullString = p_errorSource Then
    
        ' if the build failed, use the local method to build the error source.
        p_errorSource = Me.BuildErrSource(a_errProcedureName, a_errModuleName, a_errProjectName)
    
    End If
    
    ' Update the Err.Source for reporting the error.
   
    Err.Source = p_errorSource
 
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
Friend Function BuildErrorMessage(Optional ByVal a_displayWarning As Boolean = False, _
                Optional a_descriptionDelimiter As String = ": ") As String

    Dim p_builder As String: p_builder = VBA.vbNullString
  
    ' check if we have an error
    If VBA.Err.Number <> 0 Then
    
        p_builder = p_builder & IIf(a_displayWarning, "Warning ", "Error ")
        p_builder = p_builder & Format$(VBA.Err.Number)
        
        Dim p_errNumber As Long: p_errNumber = VBA.Err.Number - vbObjectError
        
        If Abs(p_errNumber) < &HFFFF& Then
            p_errNumber = p_errNumber - &H200
            p_builder = p_builder & "(+0x"
        Else
            p_errNumber = VBA.Err.Number
            p_builder = p_builder & "(0x"
        End If
        p_builder = p_builder & VBA.LCase$(VBA.Hex$(p_errNumber)) & ")"
        If Err.LastDllError <> 0 Then
            p_builder = p_builder & " Last DLL Error 0x"
            p_builder = p_builder & VBA.LCase$(VBA.Hex$(Err.LastDllError)) & ")"
        End If
        p_builder = p_builder & " occurred in "
        p_builder = p_builder & VBA.Err.Source
        p_builder = p_builder & a_descriptionDelimiter
        p_builder = p_builder & VBA.Err.Description
        
    End If

    BuildErrorMessage = p_builder

End Function

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

    On Error Resume Next
    
    ' Try to build the error message the user defined errors object.
    ' This is done in case we have an issue connecting to the external
    ' workbook and wish to prevent a crash of this workbook.
    p_builder = cc_isr_Core.UserDefinedErrors.BuildStandardErrorMessage(a_displayWarning, a_descriptionDelimiter)
    
    On Error GoTo 0
    
    If VBA.vbNullString = p_builder Then
    
        ' if the build failed, use the local method to build the error message.
        p_builder = Me.BuildErrorMessage(a_displayWarning, a_descriptionDelimiter)
    
    End If

    BuildStandardErrorMessage = p_builder

End Function

