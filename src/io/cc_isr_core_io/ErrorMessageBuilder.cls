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
''' <summary>   Constructs workbook level error source and messages. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Builds a an error message from the VBA <see cref="Err"/>. </summary>
''' <returns>   A comma-separated error message string in the form: <para>
''' Error Number=#(#/0x# (Last DLL Error 0x#)), Source={source}, Description={description}.
''' </returns>
Public Function FormattedErr() As String
    
    Dim p_errNoDetails As String
    p_errNoDetails = "(" & _
        VBA.CStr(VBA.Err.Number - VBA.vbObjectError) & _
        " 0x" & VBA.LCase$(VBA.Hex$(VBA.Err.Number))
    If Err.LastDllError <> 0 Then
        p_errNoDetails = p_errNoDetails & " (Last DLL Error 0x"
        p_errNoDetails = p_errNoDetails & VBA.LCase$(VBA.Hex$(Err.LastDllError)) & ")"
    End If
    p_errNoDetails = p_errNoDetails & ")"
    FormattedErr = "Error Number=" & VBA.Err.Number & p_errNoDetails & ", Source=" & _
                   VBA.Err.Source & ", Description=" & VBA.Err.Description

End Function

''' <summary>   Builds the Err object source string to project.module.procedure. </summary>
''' <param name="a_errProcedureName">   [String] Specifies the name of the procedure. </param>
''' <param name="a_errModuleName">      [String] Specifies the module name. </param>
''' <param name="a_errProjectName">     [Optional, String] Specifies the project name; otherwise the project
'''                                     name of the active workbook is used. </param>
''' <returns>   A Standard error source string to project.module.procedure. </returns>
Friend Function BuildErrSource(ByVal a_errProcedureName As String, ByVal a_errModuleName As String, _
        Optional ByVal a_errProjectName As String = vbNullString) As String

    ' this procedure must not trap errors because it must
    ' not alter the error object.
      
    ' thus we assume that this code is robust and will
    ' not cause errors.
    
    If VBA.vbNullString = a_errProjectName And Not ActiveWorkbook Is Nothing Then
        If Not ActiveWorkbook.VBProject Is Nothing Then _
            a_errProjectName = ActiveWorkbook.VBProject.Name
    End If
    
    Dim p_builder As String
    p_builder = IIf(VBA.vbNullString = a_errProjectName, "unknown", a_errProjectName) _
        & "." & a_errModuleName & "." & a_errProcedureName
    
    ' if the built source does not start with the Err.Source, this must be a new
    ' source, so append it to the built source.
    If 1 <> VBA.InStr(1, p_builder, Err.Source, VBA.VbCompareMethod.vbTextCompare) Then _
        p_builder = p_builder & " in " & Err.Source
    
    BuildErrSource = p_builder
 
End Function

''' <summary>   Builds the error source to append to the <see cref="Err"/>.<see cref="Err.Source"/>
'''             as project.module.procedure. </summary>
''' <param name="a_procedureName">   [String] Specifies the name of the procedure. </param>
''' <param name="a_moduleName">      [String] Specifies the module name. </param>
''' <param name="a_workbook">        [Excel.Workbook] Specifies the workbook. </param>
''' <returns>   A Standard error source string to project.module.procedure. </returns>
Friend Function BuildErrorSource(ByVal a_procedureName As String, ByVal a_moduleName As String, _
        ByVal a_workbook As Excel.Workbook) As String

    Dim p_builder As String
    If a_workbook Is Nothing Then
        p_builder = "unknown" & "." & a_moduleName & "." & a_procedureName
    Else
        p_builder = a_workbook.VBProject.Name & "." & a_moduleName & "." & a_procedureName
    End If
    
    ' append the Err.Source if it does not equal or is not contained in the built source
    If 1 <> VBA.InStr(1, p_builder, Err.Source) Then _
        p_builder = p_builder & " in " & Err.Source
    
    BuildErrorSource = p_builder
 
End Function

''' <summary>   Sets the Err object source string to project.module.procedure. </summary>
''' <param name="a_errProcedureName">   [String] Specifies the name of the procedure. </param>
''' <param name="a_errModuleName">      [String] Specifies the module name. </param>
''' <param name="a_errProjectName">     [Optional, String] Specifies the project name; otherwise the project
'''                                     name of the active workbook is used. </param>
Public Sub SetErrSource(ByVal a_errProcedureName As String, ByVal a_errModuleName As String, _
        Optional ByVal a_errProjectName As String = vbNullString)

    Dim p_errorSource As String
    p_errorSource = Me.BuildErrSource(a_errProcedureName, a_errModuleName, a_errProjectName)

    ' Update the Err.Source for reporting the error.
    Err.Source = p_errorSource
 
End Sub

''' <summary>   Appends and apply a source to the <see cref="Err"/>.<see cref="Err.Source"/> if
'''             nont already specified. </summary>
''' <param name="a_errProcedureName">   [String] Specifies the name of the procedure. </param>
''' <param name="a_errModuleName">      [String] Specifies the module name. </param>
''' <param name="a_workbook">           [<] Specifies the project name; otherwise the project
'''                                     name of the active workbook is used. </param>
''' <returns>   [String] The appedned error source. </returns>
Public Function AppendErrSource(ByVal a_errProcedureName As String, ByVal a_errModuleName As String, _
        ByVal a_workbook As Excel.Workbook) As String

    Dim p_errorSource As String
    p_errorSource = Me.BuildErrorSource(a_errProcedureName, a_errModuleName, a_workbook)

    ' Update the Err.Source for reporting the error.
    Err.Source = p_errorSource
    
    AppendErrSource = p_errorSource
 
End Function

''' <summary>   Builds a standard error message. </summary>
''' <param name="a_displayWarning">         [Boolean, false] True
'''                                         to display a warning rather than
'''                                         an error message. </param>
''' <returns>   A Standard error message string in the form: <para>
''' Error 0x# (Last DLL Error 0x#) Description occurred in <c>Source</c> at 20230901113506.123. </para><para>
''' or  </para><para>
''' Warning 0x# (Last DLL Error 0x#) Description occurred in <c>Source</c> at 20230901113506.123. </para>
''' </returns>
Friend Function BuildErrorMessage(Optional ByVal a_displayWarning As Boolean = False) As String

    Dim p_builder As String: p_builder = VBA.vbNullString
  
    ' check if we have an error
    If VBA.Err.Number <> 0 Then
    
        p_builder = p_builder & IIf(a_displayWarning, "Warning ", "Error ")
        p_builder = p_builder & VBA.CStr(VBA.Err.Number)
        
        Dim p_errNumber As Long: p_errNumber = VBA.Err.Number
        If VBA.Err.Number < 0 Then
            p_errNumber = VBA.Err.Number - VBA.vbObjectError
        Else
            p_errNumber = VBA.Err.Number
        End If
        
        If Abs(p_errNumber) > &H200 Then
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
        p_builder = p_builder & VBA.Err.Description
        p_builder = p_builder & " In "
        p_builder = p_builder & VBA.Err.Source
        p_builder = p_builder & " at "
        p_builder = p_builder & CoreExtensions.FormatDays(CoreExtensions.DaysNow())
        p_builder = p_builder & "."
        
    End If

    BuildErrorMessage = p_builder

End Function

''' <summary>   Builds a standard error message. </summary>
''' <remarks>   This method enqueues the VBA Error object if the error is not
''' a <see cref="UserDefinedError"/>. It then concatenates all the errors from the
''' <see cref="UserDefinedErrors.ErrorQueue"/>, which are moved to the
''' <see cref="UserDefinedErrors.ErrorArchveStack"/>. <para>
''' Having an error queue also addresses automation errors where the
''' reporting workbook has no access to the error object of the error source, which
''' might reside in a workbook that is not referenced by the reporting workbook.</para>
''' </remarks>
''' <param name="a_delimiter">   [String, VBCrLf] delimiter between errors. </param>
''' <returns>   [String]. </returns>
Public Function BuildStandardErrorMessage(Optional ByVal a_delimiter As String = VBA.vbCrLf) As String

    ' enqueue the error object or append its source to the last error.
    UserDefinedErrors.EnqueueErrorObject
    
    ' return the queued errors as a delimited string.
    BuildStandardErrorMessage = cc_isr_Core_IO.UserDefinedErrors.ArchiveQueuedErrors(a_delimiter)

End Function

''' <summary>   Builds a delimited string of the archived messages clearing teh archive. </summary>
''' <param name="a_delimiter">   [String, VBCrLf] delimiter between errors. </param>
''' <returns>   [String]. </returns>
Public Function BuildArchivedErrorsMessage(Optional ByVal a_delimiter As String = VBA.vbCrLf) As String

    ' return the archived errors as a delimited string.
    BuildArchivedErrorsMessage = cc_isr_Core_IO.UserDefinedErrors.PopArchivedErrors(a_delimiter)

End Function

