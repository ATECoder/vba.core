VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Type ThisData
    ExportCodeAfterSave As Boolean
    Deployed As Boolean
End Type

Private This As ThisData

''' <summary>   Gets the option for exporting code files after save. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ExportCodeAfterSave() As Boolean
    ExportCodeAfterSave = This.ExportCodeAfterSave
End Property

''' <summary>   Enables or disables exporting code files after save. </summary>
''' <para name="value">   True to enable exporting code files after save. </param>
Public Property Let ExportCodeAfterSave(ByVal a_value As Boolean)
    This.ExportCodeAfterSave = a_value
End Property

''' <summary>   Gets the deployed status. </summary>
''' <remarks>   Code is not saved if the workbook is deployed. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get Deployed() As Boolean
    Deployed = This.Deployed
End Property

''' <summary>   Sets the deployed status. </summary>
''' <para name="value">   True to set the deployed status. </param>
Public Property Let Deployed(ByVal a_value As Boolean)
    This.Deployed = a_value
End Property

''' <summary>   Sets the Err object source string to project.module.procedure. </summary>
''' <param name="a_errProcedureName">   [String] Specifies the name of the procedure. </param>
''' <param name="a_errModuleName">      [String] Specifies the module name. </param>
''' <param name="a_errProjectName">     [String] Specifies the project name; otherwise the project
'''                                     name of the active workbook is used. </param>
Public Sub SetErrSource(ByVal a_errProcedureName As String, ByVal a_errModuleName As String, _
        Optional ByVal a_errProjectName As String = vbNullString)

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

''' <summary>   Marks a workbook as saved if not saved and deployed or read only. </summary>
''' <param name="a_workbook">   [Excel.Workbook] The workbook. </param>
''' <param name="a_deployed">   [Boolean] True if the workbook was deployed. </param>
Public Sub MarkAsSaved(ByVal a_workbook As Excel.Workbook, ByVal a_deployed As Boolean)
    
    If a_deployed Or a_workbook.ReadOnly Then _
        a_workbook.Saved = True

End Sub

''' <summary>   Disables saving code if deployed or read only. </summary>
''' <param name="a_workbook">   [Excel.Workbook] The workbook. </param>
''' <param name="a_deployed">   [Boolean] True if the workbook was deployed. </param>
Public Sub DisableSavingCode(ByVal a_workbook As Excel.Workbook, ByVal a_deployed As Boolean)
    
    a_workbook.ExportCodeAfterSave = Not (a_deployed Or a_workbook.ReadOnly)

End Sub

''' <summary>   Returns true if the editor window is open and we have
''' open code editor windows. </summary>
Public Function AreCodeEditorWindowsOpen() As Boolean

    Dim p_workbook As Workbook
    Dim p_vbeWindow As Variant
    
    ' If VBE is closed the user shall allow it to open first
    
    If Application.VBE.MainWindow.Visible Then
        For Each p_workbook In Application.Workbooks
            If p_workbook.Name = ThisWorkbook.Name Then
                For Each p_vbeWindow In p_workbook.VBProject.VBE.Windows
                    If p_vbeWindow.Visible = True And p_vbeWindow.Type = 0 Then
                        AreCodeEditorWindowsOpen = True
                        Exit Function
                    End If
               Next p_vbeWindow
            End If
        Next p_workbook
    End If
    
    AreCodeEditorWindowsOpen = False

End Function

''' <summary>   Closes all currently open editor windows. </summary>
Public Sub CloseAllEditorWindows()

    Dim p_workbook As Workbook
    Dim p_codeWindow As Variant
    
    ' If VBE is closed the user shall allow it to open first
    
    If Application.VBE.MainWindow.Visible = False Then
        If MsgBox("VBE is still closed!" & vbCrLf & _
            "The operation requires VBE to be open during its process. Would you like to open it?", _
            VBA.vbInformation + VBA.vbYesNo) = VBA.vbYes Then
            Application.VBE.MainWindow.Visible = True
        Else
            Exit Sub
        End If
    End If
    
    ' Then it will close all the windows
    ' Except the Default VBE Windows (Immediate Windows, Locals Windows ...)
    For Each p_workbook In Application.Workbooks
        If p_workbook.Name = ThisWorkbook.Name Then
           For Each p_codeWindow In p_workbook.VBProject.VBE.Windows
               If p_codeWindow.Visible = True And p_codeWindow.Type = 0 Then p_codeWindow.Visible = False
           Next p_codeWindow
        End If
    Next p_workbook

End Sub

''' <summary>   Handles the Workbook activate event. </summary>
''' <remarks>   Initializes the user defined errors. </remarks>
Private Sub Workbook_Activate()
    
    Const p_procedureName As String = "Workbook_Activate"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    Me.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    VBA.MsgBox Me.BuildStandardErrorMessage(), VBA.vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the workbook after save event. </summary>
''' <remarks>   Exports source code files. </remarks>
''' <param name="a_success">   [Boolean] True if saved successfully. </param>
Private Sub Workbook_AfterSave(ByVal a_success As Boolean)
    
    Const p_procedureName As String = "Workbook_AfterSave"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If a_success And Me.ExportCodeAfterSave Then
        cc_isr_Core_IO.WorkbookExport.ExportProjectCodeFiles Excel.ThisWorkbook.VBProject
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    Me.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    VBA.MsgBox Me.BuildStandardErrorMessage(), VBA.vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the workbook before close event. </summary>
''' <remarks>   Disables the save dialog for deployed or read-only workbooks.
'''             Disposes any disposable worksheets. </remarks>
''' <para name="a_cancel">   [Boolean] Set to true to cancel closing. </param>
Private Sub Workbook_BeforeClose(ByRef a_cancel As Boolean)
    
    Const p_procedureName As String = "Workbook_BeforeClose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' dispose the disposable worksheets
    
    ' :( this does not remove the code windows from the editor on the next
    ' opening.
    
    If AreCodeEditorWindowsOpen() Then Me.CloseAllEditorWindows
    
    ' this disables the save dialog on read only workbooks.
    Me.MarkAsSaved ThisWorkbook, Me.Deployed
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    Me.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    VBA.MsgBox Me.BuildStandardErrorMessage(), VBA.vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    Const p_procedureName As String = "Workbook_BeforeSave"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    Me.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    VBA.MsgBox Me.BuildStandardErrorMessage(), VBA.vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the workbook open event. </summary>
''' <remarks>   Toggles visibility of referenced workbooks. </remarks>
Private Sub Workbook_Open()
    
    Const p_procedureName As String = "Workbook_Open"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' turn on code exporting if not deploy or read only
    DisableSavingCode ThisWorkbook, Me.Deployed
    
    ' show this work book
    Excel.Application.Windows(Excel.ThisWorkbook.Name).Visible = True
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    Me.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    VBA.MsgBox Me.BuildStandardErrorMessage(), VBA.vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

