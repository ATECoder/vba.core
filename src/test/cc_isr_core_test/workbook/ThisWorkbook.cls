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

    On Error Resume Next

    ' toggle deploy mode on referenced workbooks.

    cc_isr_Core_IO.ThisWorkbook.Deployed = a_value
    cc_isr_Core.ThisWorkbook.Deployed = a_value
    cc_isr_Test_Fx.ThisWorkbook.Deployed = a_value
    
End Property

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

''' <summary>   Handles the workbook after save event. </summary>
''' <remarks>   Exports source code files. </remarks>
''' <para name="a_success">   [Boolean] True if saved successfully. </param>
Private Sub Workbook_AfterSave(ByVal a_success As Boolean)
    
    Const p_procedureName As String = "Workbook_AfterSave"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If a_success And Me.ExportCodeAfterSave Then
        cc_isr_Core_IO.WorkbookExport.ExportProjectCodeFiles ThisWorkbook.VBProject
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    MsgBox ErrorMessageBuilder.BuildStandardErrorMessage(), vbExclamation
    
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
    
    ' this disables the save dialog on read only workbooks.
    Me.MarkAsSaved cc_isr_Core_IO.ThisWorkbook, Me.Deployed
    Me.MarkAsSaved cc_isr_Core.ThisWorkbook, Me.Deployed
    Me.MarkAsSaved cc_isr_Test_Fx.ThisWorkbook, Me.Deployed
    Me.MarkAsSaved ThisWorkbook, Me.Deployed
    
    ' dispose the disposable worksheets
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    MsgBox ErrorMessageBuilder.BuildStandardErrorMessage(), vbExclamation
    
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
    
    ' toggle the deployed state
    
    Me.Deployed = False
    
    ' turn on code exporting if not deploy or read only
    Me.DisableSavingCode cc_isr_Core_IO.ThisWorkbook, Me.Deployed
    Me.DisableSavingCode cc_isr_Core.ThisWorkbook, Me.Deployed
    Me.DisableSavingCode cc_isr_Test_Fx.ThisWorkbook, Me.Deployed
    Me.DisableSavingCode ThisWorkbook, Me.Deployed

    ' hide the referenced workbooks
    Application.Windows(cc_isr_Core_IO.ThisWorkbook.Name).Visible = False
    Application.Windows(cc_isr_Core.ThisWorkbook.Name).Visible = False
    Application.Windows(cc_isr_Test_Fx.ThisWorkbook.Name).Visible = False
    
    ' show this work book
    Application.Windows(ThisWorkbook.Name).Visible = True
    
    ' init the user defined errors.
    cc_isr_Core.UserDefinedErrors.Initialize
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    MsgBox ErrorMessageBuilder.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

