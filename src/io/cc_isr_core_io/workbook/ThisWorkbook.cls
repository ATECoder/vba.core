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

Private Type ThisType
    ExportCodeAfterSave As Boolean
End Type

Private This As ThisType

''' <summary>   Gets the option for exporting code files after save. </summary>
''' <value> </value>
Public Property Get ExportCodeAfterSave() As Boolean
    ExportCodeAfterSave = This.ExportCodeAfterSave
End Property

''' <summary>   Enables or disables exporting code files after save. </summary>
''' <para name="a_value">   [Boolean] True to enable exporting code files after save. </param>
Public Property Let ExportCodeAfterSave(ByVal a_value As Boolean)
    This.ExportCodeAfterSave = a_value
End Property

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
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    VBA.MsgBox cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(), VBA.vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the workbook after save event. </summary>
''' <remarks>   Exports source code files. </remarks>
''' <para name="a_success">   [Boolean] True if saved successfully. </param>
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
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    VBA.MsgBox cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(), VBA.vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Handles the workbook before close event. </summary>
''' <remarks>   Disposes the disposable workseets. </remarks>
''' <para name="a_cancel">   [Boolean] Set to true to cancel closing. </param>
Private Sub Workbook_BeforeClose(ByRef a_cancel As Boolean)
    
    Const p_procedureName As String = "Workbook_BeforeClose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' dispose the disposable worksheets
    
    ' :( this does not remove the code windows from the editor on the next
    ' opening.
    
    If AreCodeEditorWindowsOpen() Then
    
        ' close all editor windows.
        Me.CloseAllEditorWindows
    
        ' then save once again.
        If Not ThisWorkbook.ReadOnly Then
            ThisWorkbook.Save
        End If
    
    End If
    
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    VBA.MsgBox cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(), VBA.vbExclamation
    
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
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    VBA.MsgBox cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(), VBA.vbExclamation
    
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
    
    ' turn on code exporting for this work book
    Excel.ThisWorkbook.ExportCodeAfterSave = True
    
    Excel.Application.Windows(Excel.ThisWorkbook.Name).Visible = True
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.SetErrSource p_procedureName, Me.Name
    
    ' display the error message
    VBA.MsgBox cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(), VBA.vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

