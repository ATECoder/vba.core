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

Private Type this_
    Deployed As Boolean
    AllWorkbooks As Collection
End Type

Private This As this_

''' <summary>   Updates this workbook document info. </summary>
Private Sub UpdateDocumentInfo()
    On Error Resume Next
    
    Dim p_revision As String: p_revision = "1.0.8696"
    Dim p_created As Variant: p_created = VBA.DateValue("7/15/2023 10:01:01 AM")
    
    cc_isr_Core_IO.WorkbookExtensions.UpdateDocumentProperty _
        ThisWorkbook, cc_isr_Core_IO.WorkbookExtensions.RevisionPropertyName, "Revision " & p_revision
    
    cc_isr_Core_IO.WorkbookExtensions.UpdateDocumentProperty _
        ThisWorkbook, cc_isr_Core_IO.WorkbookExtensions.CreationDateName, p_created
    
    cc_isr_Core_IO.WorkbookExtensions.UpdateDocumentProperty _
        ThisWorkbook, cc_isr_Core_IO.WorkbookExtensions.DocuemntVersionName, p_revision

    On Error GoTo 0
End Sub

''' <summary>   Returns true if the collection is nothing or empty. </summary>
''' <param name="a_collection">   [Collection] the collection. </param>
''' <returns>   [Boolean]. </value>
Public Function CollectionIsNothingOrEmpty(ByVal a_collection As Collection)
    If a_collection Is Nothing Then
        CollectionIsNothingOrEmpty = True
    ElseIf 0 = a_collection.Count Then
        CollectionIsNothingOrEmpty = True
    Else
        CollectionIsNothingOrEmpty = False
    End If
End Function

''' <summary>   Gets all referenced workbooks for this workbook including
'''             itself. </summary>
''' <value>   [Collection]. </value>
Public Property Get AllWorkbooks() As Collection
    If Me.CollectionIsNothingOrEmpty(This.AllWorkbooks) Then
        If ThisWorkbook Is Nothing Then
            Set This.AllWorkbooks = New Collection
        Else
            Set This.AllWorkbooks = cc_isr_Core_IO.WorkbookExtensions.ReferencedWorkbooks(ThisWorkbook)
        End If
    End If
    Set AllWorkbooks = This.AllWorkbooks
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

    On Error Resume Next

    ' toggle deploy mode on all workbooks.
    This.Deployed = a_value
    Dim p_workbook As Workbook
    For Each p_workbook In Me.AllWorkbooks
        If p_workbook.Name <> ThisWorkbook.Name Then _
            p_workbook.Deployed = a_value
        DoEvents
    Next p_workbook

End Property

''' <summary>   Handles the Workbook activate event. </summary>
''' <remarks>   Initializes the user defined errors. </remarks>
Private Sub Workbook_Activate()
    
    Const p_procedureName As String = "Workbook_Activate"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' show the active workbook
    Application.Windows(ThisWorkbook.Name).Visible = True
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    MsgBox cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the workbook after save event. </summary>
''' <para name="a_success">   [Boolean] True if the save succeeded. </param>
''' <remarks>   Exports source code files. </remarks>
Private Sub Workbook_AfterSave(ByVal a_success As Boolean)

    Const p_procedureName As String = "Workbook_AfterSave"
   
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' export project files.
    
    If a_success And Not (This.Deployed Or ThisWorkbook.ReadOnly) Then
        cc_isr_Core_IO.VbProjectExtensions.ExportCodeFiles ThisWorkbook.VBProject
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    MsgBox cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
   
End Sub

Private Sub Workbook_AfterXmlImport(ByVal Map As XmlMap, ByVal IsRefresh As Boolean, ByVal Result As XlXmlImportResult)

End Sub

''' <summary>   Handles the workbook before close event. </summary>
''' <remarks>   Disables the save dialog for deployed or read-only workbooks.
'''             Disposes any disposable worksheets. </remarks>
''' <para name="a_cancel">   [Boolean] Set to true to cancel closing. </param>
Private Sub Workbook_BeforeClose(ByRef a_cancel As Boolean)

    Const p_procedureName As String = "Workbook_BeforeClose"
   
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' disable the save dialog on read only workbooks.
    
    Dim p_workbook As Workbook
    For Each p_workbook In Me.AllWorkbooks
        cc_isr_Core_IO.WorkbookExtensions.MarkAsSaved p_workbook, Me.Deployed
        DoEvents
    Next p_workbook
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    MsgBox cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

''' <summary>   Handles the workbook before save event. </summary>
''' <remarks>   Updates the document info as necessary. </remarks>
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    ' update the document info to the current revision as necessary.
    UpdateDocumentInfo

End Sub

''' <summary>   Handles the workbook open event. </summary>
''' <remarks>   Toggles visibility of referenced workbooks. </remarks>
Private Sub Workbook_Open()

    Const p_procedureName As String = "Workbook_Open"
   
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' hide this and all referenced workbooks
    Dim p_workbook As Workbook
    For Each p_workbook In Me.AllWorkbooks
        cc_isr_Core_IO.WorkbookExtensions.Hide p_workbook
        DoEvents
    Next p_workbook
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    MsgBox cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub




