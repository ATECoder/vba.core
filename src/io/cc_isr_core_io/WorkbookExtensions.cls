VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Workbook extensions. </summary>
''' <remarks> Dependencies: PathExtensions.cls, VbComponentExtesions.cls
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' Export project files
'
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>   Disables saving code if deployed or read only. </summary>
''' <param name="a_workbook">   [Excel.Workbook] The workbook. </param>
''' <param name="a_deployed">   [Boolean] True if the workbook was deployed. </param>
Public Sub DisableSavingCode(ByVal a_workbook As Excel.Workbook, ByVal a_deployed As Boolean)
    ' to be removed.
End Sub


''' <summary>   Marks a workbook as saved if not saved and deployed or read only. </summary>
''' <param name="a_workbook">   [Excel.Workbook] The workbook. </param>
''' <param name="a_deployed">   [Boolean] True if the workbook was deployed. </param>
Public Sub MarkAsSaved(ByVal a_workbook As Excel.Workbook, ByVal a_deployed As Boolean)
    
    If a_deployed Or a_workbook.ReadOnly Then _
        a_workbook.Saved = True

End Sub

''' <summary>   Hides the workbook. </summary>
''' <param name="a_workbook">   [Excel.Workbook] The workbook. </param>
Public Sub Hide(ByVal a_workbook As Excel.Workbook)
    
        Application.Windows(a_workbook.Name).Visible = False

End Sub

''' <summary>   Shows the workbook. </summary>
''' <param name="a_workbook">   [Excel.Workbook] The workbook. </param>
Public Sub Show(ByVal a_workbook As Excel.Workbook)
    
        Application.Windows(a_workbook.Name).Visible = True

End Sub


''' <summary>   Returns the workbooks that are refenced in the specified project. </summary>
''' <param name="a_workbook">   [Excel.Workbook] The workbook which project references are to be found.  </param>
''' <returns>   [Collection] of workbooks including this workbook. </returns>
Public Function ReferencedWorkbooks(ByVal a_workbook As Excel.Workbook) As Collection

    Dim p_workbooks As Collection
    Set p_workbooks = ReferenceManager.ReferencedWorkbooks(a_workbook.VBProject)
    p_workbooks.Add a_workbook
    Set ReferencedWorkbooks = p_workbooks

End Function

