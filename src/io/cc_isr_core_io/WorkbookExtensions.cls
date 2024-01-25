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

Private Type this_
    RevisionPropertyName As String
    DocuemntVersionName As String
    CreationDateName As String
End Type

Private This As this_

Private Sub Class_Initialize()

    This.RevisionPropertyName = "Content Status"
    This.DocuemntVersionName = "Document Version"
    This.CreationDateName = "Creation date"

End Sub

''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Document properties
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

''' <summary>   Gets the revision document property name. </summary>
''' <value>   [String] </value>
Public Property Get RevisionPropertyName() As String
    RevisionPropertyName = This.RevisionPropertyName
End Property

''' <summary>   Gets the document version document property name. </summary>
''' <value>   [String] </value>
Public Property Get DocuemntVersionName() As String
    DocuemntVersionName = This.DocuemntVersionName
End Property

''' <summary>   Gets the creaetion date document property name. </summary>
''' <value>   [String] </value>
Public Property Get CreationDateName() As String
    CreationDateName = This.CreationDateName
End Property

''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Export project files
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
''' <param name="a_workbook">   [Excel.Workbook] The workbook which project references are to be found. </param>
''' <returns>   [Collection] of workbooks including this workbook. </returns>
Public Function ReferencedWorkbooks(ByVal a_workbook As Excel.Workbook) As Collection

    Dim p_workbooks As Collection
    Set p_workbooks = ReferenceManager.ReferencedWorkbooks(a_workbook.VBProject)
    p_workbooks.Add a_workbook
    Set ReferencedWorkbooks = p_workbooks

End Function

''' <summary>   Update a workbook document property. </summary>
''' <param name="a_workbook">   [Excel.Workbook] The workbook which project references are to be found. </param>
''' <param name="a_name">       [String] The document property name. </param>
''' <param name="a_value">      [Variant] The document property value. </param>
Public Sub UpdateDocumentProperty(ByVal a_workbook As Excel.Workbook, _
    ByVal a_name As String, ByVal a_value As Variant)
    If a_value <> ActiveWorkbook.BuiltinDocumentProperties(a_name) Then
        ActiveWorkbook.BuiltinDocumentProperties(a_name) = a_value
    End If
End Sub

''' <summary>   Debug prints the Excel workbook document properties. </summary>
''' <param name="a_workbook">   [Excel.Workbook] The workbook which project references are to be found. </param>
Public Sub DebugPrintDocumentProperties(ByVal a_workbook As Excel.Workbook)
    
    On Error Resume Next
    Dim p_item As Variant
    Dim p_itemNumber As Integer: p_itemNumber = 0
    Debug.Print a_workbook.Name, a_workbook.VBProject.Name
    For Each p_item In a_workbook.BuiltinDocumentProperties
        p_itemNumber = p_itemNumber + 1
        Debug.Print p_itemNumber, p_item.Name, VBA.TypeName(p_item), p_item.Value
        If Err.Number <> 0 Then
            Err.Clear
            Debug.Print p_itemNumber, p_item.Name, VBA.TypeName(p_item), "error producing value"
            If Err.Number <> 0 Then
                Err.Clear
                Debug.Print p_itemNumber, p_item.Name, "error printing type"
                If Err.Number <> 0 Then
                    Err.Clear
                    Debug.Print p_itemNumber, "error printing name"
                End If
            End If
        End If
    Next
    On Error GoTo 0
   
End Sub

'  1            Title         DocumentProperty            isr.cc Core Demo Visual Basic for Application Workbook
'  2            Subject       DocumentProperty            Visual Basic for Application Core Workbook Demo
'  3            Author        DocumentProperty            david@integratedscientificresources.com
'  4            Keywords      DocumentProperty
'  5            Comments      DocumentProperty            (c) 2023 Integrated Scientific Resources, Inc. All rights reserved.
' Integrated Scientific Resources, Inc., and any contributors grant you a license to the documentation and other content in this program under the [Creative Commons Attribution 4.0 International Public License], and grant you a license to any code in the program under the [MIT License].
'
' Integrated Scientific Resources, Inc., and/or other Integrated Scientific Resources, Inc., products and services referenced in the documentation may be either trademarks or registered trademarks of Integrated Scientific Resources, Inc., in the United States and/or other countries. The licenses for this project do not grant you rights to use any Integrated Scientific Resources, Inc., names, logos, or trademarks.
'
' Integrated Scientific Resources, Inc., and any contributors reserve all other rights, whether under their respective copyrights, patents, or trademarks, whether by implication, estoppel or otherwise.
'
' [Creative Commons Attribution 4.0 International Public License]:(https://creativecommons.org/licenses/by/4.0/legalcode)
' [MIT License]:(https://opensource.org/licenses/MIT)
'  6            Template      DocumentProperty
'  7            Last author   DocumentProperty            David Hary
'  8            Revision number             DocumentProperty
'  9            Application name            DocumentProperty            Microsoft Excel
'  10           Last print date             DocumentProperty            4/23/1999 12:59:39 AM
'  11           Creation date DocumentProperty            7/16/1998 3:59:47 AM
'  12           Last save time              DocumentProperty            10/23/2023 3:35:20 PM
'  13           Total editing time          DocumentProperty            error producing value
'  14           Number of pages             DocumentProperty            error producing value
'  15           Number of words             DocumentProperty            error producing value
'  16           Number of characters        DocumentProperty            error producing value
'  17           Security      DocumentProperty             2
'  18           Category      DocumentProperty            Visual Basic For Application
'  19           Format        DocumentProperty
'  20           Manager       DocumentProperty            David
'  21           Company       DocumentProperty            Integrated Scientific Resources, Inc.
'  22           Number of bytes             DocumentProperty            error producing value
'  23           Number of lines             DocumentProperty            error producing value
'  24           Number of paragraphs        DocumentProperty            error producing value
'  25           Number of slides            DocumentProperty            error producing value
'  26           Number of notes             DocumentProperty            error producing value
'  27           Number of hidden Slides     DocumentProperty            error producing value
'  28           Number of multimedia clips  DocumentProperty            error producing value
'  29           Hyperlink base              DocumentProperty
'  30           Number of characters (with spaces)        DocumentProperty            error producing value
'  31           Content type  DocumentProperty
'  32           Content status              DocumentProperty            Revision 1.0.8619
'  33           Language      DocumentProperty
'  34           Document version            DocumentProperty


