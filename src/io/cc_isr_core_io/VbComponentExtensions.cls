VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbComponentExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Manages VBE Components. </summary>
''' <remarks> Dependencies: PathExtensions.cls.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Detects if a component is a workbook. <summary>
''' <param name="a_component">   [<see cref="VBComponent"/>] A VB Component. </param>
''' <returns> True if the component is a work book. </returns>
Public Function IsWorkbook(ByVal a_component As VBIDE.VBComponent)
    
    On Error Resume Next
    Dim p_property As VBIDE.Property: Set p_property = a_component.Properties("AutoSaveOn")
    On Error GoTo 0
    IsWorkbook = Not p_property Is Nothing

End Function

''' <summary>   Detects if a component is a sheet. <summary>
''' <param name="component">   A reference to a VB Component. </param>
''' <returns> True if the component is a sheet. </returns>
Public Function IsSheet(ByVal a_component As VBIDE.VBComponent)
    
    On Error Resume Next
    Dim p_property As VBIDE.Property: Set p_property = a_component.Properties("Cells")
    On Error GoTo 0
    IsSheet = Not p_property Is Nothing

End Function

''' <summary>   Detects if a component is a chart. <summary>
''' <param name="a_component">   [<see cref="VBComponent"/>] A VB Component. </param>
''' <returns> True if the component is a chart. </returns>
Public Function IsChart(ByVal a_component As VBIDE.VBComponent)
    
    On Error Resume Next
    Dim p_property As VBIDE.Property: Set p_property = a_component.Properties("AutoScaling")
    On Error GoTo 0
    IsChart = Not p_property Is Nothing

End Function

''' <summary>   Get the extension if the component is a file </summary>
''' <param name="a_component">   The <see cref="VBComponent"/>   </param>
''' <returns>   The file extension or an empty string if not a file. </returns>
Public Function GetFileExtension(ByVal a_component As VBComponent)
    
    Select Case a_component.Type
        Case VBIDE.vbext_ct_ClassModule, VBIDE.vbext_ct_Document
            GetFileExtension = ".cls"
        Case VBIDE.vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case VBIDE.vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = VBA.vbNullString
    End Select

End Function

