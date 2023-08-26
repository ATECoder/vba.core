VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EditorExtensions"
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

