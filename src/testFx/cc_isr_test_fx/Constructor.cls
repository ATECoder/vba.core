VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Constructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' Constructor.cls
'
' Dependencies:
'
' Assert.cls
' MacroInfo.cls
' ModuleInfo.cls
' StopWatch.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   A public creatable singleton class for creating instances of this workbook
''' classes to share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Creates an instance of the <see cref="Assert"/> class. </summary>
''' <returns>   A new instance of the <see cref="Assert"/> class. </returns>
Public Function CreateAssert() As Assert
  Set CreateAssert = New Assert
End Function

''' <summary>   Creates a new instance of the <see cref="MacroInfo"/> class to share
''' with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="MacroInfo"/> class. </returns>
Public Function CreateMacroInfo() As MacroInfo
    Set CreateMacroInfo = New MacroInfo
End Function

''' <summary>   Creates a new instance of the <see cref="ModuleInfo"/> class to share
''' with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="ModuleInfo"/> class. </returns>
Public Function CreateModuleInfo() As ModuleInfo
    Set CreateModuleInfo = New ModuleInfo
End Function

''' <summary>   Creates a new instance of the <see cref="StopWatch"/> class to share
''' with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="StopWatch"/> class. </returns>
Public Function CreateStopWatch() As StopWatch
    Set CreateStopWatch = New StopWatch
End Function

