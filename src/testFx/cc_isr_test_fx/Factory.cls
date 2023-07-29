VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Instantiates public not creatable objects for external workbooks. </summary>
''' <remarks>   Dependencies: Assert.cls, MacroInfo.cls, ModuleInfo.cls, StopWatch.cls. </remakrs>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   A public creatable singleton class for creating instances of this workbook
''' classes to share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Instantiates a new <see cref="Assert"/> object. </summary>
''' <returns>   A new instance of the <see cref="Assert"/> object. </returns>
Public Function NewAssert() As Assert
  Set NewAssert = New Assert
End Function

''' <summary>   Instantiates a new <see cref="MacroInfo"/> object. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="MacroInfo"/> object. </returns>
Public Function NewMacroInfo() As MacroInfo
    Set NewMacroInfo = New MacroInfo
End Function

''' <summary>   Instantiates a new <see cref="ModuleInfo"/> object. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="ModuleInfo"/> object. </returns>
Public Function NewModuleInfo() As ModuleInfo
    Set NewModuleInfo = New ModuleInfo
End Function

''' <summary>   Instantiates a new <see cref="StopWatch"/> object. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
''' <returns>   A new instance of the <see cref="StopWatch"/> object. </returns>
Public Function NewStopWatch() As StopWatch
    Set NewStopWatch = New StopWatch
End Function

