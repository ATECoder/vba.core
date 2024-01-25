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
''' <summary>   A public creatable singleton class for creating instances of this workbook
'''             classes to share with external workbooks. </summary>
''' <remarks>   Dependencies: Assert.cls, MacroInfo.cls, ModuleInfo.cls. </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   A public creatable singleton class for creating instances of this workbook
'''             classes to share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Instantiates a new <see cref="Assert"/>] object. </summary>
''' <returns>   {<see cref="Assert"/>]. </returns>
Public Function NewAssert() As Assert
  Set NewAssert = New Assert
End Function

''' <summary>   Instantiates a new <see cref="MacroInfo"/>] object. </summary>
''' <returns>   {<see cref="MacroInfo"/>]. </returns>
Public Function NewMacroInfo() As MacroInfo
    Set NewMacroInfo = New MacroInfo
End Function

''' <summary>   Instantiates a new <see cref="ModuleInfo"/>] object. </summary>
''' <returns>   {<see cref="ModuleInfo"/>]. </returns>
Public Function NewModuleInfo() As ModuleInfo
    Set NewModuleInfo = New ModuleInfo
End Function

''' <summary>   Instantiates a new temporary test file. </summary>
''' <returns>   [<see cref="TestFileHandle"/>]. <returns>
Public Function NewTestFileHandle() As TestFileHandle
    Set NewTestFileHandle = New TestFileHandle
End Function

