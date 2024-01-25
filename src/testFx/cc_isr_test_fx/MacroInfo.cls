VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MacroInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Macro information for enumerating macros.  </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    ProjectName As String
    ModuleName As String
    MacroName As String
End Type

Private This As this_

''' <summary>   Initializes and returns this class. </summary>
''' <param name="a_projectName">      [String] The project name. </param>
''' <param name="a_codeModuleName">   [String] The component code module name. </param>
''' <param name="a_macroName">        [String] The name of the macro. </param>
''' <returns>   [<see cref="MacroInfo"/>]. </returns>
Public Function Initialize(ByVal a_projectName As String, _
        ByVal a_codeModuleName As String, ByVal a_macroName As String) As MacroInfo
    This.ProjectName = a_projectName
    This.ModuleName = a_codeModuleName
    This.MacroName = a_macroName
    Set Initialize = Me
End Function

''' <summary>   Initializes and returns this class. </summary>
''' <param name="a_module">      [<see cref="ModuleInfo"/>] The module info. </param>
''' <param name="a_macroName">   [String] The name of the macro. </param>
''' <returns>   [<see cref="MacroInfo"/>]. </returns>
Public Function InitializeModuleMacro(ByVal a_module As ModuleInfo, _
        ByVal a_macroName As String) As MacroInfo
    This.ProjectName = a_module.ProjectName
    This.ModuleName = a_module.ModuleName
    This.MacroName = a_macroName
    Set InitializeModuleMacro = Me
End Function

''' <summary>   Initializes and returns this class from a macro full name. </summary>
''' <param name="a_macroFullName">     [String] the macro full name. </param>
''' <returns>   [<see cref="MacroInfo"/>]. </returns>
Public Function FromMacroFullName(ByVal a_macroFullName As String) As MacroInfo
    
    Dim p_names() As String: p_names() = Split(a_macroFullName, ".")
    This.ProjectName = p_names(LBound(p_names))
    This.ModuleName = p_names(LBound(p_names) + 1)
    This.MacroName = p_names(LBound(p_names) + 2)
    Set FromMacroFullName = Me

End Function

''' <summary>   Creates a clone of the this macro info. </summary>
''' <returns>   [<see cref="MacroInfo"/>]. </returns>
Public Function Clone() As MacroInfo
    
    Set Clone = Factory.NewMacroInfo.Initialize(Me.ProjectName, Me.ModuleName, Me.MacroName)

End Function

''' <summary>   Gets the macro name. </summary>
Public Property Get MacroName() As String
    
    MacroName = This.MacroName

End Property

''' <summary>   Gets the project name. </summary>
Public Property Get ProjectName() As String
    
    ProjectName = This.ProjectName

End Property

''' <summary>   Gets the module name. </summary>
Public Property Get ModuleName() As String
    
    ModuleName = This.ModuleName

End Property

''' <summary>   Gets the full name. </summary>
Public Property Get FullName() As String
Attribute FullName.VB_Description = "Get the macro full name"
Attribute FullName.VB_UserMemId = 0
    
    FullName = This.ProjectName & "." & This.ModuleName & "." & This.MacroName

End Property

''' <summary>   Checks if the two macros are equal referring to the same
''  macro in the project. </summary>
''' <param name="a_macro">   [<see cref=MacroInfo"/>] The macro to check for equality. </param>
Public Function Equals(ByVal a_macro As MacroInfo) As Boolean
    
    If VBA.IsObject(a_macro) Then
        Equals = (0 = StrComp(Me.ProjectName, a_macro.ProjectName)) And _
                 (0 = StrComp(Me.ModuleName, a_macro.ModuleName)) And _
                 (0 = StrComp(Me.MacroName, a_macro.MacroName))
    Else
        Equals = False
    End If

End Function

