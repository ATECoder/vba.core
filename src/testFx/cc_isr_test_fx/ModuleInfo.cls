VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModuleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Module information for enumerating modules. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    ProjectName As String
    ModuleName As String
End Type

Private This As this_

''' <summary>   Initializes and returns this class. </summary>
''' <param name="a_projectName">      [String] The project name. </param>
''' <param name="a_codeModuleName">   [String] The component code module name. </param>
''' <returns>   [<see cref="ModuleInfo"/>]. </returns>
Public Function Initialize(ByVal a_projectName As String, ByVal a_codeModuleName As String) As ModuleInfo
    This.ProjectName = a_projectName
    This.ModuleName = a_codeModuleName
    Set Initialize = Me
End Function

''' <summary>   Initializes and returns this class from a module full name. </summary>
''' <param name="a_moduleFullName">   [String] The module full name. </param>
''' <returns>   [<see cref="ModuleInfo"/>]. </returns>
Public Function FromModuleFullName(ByVal a_moduleFullName As String) As ModuleInfo
    
    Dim p_names() As String: p_names() = Split(a_moduleFullName, ".")
    This.ProjectName = p_names(LBound(p_names))
    This.ModuleName = p_names(LBound(p_names) + 1)
    Set FromModuleFullName = Me

End Function

''' <summary>   Initializes and returns this class from a macro full name. </summary>
''' <param name="a_macroFullName">   [String] The macro full name. </param>
''' <returns>   [<see cref="ModuleInfo"/>]. </returns>
Public Function FromMacroFullName(ByVal a_macroFullName As String) As ModuleInfo
    
    Dim p_names() As String: p_names() = Split(a_macroFullName, ".")
    This.ProjectName = p_names(LBound(p_names))
    This.ModuleName = p_names(LBound(p_names) + 1)
    Set FromMacroFullName = Me
    
End Function

''' <summary>   Creates a clone of the this module info. </summary>
''' <returns>   [<see cref="ModuleInfo"/>]. </returns>
Public Function Clone() As ModuleInfo
    
    Set Clone = Factory.NewModuleInfo.Initialize(Me.ProjectName, Me.ModuleName)

End Function

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
Attribute FullName.VB_Description = "This is a description of a member variable"
Attribute FullName.VB_UserMemId = 0
    FullName = This.ProjectName & "." & This.ModuleName
End Property

''' <summary>   Checks if the two modules are equal referring to the same
''  component in the project. </summary>
''' <param name="a_module">   [See cref="ModuleInfo"/>] The module to check for equality. </param>
Public Function Equals(ByVal a_module As ModuleInfo) As Boolean
    If IsObject(a_module) Then
        Equals = (0 = StrComp(Me.ProjectName, a_module.ProjectName)) And _
                 (0 = StrComp(Me.ModuleName, a_module.ModuleName))
    Else
        Equals = False
    End If
End Function


