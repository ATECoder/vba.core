VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MacroInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' MacroInfo.cls
'
' Dependencies:
'
' TBA
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private m_projectName As String
Private m_moduleName As String
Private m_macroName As String

''' <summary>   Initializes this class. </summary>
''' <param name="a_projectName">      [String] The project name. </param>
''' <param name="a_codeModuleName">   [String] The component code module name. </param>
''' <param name="a_macroName">        [String] The name of the macro. </param>
Public Sub Initialize(ByVal a_projectName As String, ByVal a_codeModuleName As String, ByVal a_macroName As String)
    m_projectName = a_projectName
    m_moduleName = a_codeModuleName
    m_macroName = a_macroName
End Sub

''' <summary>   Initializes this class. </summary>
''' <param name="a_module">      [<see cref="ModuleInfo"/>] The module info. </param>
''' <param name="a_macroName">   [String] The name of the macro. </param>
Public Sub InitializeModuelMacro(ByVal a_module As ModuleInfo, ByVal a_macroName As String)
    m_projectName = a_module.ProjectName
    m_moduleName = a_module.ModuleName
    m_macroName = a_macroName
End Sub

''' <summary>   Initializes this class from a macro full name. </summary>
''' <param name="a_macroFullName">     [String] the macro full name. </param>
Public Sub FromMacroFullName(ByVal a_macroFullName As String)
    
    Dim p_names() As String: p_names() = Split(a_macroFullName, ".")
    m_projectName = p_names(LBound(p_names))
    m_moduleName = p_names(LBound(p_names) + 1)
    m_macroName = p_names(LBound(p_names) + 2)

End Sub

''' <summary>   Creates a clone of the this macro info. </summary>
Public Function Clone() As cc_isr_Test_Fx.MacroInfo
    
    Set Clone = New cc_isr_Test_Fx.MacroInfo
    Clone.Initialize Me.ProjectName, Me.ModuleName, Me.MacroName

End Function

''' <summary>   Gets the macro name. </summary>
Public Property Get MacroName() As String
    
    MacroName = m_macroName

End Property

''' <summary>   Gets the project name. </summary>
Public Property Get ProjectName() As String
    
    ProjectName = m_projectName

End Property

''' <summary>   Gets the module name. </summary>
Public Property Get ModuleName() As String
    
    ModuleName = m_moduleName

End Property

''' <summary>   Gets the full name. </summary>
Public Property Get FullName() As String
Attribute FullName.VB_Description = "Get the macro full name"
Attribute FullName.VB_UserMemId = 0
    
    FullName = m_projectName & "." & m_moduleName & "." & m_macroName

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

