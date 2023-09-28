VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Enumerates modules and macros for testing. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Adds items from the source to the destination collection. </summary>
''' <param name="a_source">        [Collection] The source collection. </param>
''' <param name="a_destination">   [Collection] The destination collection which receives the added items. </param>
''' <returns>   The number of items that were added. </returns>
Public Function AddItems(ByVal a_source As VBA.Collection, ByVal a_destination As VBA.Collection) As Integer
    
    Dim p_count As Integer: p_count = 0
    Dim p_item As Variant
    For Each p_item In a_source
        VBA.DoEvents
        a_destination.Add p_item
        p_count = p_count + 1
    Next p_item
    AddItems = p_count
    
End Function

''' <summary>   Checks whether the given text starts with the given sequence. </summary>
''' <remarks>   </remarks>
''' <param name="a_text">            [String] The text to check for the sequence. </param>
''' <param name="a_startText">       [String] The text to be located at the start. </param>
''' <param name="a_compareMethod">   [VbCompareMethod, vbBinaryCompare] The compare method. </param>
''' <returns>   A Boolean <returns>
Public Function StartsWith(ByVal a_text As String, ByVal a_startText As String, _
                           Optional ByVal a_compareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
    
    StartsWith = 1 = InStr(1, a_text, a_startText, a_compareMethod)

End Function

''' <summary>   Checks whether the given text ends with the given sequence. </summary>
''' <remarks>   </remarks>
''' <param name="a_text">            [String] The text to check for the sequence. </param>
''' <param name="a_endText">         [String] The text to be located at the end. </param>
''' <param name="a_compareMethod">   [VbCompareMethod, vbBinaryCompare] The compare method. </param>
''' <returns>   A Boolean. <returns>
Public Function EndsWith(ByVal a_text As String, ByVal a_endText As String, _
                           Optional ByVal a_compareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
    
    EndsWith = 0 = StrComp(Right$(a_text, Len(a_endText)), a_endText, a_compareMethod)

End Function


''' <summary>   List all modules that end with the specified suffix and
''' contain methods that start with the specified prefix. <summary>
''' <param name="macroNamePrefix">    ['Test'] The macro name prefix. </param>
''' <param name="moduleNameSuffix">   ['Tests'] The module name suffix </param>
Public Function EnumerateModules(Optional ByVal a_macroNamePrefix As String = "Test", _
                                 Optional ByVal a_moduleNameSuffix As String = "Tests") As VBA.Collection

    Dim p_modules As VBA.Collection: Set p_modules = New VBA.Collection
    Dim p_vbProject As VBProject
    For Each p_vbProject In Application.VBE.VBProjects
        VBA.DoEvents
        Me.AddItems EnumerateProjectModules(p_vbProject, _
                                            a_macroNamePrefix, a_moduleNameSuffix), _
                                            p_modules
    Next p_vbProject
    Set EnumerateModules = p_modules

End Function

''' <summary>   List all project modules that end with the specified suffix and
''' contain methods that start with the specified prefix. <summary>
''' <param name="a_project">            [<see cref="VBProject"/> The project which modules are to be added. </param>
''' <param name="a_macroNamePrefix">    [String, 'Test'] The macro name prefix. </param>
''' <param name="a_moduleNameSuffix">   [String, 'Tests'] The module name suffix </param>
Public Function EnumerateProjectModules(ByVal a_project As VBProject, _
                                        Optional ByVal a_macroNamePrefix As String = "Test", _
                                        Optional ByVal a_moduleSuffix As String = "Tests") As VBA.Collection
    
    On Error Resume Next
    Dim p_modules As VBA.Collection: Set p_modules = New VBA.Collection
    Dim p_currentModule As ModuleInfo: Set p_currentModule = New ModuleInfo
    Dim p_newModule As ModuleInfo: Set p_newModule = New ModuleInfo
    
    Dim p_vbComponent As VBComponent
    For Each p_vbComponent In a_project.VBComponents
        VBA.DoEvents
        If IsObject(p_vbComponent) Then
            p_newModule.Initialize a_project.Name, p_vbComponent.CodeModule
            If Not p_newModule.Equals(p_currentModule) Then
                Set p_currentModule = p_newModule.Clone
                If Me.EndsWith(p_currentModule.ModuleName, a_moduleSuffix) Then
                    If HasMacros(p_vbComponent, a_macroNamePrefix) Then
                        p_modules.Add p_currentModule
                    End If
                End If
            End If
        End If
    Next
    Set EnumerateProjectModules = p_modules
End Function

''' <summary>   Checks if the <see cref="VbComponent"/> has macros. <summary>
''' <param name="a_component">   [<see cref="VBComponent"/>] The component to check for macro methods. </param>
''' <param name="a_prefix">      ['Test'] The prefix of the macro method name. </param>
Function HasMacros(ByVal a_component As VBComponent, Optional ByVal a_prefix As String = "Test") As Boolean

    Dim p_currentName As String, p_newName As String
    HasMacros = False
    
    On Error Resume Next
    p_currentName = vbNullString
    
    If Not a_component Is Nothing Then
        Dim p_lineNo As Integer
        For p_lineNo = 1 To a_component.CodeModule.CountOfLines
            p_newName = a_component.CodeModule.ProcOfLine(Line:=p_lineNo, prockind:=vbext_pk_Proc)
            If p_currentName <> p_newName Then
                p_currentName = p_newName
                If Me.StartsWith(p_currentName, a_prefix) Then
                   HasMacros = True
                   Exit For
                End If
            End If
        Next
    End If

End Function

''' <summary>   Finds a project. <summary>
''' <param name="a_projectName">   [String] The projectName. </param>
''' <returns>   A reference to the found project or nothing. </returns>
Public Function FindProjectByName(ByVal a_projectName As String) As VBProject
    
    Dim p_vbProject As VBProject
    For Each p_vbProject In Application.VBE.VBProjects
        VBA.DoEvents
        If 0 = StrComp(p_vbProject.Name, a_projectName, vbTextCompare) Then
            Set FindProjectByName = p_vbProject
            Exit For
        End If
    Next p_vbProject

End Function

''' <summary>   Finds a code module by name. <summary>
''' <param name="a_project">      [<see cref="VBProject"/> The project. </param>
''' <param name="a_moduleName">   [String] The component (e.g., module) name. </param>
''' <returns>   A reference to the component of found code module or nothing. </returns>
Public Function FindCodeModuleByName(ByVal a_project As VBProject, ByVal a_moduleName As String) As VBComponent
    
    If Not a_project Is Nothing Then
        Dim p_vbComponent As VBComponent
        For Each p_vbComponent In a_project.VBComponents
            VBA.DoEvents
            If 0 = StrComp(p_vbComponent.CodeModule, a_moduleName, vbTextCompare) Then
                Set FindCodeModuleByName = p_vbComponent
                Exit For
            End If
        Next p_vbComponent
    End If

End Function

''' <summary>   Enumerates all macros in the specified module. <summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/28132276/get-a-list-of-the-macros-of-a-module-in-excel-and-then-call-all-those-macros"/>
''' <remarks>
''' <param name="a_module">   [<see cref="ModuleInfo"/>] ]The module info of the module containing the macros. </param>
''' <param name="a_prefix">   ['Test'] The prefix of the macro name. </param>
''' <returns>   A Collection of macro names. </returns>
Public Function EnumerateMacros(ByVal a_module As ModuleInfo, Optional ByVal a_prefix As String = "Test") As VBA.Collection

    Dim p_macros As Collection: Set p_macros = New Collection
    If Not a_module Is Nothing Then
    
        Dim p_vbProject As VBProject: Set p_vbProject = FindProjectByName(a_module.ProjectName)
        
        If Not p_vbProject Is Nothing Then
        
            Dim p_vbComponent As VBComponent
            Set p_vbComponent = FindCodeModuleByName(p_vbProject, a_module.ModuleName)
            
            If Not p_vbComponent Is Nothing Then
            
                Set p_macros = EnumerateComponentMacros(a_module, p_vbComponent, a_prefix)
            
            End If
        
        End If
    
    End If
    Set EnumerateMacros = p_macros

End Function

''' <summary>   Enumerates all macros in the specified module. <summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/28132276/get-a-list-of-the-macros-of-a-module-in-excel-and-then-call-all-those-macros"/>
''' <remarks>
''' <param name="a_module">      [<see cref="ModuleInfo"/> The module info of the module containing the macros. </param>
''' <param name="a_prefix">      [String, 'Test'] The prefix of the macro name. </param>
''' <param name="a_component">   [<see cref="VBComponent"/>] The component to check for macro methods. </param>
''' <returns>   A Collection of macro names. </returns>
Public Function EnumerateComponentMacros(ByVal a_module As ModuleInfo, _
        ByVal a_component As VBComponent, _
        Optional ByVal a_prefix As String = "Test") As VBA.Collection

    Dim p_macros As VBA.Collection: Set p_macros = New VBA.Collection
    Dim p_currentName As String: p_currentName = vbNullString
    Dim p_newName As String
    Dim p_lineNo As Integer
    For p_lineNo = 1 To a_component.CodeModule.CountOfLines
        p_newName = a_component.CodeModule.ProcOfLine(Line:=p_lineNo, prockind:=vbext_pk_Proc)
        If p_currentName <> p_newName Then
            
            p_currentName = p_newName
            
            If Me.StartsWith(p_currentName, a_prefix) Then
                
                p_macros.Add Factory.NewMacroInfo.InitializeModuleMacro(a_module, p_currentName)
            
            End If
        End If
    Next
    Set EnumerateComponentMacros = p_macros

End Function

''' <summary>   Enumerates all macros in the specified module. <summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/28132276/get-a-list-of-the-macros-of-a-module-in-excel-and-then-call-all-those-macros"/>
''' <remarks>
''' <param name="a_module">    [<see cref="ModuleInfo"/>] ]The module info of the module containing the macros. </param>
''' <param name="a_methods">   [Out, TestMethodFlags] The test methods that are included in this test module. </param>
''' <returns>   A Collection of macro names. </returns>
Public Function EnumerateTestMacros(ByVal a_module As ModuleInfo, _
        ByRef a_methods As TestMethodFlags) As VBA.Collection

    Dim p_macros As Collection: Set p_macros = New Collection
    If Not a_module Is Nothing Then
    
        Dim p_vbProject As VBProject: Set p_vbProject = FindProjectByName(a_module.ProjectName)
        
        If Not p_vbProject Is Nothing Then
        
            Dim p_vbComponent As VBComponent
            Set p_vbComponent = FindCodeModuleByName(p_vbProject, a_module.ModuleName)
            
            If Not p_vbComponent Is Nothing Then
            
                Set p_macros = EnumerateComponentTestMacros(a_module, p_vbComponent, a_methods)
            
            End If
        
        End If
    
    End If
    Set EnumerateTestMacros = p_macros

End Function

''' <summary>   Enumerates all test macros in the specified module. <summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/28132276/get-a-list-of-the-macros-of-a-module-in-excel-and-then-call-all-those-macros"/>
''' <remarks>
''' <param name="a_module">      [<see cref="ModuleInfo"/> The module info of the module containing the macros. </param>
''' <param name="a_component">   [<see cref="VBComponent"/>] The component to check for macro methods. </param>
''' <param name="a_methods">     [Out, <see cref="TestMethodFlags"/>] The test method that were found in this test module. </param>
''' <returns>   A Collection of macro names. </returns>
Public Function EnumerateComponentTestMacros(ByVal a_module As ModuleInfo, _
        ByVal a_component As VBComponent, _
        ByRef a_methods As TestMethodFlags) As VBA.Collection

    a_methods = None
    Dim p_macros As VBA.Collection: Set p_macros = New VBA.Collection
    Dim p_currentName As String: p_currentName = vbNullString
    Dim p_newName As String
    Dim p_lineNo As Integer
    For p_lineNo = 1 To a_component.CodeModule.CountOfLines
        p_newName = a_component.CodeModule.ProcOfLine(Line:=p_lineNo, prockind:=vbext_pk_Proc)
        If p_currentName <> p_newName Then
            p_currentName = p_newName
            If Me.StartsWith(p_currentName, TestMethodPrefix) Then
                p_macros.Add Factory.NewMacroInfo.InitializeModuleMacro(a_module, p_currentName)
            ElseIf 1 = VBA.InStr(1, p_currentName, ConstantsEnums.BeforeAllMethodName, vbTextCompare) Then
                a_methods = a_methods Or TestMethodFlags.BeforeAll
            ElseIf 1 = VBA.InStr(1, p_currentName, ConstantsEnums.AfterAllMethodName, vbTextCompare) Then
                a_methods = a_methods Or TestMethodFlags.AfterAll
            ElseIf 1 = VBA.InStr(1, p_currentName, ConstantsEnums.BeforeEachMethodName, vbTextCompare) Then
                a_methods = a_methods Or TestMethodFlags.BeforeEach
            ElseIf 1 = VBA.InStr(1, p_currentName, ConstantsEnums.AfterEachMethodName, vbTextCompare) Then
                a_methods = a_methods Or TestMethodFlags.AfterEach
            End If
        End If
    Next
    Set EnumerateComponentTestMacros = p_macros

End Function



