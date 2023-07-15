Attribute VB_Name = "WorkbookUtilitiesTests"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' WorkbookUtilitiesTests.bas
'
' Dependencies:
'
' Assert.cls
' MactroInfo.cls
' ModuleInfo.cls
' WorkbookUtilites.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Sub AddModule(ByVal a_col As VBA.Collection, ByVal a_moduleFullName As String)
    
    Dim p_module As ModuleInfo
    Set p_module = Constructor.CreateModuleInfo
    p_module.FromModuleFullName a_moduleFullName
    a_col.Add p_module

End Sub

Public Function ContainsModule(ByVal a_col As VBA.Collection, ByVal a_findModule As ModuleInfo) As Boolean
    
    Dim p_found As Boolean
    p_found = False
    Dim p_moduleInfo As ModuleInfo
    For Each p_moduleInfo In a_col
        DoEvents
        If p_moduleInfo.Equals(a_findModule) Then
            p_found = True
            Exit For
        End If
    Next p_moduleInfo
    ContainsModule = p_found

End Function

Private Function ContainsAllModules(ByVal a_leftCol As VBA.Collection, ByVal a_rightCol As VBA.Collection)

    Dim p_result As Boolean: p_result = False
    Dim p_rightModuleInfo As ModuleInfo
    For Each p_rightModuleInfo In a_rightCol
        DoEvents
        If Not ContainsModule(a_leftCol, p_rightModuleInfo) Then
            p_result = False
            Exit Function
        End If
    Next p_rightModuleInfo
    ContainsAllModules = p_result

End Function


''' <summary>   Adds the test modules. </summary>
Private Sub AddTestModules(ByVal a_knownTestModules As VBA.Collection)
    
    Dim p_projectName As String: p_projectName = Application.ActiveWorkbook.VBProject.name
    AddModule a_knownTestModules, p_projectName & ".CollectionExtensionsTests"
    AddModule a_knownTestModules, p_projectName & ".MarshalTests"
    AddModule a_knownTestModules, p_projectName & ".PathExtensionsTests"
    AddModule a_knownTestModules, p_projectName & ".StopWatchTests"
    AddModule a_knownTestModules, p_projectName & ".StringBuilderTests"
    AddModule a_knownTestModules, p_projectName & ".StringExtensionsTests"
    AddModule a_knownTestModules, p_projectName & ".UserDefinedErrorsTests"
    AddModule a_knownTestModules, p_projectName & ".WorkbookUtilitiesTests"

End Sub

''' <summary>   Unit test. Asserts creating a list of test modules. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestModuleList() As Assert

    Dim p_modules As VBA.Collection
    Set p_modules = WorkbookUtilities.EnumerateProjectModules(Application.ActiveWorkbook.VBProject)
    
    ' this includes all modules that start with test.
    Dim p_knownTestModules As VBA.Collection
    Set p_knownTestModules = New VBA.Collection
    AddTestModules p_knownTestModules
    
    Set TestModuleList = Assert.AreEqual(p_knownTestModules.count, p_modules.count, _
        "Expecting " & CStr(p_knownTestModules.count) & " but found  " & _
        CStr(p_modules.count) & " test modules")
    
    If Not TestModuleList.AssertSuccessful Then
        Exit Function
    End If
    
    Dim p_missingItem As Variant: Set p_missingItem = Nothing
    Set p_missingItem = CollectionExtensions.FindMissingItem(p_modules, p_knownTestModules)
    
    If Not p_missingItem Is Nothing Then
        Set TestModuleList = Assert.IsTrue(CollectionExtensions.ContainsAll(p_modules, p_knownTestModules), _
            "item " & CStr(p_missingItem) & " from the expected test module is not found in the actual collection of test modules")
        Exit Function
    End If
  
    Set p_missingItem = CollectionExtensions.FindMissingItem(p_knownTestModules, p_modules)
    
    If Not p_missingItem Is Nothing Then
        Set TestModuleList = Assert.IsTrue(CollectionExtensions.ContainsAll(p_modules, p_knownTestModules), _
            "item " & CStr(p_missingItem) & " from the actual test module is not found in the exected collection of test modules")
        Exit Function
    End If
  
  
End Function

