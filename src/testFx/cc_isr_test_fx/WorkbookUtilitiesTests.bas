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

''' <summary>   Returns true if the collection contains the specified key. </summary>
''' <param name="a_col">     [Collection] The subject collection. </param>
''' <param name="a_key">     [Variant] The key to check for in the collection. </param>
''' <returns>   True if the key is contained in the collection. </returns>
Public Function ContainsKey(ByVal a_col As VBA.Collection, ByVal a_key As Variant) As Boolean
    
    Dim p_found As Boolean
    p_found = False
    Dim colItem As Variant
    For Each colItem In a_col
        VBA.DoEvents
        If colItem = a_key Then
            p_found = True
            Exit For
        End If
    Next colItem
    ContainsKey = p_found

End Function

''' <summary>   Returns true if the object is contained in the collection. </summary>
''' <param name="a_col">         [Collection] The subject collection. </param>
''' <param name="a_contained">   [Collection] The collection which to check for being contained in the
'''                              subject collection. </param>
''' <returns>   True if the contained collection is fully contained in the collection. </returns>
Public Function ContainsAll(ByVal a_col As VBA.Collection, ByVal a_contained As VBA.Collection) As Boolean
    
    Dim p_result As Boolean: p_result = True
    Dim p_key As Variant
    For Each p_key In a_contained
        VBA.DoEvents
        If Not ContainsKey(a_col, p_key) Then
            p_result = False
            Exit For
        End If
    Next p_key
    ContainsAll = p_result
    
End Function

''' <summary>   Returns the first item that exists in <paramref name="a_contained"/>
''' not existing in <paramref name="a_col"/>. </summary>
''' <param name="a_col">         [Collection] The subject collection. </param>
''' <param name="a_contained">   [Collection] The collection which to check for being contained in the
'''                              subject collection. </param>
Public Function FindMissingItem(ByVal a_col As VBA.Collection, ByVal a_contained As VBA.Collection) As Variant
    
    Dim p_result As Variant: Set p_result = Nothing
    Dim p_key As Variant
    For Each p_key In a_contained
        VBA.DoEvents
        If Not ContainsKey(a_col, p_key) Then
            p_result = p_key
            Exit For
        End If
    Next p_key
    Set FindMissingItem = p_result
    
End Function

Private Sub AddModule(ByVal a_col As VBA.Collection, ByVal a_moduleFullName As String)
    
    Dim p_module As cc_isr_Test_Fx.ModuleInfo
    Set p_module = Constructor.CreateModuleInfo
    p_module.FromModuleFullName a_moduleFullName
    a_col.Add p_module

End Sub

Public Function ContainsModule(ByVal a_col As VBA.Collection, ByVal a_findModule As ModuleInfo) As Boolean
    
    Dim p_found As Boolean
    p_found = False
    Dim p_moduleInfo As cc_isr_Test_Fx.ModuleInfo
    For Each p_moduleInfo In a_col
        VBA.DoEvents
        If p_moduleInfo.Equals(a_findModule) Then
            p_found = True
            Exit For
        End If
    Next p_moduleInfo
    ContainsModule = p_found

End Function

Private Function ContainsAllModules(ByVal a_leftCol As VBA.Collection, ByVal a_rightCol As VBA.Collection)

    Dim p_result As Boolean: p_result = False
    Dim p_rightModuleInfo As cc_isr_Test_Fx.ModuleInfo
    For Each p_rightModuleInfo In a_rightCol
        VBA.DoEvents
        If Not ContainsModule(a_leftCol, p_rightModuleInfo) Then
            p_result = False
            Exit Function
        End If
    Next p_rightModuleInfo
    ContainsAllModules = p_result

End Function


''' <summary>   Adds the test modules. </summary>
Private Sub AddTestModules(ByVal a_knownTestModules As VBA.Collection)
    
    Dim p_projectName As String: p_projectName = Excel.Application.ActiveWorkbook.VBProject.Name
    AddModule a_knownTestModules, p_projectName & ".WorkbookUtilitiesTests"

End Sub

Public Sub BeforeAll()
    Debug.Print "@ Before All macro"
End Sub

Public Sub BeforeEach()
    Debug.Print "@ Before Each macro"
End Sub

Public Sub AfterEach()
    Debug.Print "@ After Each macro"
End Sub

Public Sub AfterAll()
    Debug.Print "@ After All macro"
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
    Set p_missingItem = FindMissingItem(p_modules, p_knownTestModules)
    
    If Not p_missingItem Is Nothing Then
        Set TestModuleList = Assert.IsTrue(ContainsAll(p_modules, p_knownTestModules), _
            "item " & CStr(p_missingItem) & " from the expected test module is not found in the actual collection of test modules")
        Exit Function
    End If
  
    Set p_missingItem = FindMissingItem(p_knownTestModules, p_modules)
    
    If Not p_missingItem Is Nothing Then
        Set TestModuleList = Assert.IsTrue(ContainsAll(p_modules, p_knownTestModules), _
            "item " & CStr(p_missingItem) & _
            " from the actual test module is not found in the exected collection of test modules")
        Exit Function
    End If
  
  
End Function

