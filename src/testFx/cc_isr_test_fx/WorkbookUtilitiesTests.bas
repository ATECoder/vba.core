Attribute VB_Name = "WorkbookUtilitiesTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Workbook utility tests.  </summary>
''' <remarks>   Dependencies: Assert.cls, MacroInfo.cls, ModuleInfo.cls, WorkbookUtilities.cls.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Name As String
    TestNumber As Integer
    BeforeAllAssert As Assert
    BeforeEachAssert As Assert
    ErrTracer As IErrTracer
End Type

Private This As this_

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Test runners
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Runs the specified test. </summary>
Public Sub RunTest(ByVal a_testNumber As Integer)
    BeforeEach
    Select Case a_testNumber
        Case 1
            TestModuleList
        Case Else
    End Select
    AfterEach
End Sub

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    BeforeAll
    RunTest 1
    AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    BeforeAll
    Dim p_testNumber As Integer
    For p_testNumber = 1 To 1
        RunTest p_testNumber
        DoEvents
    Next p_testNumber
    AfterAll
End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Tests initialize and cleanup.
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Prepares all tests. </summary>
''' <remarks>   This method sets up the 'Before All' <see cref="cc_isr_Test_Fx.Assert"/>
''' which serves to set the 'Before Each' <see cref="cc_isr_Test_Fx.Assert"/>.
''' The error object and user defined errors state are left clear after this method. </remarks>
Public Sub BeforeAll()

    Const p_procedureName As String = "BeforeAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = Assert.Pass("Primed to run all tests.")

    This.Name = "WorkbookUtilitiesTests"
    
    Set This.ErrTracer = New ErrTracer
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState

    ' Prime all tests

    This.TestNumber = 0
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        
        ' report any leftover errors.
        If Not This.ErrTracer Is Nothing Then _
            Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        
        If p_outcome.AssertSuccessful Then
            Set p_outcome = Assert.Pass("Primed to run all tests.")
        Else
            Set p_outcome = Assert.Inconclusive("Failed priming all tests;" & _
                VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    End If
    
    Set This.BeforeAllAssert = p_outcome
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Prepares each test before it is run. </summary>
''' <remarks>   This method sets up the 'Before Each' <see cref="cc_isr_Test_Fx.Assert"/>
''' which serves to initialize the <see cref="cc_isr_Test_Fx.Assert"/> of each test.
''' The error object and user defined errors state are left clear after this method. </remarks>
Public Sub BeforeEach()

    Const p_procedureName As String = "BeforeEach"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler

    This.TestNumber = This.TestNumber + 1

    Dim p_outcome As cc_isr_Test_Fx.Assert

    If This.BeforeAllAssert.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber) & ".")
    Else
        Set p_outcome = Assert.Inconclusive("Unable to prime pre-test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & This.BeforeAllAssert.AssertMessage)
    End If
    
    ' clear the error state.
    cc_isr_Core_IO.UserDefinedErrors.ClearErrorState
   
    ' Prepare the next test

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful Then
        
        ' report any leftover errors.
        If Not This.ErrTracer Is Nothing Then _
            Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
        
        If p_outcome.AssertSuccessful Then
             Set p_outcome = Assert.Pass("Primed pre-test #" & VBA.CStr(This.TestNumber))
        Else
            Set p_outcome = Assert.Inconclusive("Failed priming pre-test #" & VBA.CStr(This.TestNumber) & _
                ";" & VBA.vbCrLf & p_outcome.AssertMessage)
        End If
    End If
    
    Set This.BeforeEachAssert = p_outcome

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
                       
End Sub

''' <summary>   Releases test elements after each tests is run. </summary>
''' <remarks>   This method uses the <see cref="ErrTracer"/> to report any leftover errors
''' in the user defined errors queue and stack. The error object and user defined errors
''' state are left clear after this method. </remarks>
Public Sub AfterEach()
    
    Const p_procedureName As String = "AfterEach"
    
    ' Trap errors to the error handler.
    On Error GoTo err_Handler

    Dim p_outcome As cc_isr_Test_Fx.Assert
    Set p_outcome = Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")

    ' cleanup after each test.
    If This.BeforeEachAssert.AssertSuccessful Then
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before Each' assert.
    Set This.BeforeEachAssert = Nothing

    ' report any leftover errors.
    If Not This.ErrTracer Is Nothing Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
    Else
        Set p_outcome = Assert.Inconclusive("Errors reported cleaning up test #" & VBA.CStr(This.TestNumber) & _
            ";" & VBA.vbCrLf & p_outcome.AssertMessage)
    End If
    
    If Not p_outcome.AssertSuccessful Then _
        This.ErrTracer.TraceError p_outcome.AssertMessage
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

''' <summary>   Releases the test class after all tests run. </summary>
''' <remarks>   This method uses the <see cref="ErrTracer"/> to report any leftover errors
''' in the user defined errors queue and stack. The error object and user defined errors
''' state are left clear after this method. </remarks>
Public Sub AfterAll()
    
    Const p_procedureName As String = "AfterAll"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    Dim p_outcome As cc_isr_Test_Fx.Assert: Set p_outcome = Assert.Pass("All tests cleaned up.")
    
    ' cleanup after all tests.
    
    If This.BeforeAllAssert.AssertSuccessful Then
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    ' release the 'Before All' assert.
    Set This.BeforeAllAssert = Nothing

    ' report any leftover errors.
    If Not This.ErrTracer Is Nothing Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors()
    
    If p_outcome.AssertSuccessful Then
        Set p_outcome = Assert.Pass("Test #" & VBA.CStr(This.TestNumber) & " cleaned up.")
    Else
        Set p_outcome = Assert.Inconclusive("Errors reported cleaning up all tests;" & _
            VBA.vbCrLf & p_outcome.AssertMessage)
    End If
    
    If Not p_outcome.AssertSuccessful Then _
        This.ErrTracer.TraceError p_outcome.AssertMessage
    
    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Generic collection methods
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns true if the collection contains the specified key. </summary>
''' <param name="a_col">     [Collection] The subject collection. </param>
''' <param name="a_key">     [Variant] The key to check for in the collection. </param>
''' <returns> True if the key is contained in the collection. </returns>
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
''' <returns> True if the contained collection is fully contained in the collection. </returns>
Public Function ContainsAll(ByVal a_col As VBA.Collection, ByVal a_contained As VBA.Collection) As Boolean
    
    Dim p_result As Boolean: p_result = True
    Dim p_key As ModuleInfo
    For Each p_key In a_contained
        VBA.DoEvents
        If Not ContainsKey(a_col, p_key) Then
            p_result = False
            Exit For
        End If
    Next p_key
    ContainsAll = p_result
    
End Function

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Module Info collection methods
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Returns the first module that exists in <paramref name="a_contained"/> but is
''' not existing in <paramref name="a_col"/>. </summary>
''' <param name="a_col">         [Collection] The subject collection. </param>
''' <param name="a_contained">   [Collection] The collection which to check for being contained in the
'''                              subject collection. </param>
''' <returns>   [ModuleInfo] The missing module. </returns>
Public Function FindMissingModule(ByVal a_col As VBA.Collection, ByVal a_contained As VBA.Collection) As ModuleInfo
    
    Dim p_result As ModuleInfo: Set p_result = Nothing
    Dim p_key As ModuleInfo
    For Each p_key In a_contained
        VBA.DoEvents
        If Not ContainsKey(a_col, p_key) Then
            Set p_result = p_key
            Exit For
        End If
    Next p_key
    Set FindMissingModule = p_result
    
End Function

''' <summary>   Adds a module to the collectioon. </summary>
''' <param name="a_col">              [Collection] The module collection. </param>
''' <param name="a_moduleFullName">   [String] the module to add. </param>
Private Sub AddModule(ByVal a_col As VBA.Collection, ByVal a_moduleFullName As String)
    
    a_col.Add Factory.NewModuleInfo.FromModuleFullName(a_moduleFullName)

End Sub

''' <summary>   Returns true if the collection contains the specified module. </summary>
''' <param name="a_col">          [Collection] The subject collection. </param>
''' <param name="a_findModule">   [ModuleInfo] The module to find in the collection. </param>
''' <returns>   [Boolean] True if the module is contained in the collection. </returns>
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

''' <summary>   Returns true if the left collection contains all the modules of the right collection. </summary>
''' <param name="a_leftCol">    [Collection] the containing collection. </param>
''' <param name="a_rightCol">   [Collection] The collection which to check for being contained in the
'''                              containing collection. </param>
''' <returns>   [Boolean] True if the contained collection is fully contained in the collection. </returns>
Private Function ContainsAllModules(ByVal a_leftCol As VBA.Collection, ByVal a_rightCol As VBA.Collection)

    Dim p_result As Boolean: p_result = True
    Dim p_rightModuleInfo As cc_isr_Test_Fx.ModuleInfo
    For Each p_rightModuleInfo In a_rightCol
        VBA.DoEvents
        If Not ContainsModule(a_leftCol, p_rightModuleInfo) Then
            p_result = False
            Exit For
        End If
    Next p_rightModuleInfo
    ContainsAllModules = p_result

End Function

''' <summary>   Adds the known test modules. </summary>
''' <para name="a_knownTestModules">   [Collection] holds the know test modules. </param>
Private Sub AddTestModules(ByVal a_knownTestModules As VBA.Collection)
    
    Dim p_projectName As String: p_projectName = Excel.Application.ActiveWorkbook.VBProject.Name
    AddModule a_knownTestModules, p_projectName & ".WorkbookUtilitiesTests"
    AddModule a_knownTestModules, p_projectName & ".AssertTests"

End Sub

' + + + + + + + + + + + + + + + + + + + + + + + + + + +
'  Tests
' + + + + + + + + + + + + + + + + + + + + + + + + + + +

''' <summary>   Unit test. Asserts creating a list of test modules. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestModuleList() As Assert

    Const p_procedureName As String = "TestModuleList"

    ' Trap errors to the error handler
    On Error GoTo err_Handler

    Dim p_outcome As Assert
    Dim p_isDone As Boolean: p_isDone = False

    Dim p_modules As VBA.Collection
    Set p_modules = WorkbookUtilities.EnumerateProjectModules(Application.ActiveWorkbook.VBProject)
    
    ' this includes all modules that start with test.
    Dim p_knownTestModules As VBA.Collection
    Set p_knownTestModules = New VBA.Collection
    AddTestModules p_knownTestModules
    
    Set p_outcome = Assert.AreEqual(p_knownTestModules.Count, p_modules.Count, _
        "Expecting " & CStr(p_knownTestModules.Count) & " but found  " & _
        CStr(p_modules.Count) & " test modules")
    
    Dim p_missingModule As ModuleInfo

    If Not p_isDone And p_outcome.AssertSuccessful Then
    
        Set p_missingModule = FindMissingModule(p_modules, p_knownTestModules)
    
        If Not p_missingModule Is Nothing Then
            Set p_outcome = Assert.IsTrue(ContainsAllModules(p_modules, p_knownTestModules), _
                "Module " & p_missingModule.ModuleName & _
                " from the expected test modules is not found in the actual collection of test modules")
            p_isDone = True
        End If
    
    End If
    
    If Not p_isDone And p_outcome.AssertSuccessful Then
    
        Set p_missingModule = FindMissingModule(p_knownTestModules, p_modules)
        
        If Not p_missingModule Is Nothing Then
            Set TestModuleList = Assert.IsTrue(ContainsAllModules(p_modules, p_knownTestModules), _
                "Module " & p_missingModule.ModuleName & _
                " from the actual test module is not found in the expected collection of test modules")
            p_isDone = True
        End If
    
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    If p_outcome.AssertSuccessful And Not This.ErrTracer Is Nothing Then _
        Set p_outcome = This.ErrTracer.AssertLeftoverErrors
    
    Debug.Print p_outcome.BuildReport("TestModuleList")
    
    Set TestModuleList = p_outcome
    
    On Error GoTo 0
    Exit Function

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, This.Name, ThisWorkbook
    
    ' enqueue the error or append its source to the last error.
    cc_isr_Core_IO.UserDefinedErrors.EnqueueErrorObject
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
  
End Function

