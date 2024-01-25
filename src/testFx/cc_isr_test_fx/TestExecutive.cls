VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestExecutive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An executive for executing test methods. </summary>
''' <remarks>   Dependencies: MacroInfo.cls, ModuleInfo.cls, WorkbookUtilities.cls. </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit
Private Const ModuleFullNameRow As String = "1"
Private Const ModuleFullNameColumn As String = "B"
Private Const LabelColumn As String = "A"
Private Const ItemColumn As String = "B"
Private Const ProjectLabel As String = "Project"
Private Const ModuleLabel As String = "Module"

Private Const TestNameColumn As String = "A"
Private Const OutcomeColumn As String = "B"
Private Const ClearColorIndex As Integer = 0
Private Const PassColorIndex As Integer = 4
Private Const PassLabel As String = "Passed"
Private Const PassOutcomeLabel As String = "passed"
Private Const FailColorIndex As Integer = 3
Private Const FailLabel As String = "Failed"
Private Const FailOutcomeLabel As String = "failed"
Private Const TestNameLabel As String = "Test Name"
Private Const TestOutcomeLabel As String = "Outcome"
Private Const TestDurationLabel As String = "Duration"
Private Const TestDurationUnits As String = " ms"
Private Const InconclusiveColorIndex As Integer = 6
Private Const InconclusiveOutcomeLabel As String = "inconclusive"
Private Const InconclusiveLabel As String = "Inconclusive"

''' <summary>   Executes all test macros in the module specified in the test sheet. </summary>
''' <para name="a_testSheet">             [Worksheet] Specifies the work sheet were test data are set and logged. </param>
''' <para name="a_usingRemoteWorkbook">   [Boolean, True] True if using a remote workbook
'''                                       where workbook file needs to be added
''' to the application run call. </param>
Public Sub Execute(ByVal a_testSheet As Worksheet, Optional a_usingRemoteWorkbook As Boolean = False)
    
    On Error GoTo Finally
    
    If a_usingRemoteWorkbook Then
        Dim p_workbook As Workbook: Set p_workbook = a_testSheet.Parent
        Dim p_filePath As String: p_filePath = p_workbook.FullName
    End If

    Dim p_moduleFullName As String
   
    Dim p_inconclusiveCount As Integer
    Dim p_passedCount As Integer
    Dim p_failedCount As Integer
    Dim p_rowNo As Integer
    p_rowNo = ModuleFullNameRow
    
    a_testSheet.Rows(CStr(ModuleFullNameRow + 1) & ":" & a_testSheet.Rows.Count).ClearContents
    a_testSheet.Range(OutcomeColumn & CStr(ModuleFullNameRow), _
                      OutcomeColumn & a_testSheet.Rows.Count).Interior.ColorIndex = ClearColorIndex
    
    p_moduleFullName = a_testSheet.Range(ModuleFullNameColumn & ModuleFullNameRow).value
    
    ' get the module info
    Dim p_moduleInfo As ModuleInfo
    Set p_moduleInfo = Factory.NewModuleInfo.FromModuleFullName(p_moduleFullName)
    
    p_rowNo = p_rowNo + 1
    a_testSheet.Range(LabelColumn & p_rowNo).value = ProjectLabel
    a_testSheet.Range(ItemColumn & p_rowNo).value = p_moduleInfo.ProjectName
    
    p_rowNo = p_rowNo + 1
    a_testSheet.Range(LabelColumn & p_rowNo).value = ModuleLabel
    a_testSheet.Range(ItemColumn & p_rowNo).value = p_moduleInfo.ModuleName
   
    p_rowNo = p_rowNo + 1
    Dim p_passCountRow As Integer: p_passCountRow = p_rowNo
    a_testSheet.Range(LabelColumn & p_passCountRow).value = PassLabel
    
    p_rowNo = p_rowNo + 1
    Dim p_failCountRow As Integer: p_failCountRow = p_rowNo
    a_testSheet.Range(LabelColumn & p_failCountRow).value = FailLabel
    
    p_rowNo = p_rowNo + 1
    Dim p_inconclusiveCountRow As Integer: p_inconclusiveCountRow = p_rowNo
    a_testSheet.Range(LabelColumn & p_inconclusiveCountRow).value = InconclusiveLabel

    p_rowNo = p_rowNo + 1
    Dim p_durationRow As Integer: p_durationRow = p_rowNo
    a_testSheet.Range(LabelColumn & p_durationRow).value = TestDurationLabel
    
    p_rowNo = p_rowNo + 1
    a_testSheet.Range(TestNameColumn & p_rowNo).value = TestNameLabel
    a_testSheet.Range(OutcomeColumn & p_rowNo).value = TestOutcomeLabel
    
    Dim p_testMethods As TestMethodFlags
    Dim p_macros As Collection: Set p_macros = WorkbookUtilities.EnumerateTestMacros(p_moduleInfo, p_testMethods)
    Dim p_macroInfo As MacroInfo
    
    Dim p_stopper As cc_isr_Core_IO.StopWatch: Set p_stopper = cc_isr_Core_IO.Factory.NewStopwatch
    
    Dim p_macroFullName As String
    
    Dim p_moduleMacroInfo As MacroInfo: Set p_moduleMacroInfo = cc_isr_Test_Fx.Factory.NewMacroInfo()
    
    If TestMethodFlags.BeforeAll = (p_testMethods And TestMethodFlags.BeforeAll) Then
        p_moduleMacroInfo.InitializeModuleMacro p_moduleInfo, ConstantsEnums.BeforeAllMethodName
        p_macroFullName = p_moduleMacroInfo.FullName
        If a_usingRemoteWorkbook Then
            p_macroFullName = "'" & p_filePath & "'!" & p_macroFullName
        End If
        Application.Run (p_macroFullName)
    End If
    
    For Each p_macroInfo In p_macros
        
        VBA.DoEvents
        
        ' execute the Before Each module test method.
        
        If TestMethodFlags.BeforeEach = (p_testMethods And TestMethodFlags.BeforeEach) Then
            p_moduleMacroInfo.InitializeModuleMacro p_moduleInfo, ConstantsEnums.BeforeEachMethodName
            p_macroFullName = p_moduleMacroInfo.FullName
            If a_usingRemoteWorkbook Then
                p_macroFullName = "'" & p_filePath & "'!" & p_macroFullName
            End If
            Application.Run (p_macroFullName)
        End If
        
        p_macroFullName = p_macroInfo.FullName
        If a_usingRemoteWorkbook Then
            p_macroFullName = "'" & p_filePath & "'!" & p_macroFullName
        End If
        
        Set Assert = Application.Run(p_macroFullName)
        
        p_rowNo = p_rowNo + 1
        
        If IsObject(Assert) Then
            a_testSheet.Range(TestNameColumn & p_rowNo).value = p_macroInfo.MacroName
            If Assert.AssertInconclusive Then
                p_inconclusiveCount = p_inconclusiveCount + 1
                a_testSheet.Range(OutcomeColumn & p_rowNo).value = InconclusiveOutcomeLabel
                a_testSheet.Range(OutcomeColumn & p_rowNo).Interior.ColorIndex = InconclusiveColorIndex
            ElseIf Assert.AssertSuccessful Then
                p_passedCount = p_passedCount + 1
                a_testSheet.Range(OutcomeColumn & p_rowNo).value = PassOutcomeLabel
                a_testSheet.Range(OutcomeColumn & p_rowNo).Interior.ColorIndex = PassColorIndex
            Else
                p_failedCount = p_failedCount + 1
                a_testSheet.Range(OutcomeColumn & p_rowNo).value = Assert.AssertMessage
                a_testSheet.Range(OutcomeColumn & p_rowNo).Interior.ColorIndex = FailColorIndex
            End If
        Else
            a_testSheet.Range(OutcomeColumn & p_rowNo).value = "assert not defined"
        End If
        
        ' execute the After Each module test method.
        
        If TestMethodFlags.AfterEach = (p_testMethods And TestMethodFlags.AfterEach) Then
            p_moduleMacroInfo.InitializeModuleMacro p_moduleInfo, ConstantsEnums.AfterEachMethodName
            p_macroFullName = p_moduleMacroInfo.FullName
            If a_usingRemoteWorkbook Then
                p_macroFullName = "'" & p_filePath & "'!" & p_macroFullName
            End If
            Application.Run (p_macroFullName)
        End If
        
    Next p_macroInfo
    
    ' execute the After All module test method.
    
    If TestMethodFlags.AfterAll = (p_testMethods And TestMethodFlags.AfterAll) Then
        p_moduleMacroInfo.InitializeModuleMacro p_moduleInfo, ConstantsEnums.AfterAllMethodName
        p_macroFullName = p_moduleMacroInfo.FullName
        If a_usingRemoteWorkbook Then
            p_macroFullName = "'" & p_filePath & "'!" & p_macroFullName
        End If
        Application.Run (p_macroFullName)
    End If
    
    
    p_stopper.StopCounter
    
    a_testSheet.Range(ItemColumn & p_inconclusiveCountRow).value = p_inconclusiveCount
    a_testSheet.Range(ItemColumn & p_passCountRow).value = p_passedCount
    a_testSheet.Range(ItemColumn & p_failCountRow).value = p_failedCount
    a_testSheet.Range(ItemColumn & p_durationRow).value = CStr(p_stopper.ElapsedMilliseconds) & TestDurationUnits
    
    Exit Sub

Finally:
    
    p_rowNo = p_rowNo + 1
    a_testSheet.Range(TestNameColumn & p_rowNo).value = "Test execution failed"
    a_testSheet.Range(OutcomeColumn & p_rowNo).value = cc_isr_Core_IO.ErrorMessageBuilder.FormattedErr()
    a_testSheet.Range(OutcomeColumn & p_rowNo).Interior.ColorIndex = FailColorIndex
    p_rowNo = p_rowNo + 1
    If 0 <> Err.LastDllError Then
        a_testSheet.Range(TestNameColumn & p_rowNo).value = "Last DLL Error"
        a_testSheet.Range(OutcomeColumn & p_rowNo).value = "0x" & VBA.LCase$(VBA.Hex$(Err.LastDllError))
    End If

End Sub


