VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Testing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

''' <summary>   Executes the <see cref="WorkbookUtilitiesTests"/>. </summary>
Public Sub RunTests()

    TestExecutive.Execute Me
    
End Sub

Public Sub TestCallByName()

    ' call by name does not work on modules.
    
    ' application run does not work on classes.
    
    ' this works
    ' VBA.CallByName UtilitiesTests, "TestModuleList", VbMethod
    
    ' this works:
    ' VBA.CallByName Testing, "RunTests", VbMethod
    
End Sub


