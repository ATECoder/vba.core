VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTestSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const m_lastErrorMessageCell As String = "C4"

''' <summary>   Gets the Last Error Message. </summary>
Friend Property Get LastErrorMessage() As String
    LastErrorMessage = Range(m_lastErrorMessageCell).value
End Property

''' <summary>   Sets the Last Error Message. </summary>
Friend Property Let LastErrorMessage(ByVal a_value As String)
    Range(m_lastErrorMessageCell).value = a_value
End Property

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


