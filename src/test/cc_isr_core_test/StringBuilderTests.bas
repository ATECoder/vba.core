Attribute VB_Name = "StringBuilderTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. string builder methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Name As String
    TestNumber As Integer
    BeforeAllAssert As cc_isr_Test_Fx.Assert
    BeforeEachAssert As cc_isr_Test_Fx.Assert
    ErrTracer As IErrTracer
    TestCount As Integer
    RunCount As Integer
    PassedCount As Integer
    FailedCount As Integer
    InconclusiveCount As Integer
End Type

Private This As this_

''' <summary>   Runs the specified test. </summary>
Public Function RunTest(ByVal a_testNumber As Integer) As cc_isr_Test_Fx.Assert
    Dim p_outcome As cc_isr_Test_Fx.Assert
    'BeforeEach
    Select Case a_testNumber
        Case 1
            Set p_outcome = TestAppendingToEmptyBuilder
        Case 2
            Set p_outcome = TestAppendingEmptyString
        Case 3
            Set p_outcome = TestAppendingLongString
        Case 4
            Set p_outcome = TestAppendingLineFeed
        Case 5
            Set p_outcome = TestAppendFormat
        Case Else
    End Select
    Set RunTest = p_outcome
    'AfterEach
End Function

''' <summary>   Runs a single test. </summary>
Public Sub RunOneTest()
    'BeforeAll
    RunTest 1
    'AfterAll
End Sub

''' <summary>   Runs all tests. </summary>
Public Sub RunAllTests()
    This.Name = "BinaryExtensionTests"
    'BeforeAll
    Dim p_outcome As cc_isr_Test_Fx.Assert
    This.RunCount = 0
    This.PassedCount = 0
    This.FailedCount = 0
    This.InconclusiveCount = 0
    This.TestCount = 5
    Dim p_testNumber As Integer
    For p_testNumber = 1 To This.TestCount
        Set p_outcome = RunTest(p_testNumber)
        If Not p_outcome Is Nothing Then
            This.RunCount = This.RunCount + 1
            If p_outcome.AssertInconclusive Then
                This.InconclusiveCount = This.InconclusiveCount + 1
            ElseIf p_outcome.AssertSuccessful Then
                This.PassedCount = This.PassedCount + 1
            Else
                This.FailedCount = This.FailedCount + 1
            End If
        End If
        DoEvents
    Next p_testNumber
    'AfterAll
    Debug.Print "Ran " & VBA.CStr(This.RunCount) & " out of " & VBA.CStr(This.TestCount) & " tests."
    Debug.Print "Passed: " & VBA.CStr(This.PassedCount) & "; Failed: " & VBA.CStr(This.FailedCount) & _
                "; Inconclusive: " & VBA.CStr(This.InconclusiveCount) & "."
End Sub


''' <summary>   Unit test. Tests appending items to string builder. </summary>
''' <returns>   An instance of the <see cref="cc_isr_Test_Fx.Assert"/>   class. </returns>
Public Function TestAppendingToEmptyBuilder() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder
    Dim p_expected As String
    p_expected = "a"
    p_builder.Append p_expected
    
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expected, p_builder.ToString, _
            "Appended value should equal expected value")

    Debug.Print p_outcome.BuildReport("TestAppendingToEmptyBuilder")
    
    Set TestAppendingToEmptyBuilder = p_outcome

End Function

''' <summary>   Unit test. Tests appending an empty string to the string builder. </summary>
''' <returns>   An instance of the <see cref="cc_isr_Test_Fx.Assert"/>   class. </returns>
Public Function TestAppendingEmptyString() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder
    Dim p_expected As String
    p_expected = vbNullString
    p_builder.Append p_expected
    
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expected, p_builder.ToString, _
            "Appended empty value should equal p_expected value")

    Debug.Print p_outcome.BuildReport("TestAppendingEmptyString")
    
    Set TestAppendingEmptyString = p_outcome

End Function

''' <summary>   Unit test. Tests appending a long string to the string builder. </summary>
''' <returns>   An instance of the <see cref="cc_isr_Test_Fx.Assert"/>   class. </returns>
Public Function TestAppendingLongString() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder
    Dim p_expected As String
    p_expected = StringExtensions.Repeat("a", 1000)
    p_builder.Append p_expected
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expected, p_builder.ToString, _
            "Appended a long value should equal p_expected value")

    Debug.Print p_outcome.BuildReport("TestAppendingLongString")
    
    Set TestAppendingLongString = p_outcome

End Function

''' <summary>   Unit test. Tests appending a line feed to the string builder. </summary>
''' <returns>   An instance of the <see cref="cc_isr_Test_Fx.Assert"/>   class. </returns>
Public Function TestAppendingLineFeed() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder
    Dim p_expected As String
    p_expected = "a" & vbLf
    p_builder.Append p_expected
    
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expected, p_builder.ToString, _
            "Appended value with line feed should equal expected value")

    Debug.Print p_outcome.BuildReport("TestAppendingLineFeed")
    
    Set TestAppendingLineFeed = p_outcome

End Function

''' <summary>   Unit test. Tests appending a formatted stringto the string builder. </summary>
''' <returns>   An instance of the <see cref="cc_isr_Test_Fx.Assert"/>   class. </returns>
Public Function TestAppendFormat() As cc_isr_Test_Fx.Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder
    Dim p_expected As String
    p_expected = "a+b+c"
    Dim p_format As String: p_format = "{0}+{1}+{2}"
    p_builder.Appendformat p_format, "a", "b", "c"
    
    Set p_outcome = cc_isr_Test_Fx.Assert.areEqual(p_expected, p_builder.ToString, _
            "Appended value with line feed should equal expected value")

    Debug.Print p_outcome.BuildReport("TestAppendFormat")
    
    Set TestAppendFormat = p_outcome

End Function


