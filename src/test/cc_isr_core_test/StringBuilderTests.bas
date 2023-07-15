Attribute VB_Name = "StringBuilderTests"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' StringBuilderTests.bas
'
' Dependencies:
'
' Assert.cls
' StringBuilder.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Tests appending items to string builder. </summary>
''' <returns>   An instance of the <see cref="Assert"/>   class. </returns>
Public Function TestAppendingToEmptyBuilder() As Assert
    
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Constructor.CreateStringBuilder
    Dim p_expected As String
    p_expected = "a"
    p_builder.Append p_expected
    Set TestAppendingToEmptyBuilder = Assert.AreEqual(p_expected, p_builder.ToString, _
            "Appended value should equal expected value")

End Function

''' <summary>   Unit test. Tests appending an empty string to the string builder. </summary>
''' <returns>   An instance of the <see cref="Assert"/>   class. </returns>
Public Function TestAppendingEmptyString() As Assert
    
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Constructor.CreateStringBuilder
    Dim p_expected As String
    p_expected = vbNullString
    p_builder.Append p_expected
    Set TestAppendingEmptyString = Assert.AreEqual(p_expected, p_builder.ToString, _
            "Appended empty value should equal p_expected value")

End Function

''' <summary>   Unit test. Tests appending a long string to the string builder. </summary>
''' <returns>   An instance of the <see cref="Assert"/>   class. </returns>
Public Function TestAppendingLongString() As Assert
    
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Constructor.CreateStringBuilder
    Dim p_expected As String
    p_expected = StringExtensions.Repeat("a", 1000)
    p_builder.Append p_expected
    Set TestAppendingLongString = Assert.AreEqual(p_expected, p_builder.ToString, _
            "Appended a long value should equal p_expected value")

End Function

Public Function TestAppendingLineFeed() As Assert
    
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Constructor.CreateStringBuilder
    Dim p_expected As String
    p_expected = "a" & vbLf
    p_builder.Append p_expected
    Set TestAppendingLineFeed = Assert.AreEqual(p_expected, p_builder.ToString, _
            "Appended value with line feed should equal expected value")

End Function



