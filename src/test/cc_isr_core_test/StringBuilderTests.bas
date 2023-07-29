Attribute VB_Name = "StringBuilderTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. string builder methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Tests appending items to string builder. </summary>
''' <returns>   An instance of the <see cref="cc_isr_Test_Fx.Assert"/>   class. </returns>
Public Function TestAppendingToEmptyBuilder() As cc_isr_Test_Fx.Assert
    
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder
    Dim p_expected As String
    p_expected = "a"
    p_builder.Append p_expected
    
    Set TestAppendingToEmptyBuilder = cc_isr_Test_Fx.Assert.AreEqual(p_expected, p_builder.ToString, _
            "Appended value should equal expected value")

End Function

''' <summary>   Unit test. Tests appending an empty string to the string builder. </summary>
''' <returns>   An instance of the <see cref="cc_isr_Test_Fx.Assert"/>   class. </returns>
Public Function TestAppendingEmptyString() As cc_isr_Test_Fx.Assert
    
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder
    Dim p_expected As String
    p_expected = vbNullString
    p_builder.Append p_expected
    
    Set TestAppendingEmptyString = cc_isr_Test_Fx.Assert.AreEqual(p_expected, p_builder.ToString, _
            "Appended empty value should equal p_expected value")

End Function

''' <summary>   Unit test. Tests appending a long string to the string builder. </summary>
''' <returns>   An instance of the <see cref="cc_isr_Test_Fx.Assert"/>   class. </returns>
Public Function TestAppendingLongString() As cc_isr_Test_Fx.Assert
    
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder
    Dim p_expected As String
    p_expected = StringExtensions.Repeat("a", 1000)
    p_builder.Append p_expected
    Set TestAppendingLongString = cc_isr_Test_Fx.Assert.AreEqual(p_expected, p_builder.ToString, _
            "Appended a long value should equal p_expected value")

End Function

''' <summary>   Unit test. Tests appending a line feed to the string builder. </summary>
''' <returns>   An instance of the <see cref="cc_isr_Test_Fx.Assert"/>   class. </returns>
Public Function TestAppendingLineFeed() As cc_isr_Test_Fx.Assert
    
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder
    Dim p_expected As String
    p_expected = "a" & vbLf
    p_builder.Append p_expected
    
    Set TestAppendingLineFeed = cc_isr_Test_Fx.Assert.AreEqual(p_expected, p_builder.ToString, _
            "Appended value with line feed should equal expected value")

End Function

''' <summary>   Unit test. Tests appending a formatted stringto the string builder. </summary>
''' <returns>   An instance of the <see cref="cc_isr_Test_Fx.Assert"/>   class. </returns>
Public Function TestAppendFormat() As cc_isr_Test_Fx.Assert
    
    Dim p_builder As StringBuilder
    Set p_builder = cc_isr_Core.Factory.NewStringBuilder
    Dim p_expected As String
    p_expected = "a+b+c"
    Dim p_format As String: p_format = "{0}+{1}+{2}"
    p_builder.Appendformat p_format, "a", "b", "c"
    
    Set TestAppendFormat = cc_isr_Test_Fx.Assert.AreEqual(p_expected, p_builder.ToString, _
            "Appended value with line feed should equal expected value")

End Function


