VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Asserts and provides outcome of test conditions.  </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Indicates whether the assertion was successful. </summary>
''' <value>   [Boolean]. </value>
Public AssertSuccessful As Boolean

''' <summary>   The outcome message of this assertion. </summary>
''' <value>   A String. </value>
Public AssertMessage As String

''' <summary>   Asserts that values are equal. If not, this will return a failed assertion. </summary>
''' <param name="a_expected">   [Variant] The value the caller expects to have. </param
''' <param name="a_actual">     [Variant] The actual value. </param
''' <param name=a_"message">    [String] A message to show in case the assertion failed. </param
''' <returns>   An <see cref="Assert"/>   instant of True if values are equal; otherwise, False. </returns>
Public Function AreEqual(ByVal a_expected As Variant, ByVal a_actual As Variant, ByRef a_message As String) As Assert

    Set AreEqual = IsTrue(Not a_expected <> a_actual, a_message)
    
    If Not AreEqual.AssertSuccessful Then
        AreEqual.AssertMessage = AreEqual.AssertMessage & vbCrLf & _
                "    Expected: " & a_expected & vbCrLf & _
                "    Actual:   " & a_actual
    End If
End Function

''' <summary>   Asserts that values are not equal. If not, this will return a failed assertion. </summary>
''' <param name="a_expected">   [Variant] The value the caller expects to have. </param
''' <param name="a_actual">     [Variant] The actual value. </param
''' <param name="a_message">    [String] A message to show in case the assertion failed. </param
''' <returns>   True if values are not equal; otherwise, false. </returns>
Public Function AreNotEqual(ByVal a_expected As Variant, ByVal a_actual As Variant, ByRef a_message As String) As Assert
    
    Set AreNotEqual = IsTrue(a_expected <> a_actual, a_message)
    
    If Not AreNotEqual.AssertSuccessful Then
        AreNotEqual.AssertMessage = AreNotEqual.AssertMessage & vbCrLf & _
                "    Expected: " & a_expected & vbCrLf & _
                "    Actual:   " & a_actual
    End If
End Function

''' <summary>   Asserts that the given value is null. </summary>
''' <param name="a_actual">    [Variant] The actual value. </param
''' <param name="a_message">   [String] A message to show in case the assertion failed. </param
''' <returns>   True if <paramref name="a_actual"/>   is null; otherwise, false. </returns>
Public Function IsNull(ByVal a_actual As Variant, ByRef a_message As String) As Assert
    Set IsNull = IsFalse(IsObject(a_actual), a_message)
End Function

''' <summary>   Asserts that the given value is not null. </summary>
''' <param name="a_actual">    [Variant] The actual value. </param
''' <param name="a_message">   [String] A message to show in case the assertion failed. </param
''' <returns>   True if <paramref name="a_actual"/>   is not null; otherwise, false. </returns>
Public Function IsNotNull(ByVal a_actual As Variant, ByRef a_message As String) As Assert
    
    Set IsNotNull = IsTrue(IsObject(a_actual), a_message)

End Function

''' <summary>   Asserts that the given value is true. </summary>
''' <param name="a_truthy">    [Variant] A value that can be seen as truthy. Very likely should be boolean,
'''                             but maybe you like to live dangerously. </param
''' <param name="a_message">   [String] A message to show in case the assertion failed. </param
''' <returns>   True if <paramref name="truthy"/>   is true; otherwise, false. </returns>
Public Function IsTrue(ByVal a_truthy As Variant, ByRef a_message As String) As Assert
    
    Set IsTrue = New Assert
    IsTrue.AssertSuccessful = True
    IsTrue.AssertMessage = a_message

    If Not a_truthy Then
        IsTrue.AssertSuccessful = False
        IsTrue.AssertMessage = "Assert failed: " & a_message
    End If

End Function

''' <summary>   Asserts that the given value is False. </summary>
''' <param name="a_falsy">     [Variant] A value that can be seen as falsy. Very likely should be boolean,
'''                            but maybe you like to live dangerously. </param
''' <param name="a_message">   [String] A message to show in case the assertion failed. </param
''' <returns>   False if <paramref name="falsy"/>   is False; otherwise, false. </returns>
Public Function IsFalse(ByVal a_falsy As Variant, ByRef a_message As String) As Assert
    
    Set IsFalse = New Assert
    IsFalse.AssertSuccessful = True
    IsFalse.AssertMessage = a_message

    If a_falsy Then
        IsFalse.AssertSuccessful = False
        IsFalse.AssertMessage = "Assert failed: " & a_message
    End If

End Function


