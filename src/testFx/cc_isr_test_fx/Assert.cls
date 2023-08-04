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
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals True if values
'''             are equal; otherwise, False. </returns>
Public Function AreEqual(ByVal a_expected As Variant, ByVal a_actual As Variant, ByVal a_message As String) As Assert

    Set AreEqual = Me.IsTrue(Not a_expected <> a_actual, a_message)
    
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
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals True if
'''             values are not equal; otherwise, false. </returns>
Public Function AreNotEqual(ByVal a_expected As Variant, ByVal a_actual As Variant, ByVal a_message As String) As Assert
    
    Set AreNotEqual = Me.IsTrue(a_expected <> a_actual, a_message)
    
    If Not AreNotEqual.AssertSuccessful Then
        AreNotEqual.AssertMessage = AreNotEqual.AssertMessage & vbCrLf & _
                "    Expected: " & a_expected & vbCrLf & _
                "    Actual:   " & a_actual
    End If
End Function

''' <summary>   Asserts that the given value is Nothing. </summary>
''' <param name="a_actual">    [Variant] The actual value. </param
''' <param name="a_message">   [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals True if
'''             <paramref name="a_actual"/> is Nothing; otherwise, false. </returns>
Public Function IsNothing(ByVal a_actual As Variant, ByVal a_message As String) As Assert
    Set IsNothing = Me.IsTrue(a_actual Is Nothing, a_message)
End Function

''' <summary>   Asserts that the given value is not Nothing. </summary>
''' <param name="a_actual">    [Variant] The actual value. </param
''' <param name="a_message">   [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals True if
'''             <paramref name="a_actual"/> is not Nothing; otherwise, false. </returns>
Public Function IsNotNothing(ByVal a_actual As Variant, ByVal a_message As String) As Assert
    
    Set IsNotNothing = Me.IsFalse(a_actual Is Nothing, a_message)

End Function

''' <summary>   Asserts that the given value is null. </summary>
''' <param name="a_actual">    [Variant] The actual value. </param
''' <param name="a_message">   [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals True if
'''             <paramref name="a_actual"/> is null; otherwise, false. </returns>
Public Function IsNull(ByVal a_actual As Variant, ByVal a_message As String) As Assert
    Set IsNull = Me.IsFalse(VBA.IsObject(a_actual), a_message)
End Function

''' <summary>   Asserts that the given value is not null. </summary>
''' <param name="a_actual">    [Variant] The actual value. </param
''' <param name="a_message">   [String] A message to show in case the assertion failed. </param
''' <returns> True if <paramref name="a_actual"/> is not null; otherwise, false. </returns>
Public Function IsNotNull(ByVal a_actual As Variant, ByVal a_message As String) As Assert
    
    Set IsNotNull = Me.IsTrue(VBA.IsObject(a_actual), a_message)

End Function

''' <summary>   Asserts that the given value is true. </summary>
''' <param name="a_truthy">    [Variant] A value that can be seen as truthy. Very likely should be boolean,
'''                             but maybe you like to live dangerously. </param
''' <param name="a_message">   [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals True if
'''             <paramref name="truthy"/> is true; otherwise, false. </returns>
Public Function IsTrue(ByVal a_truthy As Variant, ByVal a_message As String) As Assert
    
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
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals False if
'''             <paramref name="falsy"/> is False; otherwise, True. </returns>
Public Function IsFalse(ByVal a_falsy As Variant, ByVal a_message As String) As Assert
    
    Set IsFalse = New Assert
    IsFalse.AssertSuccessful = True
    IsFalse.AssertMessage = a_message

    If a_falsy Then
        IsFalse.AssertSuccessful = False
        IsFalse.AssertMessage = "Assert failed: " & a_message
    End If

End Function

''' <summary>   Asserts a failure with the specified message. </summary>
''' <param name="a_message">   [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals False.. </returns>
Public Function Fail(ByVal a_message As String) As Assert
    
    Set Fail = New Assert
    Fail.AssertSuccessful = False
    Fail.AssertMessage = a_message

End Function

''' <summary>   Asserts a failure with the specified message. </summary>
''' <param name="a_message">   [String, vbNullString] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals False.. </returns>
Public Function Pass(Optional ByVal a_message As String = VBA.vbNullString) As Assert
    
    Set Pass = New Assert
    Pass.AssertSuccessful = True
    Pass.AssertMessage = a_message

End Function

