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

Private Type this_
    AssertSuccessful As Boolean
    AssertInconclusive As Boolean
    AssertMessage As String
End Type

Private This As this_

''' <summary>   Gets or set the indication of whether the assertion was successful. </summary>
''' <value>   [Boolean]. </value>
Public Property Get AssertSuccessful() As Boolean
    AssertSuccessful = This.AssertSuccessful
End Property

Public Property Let AssertSuccessful(ByVal a_value As Boolean)
    This.AssertSuccessful = a_value
End Property

''' <summary>   Gets or sets the indication of whether the assertion was inconclusive. </summary>
''' <remarks>   An assertion might be inconclusive if, for example, a device test cannot
'''             run because the device does not exist on the system as could be
'''             determined by a failure to ping the device. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get AssertInconclusive() As Boolean
    AssertInconclusive = This.AssertInconclusive
End Property

Public Property Let AssertInconclusive(ByVal a_value As Boolean)
    This.AssertInconclusive = a_value
End Property

''' <summary>   Gets or sets the outcome message of this assertion. </summary>
''' <value>   [String]. </value>
Public Property Get AssertMessage() As String
    AssertMessage = This.AssertMessage
End Property

Public Property Let AssertMessage(ByVal a_value As String)
    This.AssertMessage = a_value
End Property

''' <summary>   Selects an outcome. </summary>
''' <param name="a_passed">         [Variant] The value or object representing the passed outcome. </param
''' <param name="a_failed">         [Variant] The value or object representing the failed outcome. </param
''' <param name="a_inconclusive">   [Variant] The value or object representing the inconclusive outcome. </param
''' <returns>   [Variant]. </returns>
Public Function Choose(ByVal a_passed As Variant, ByVal a_failed As Variant, ByVal a_inconclusive As Variant) As Variant

    If Me.AssertInconclusive Then
        If IsObject(a_inconclusive) Then
            Set Choose = a_inconclusive
        Else
            Choose = a_inconclusive
        End If
    ElseIf Me.AssertSuccessful Then
        If IsObject(a_passed) Then
            Set Choose = a_passed
        Else
            Choose = a_passed
        End If
    Else
        If IsObject(a_failed) Then
            Set Choose = a_failed
        Else
            Choose = a_failed
        End If
    End If

End Function

''' <summary>   Builds a report consisting of the supplied test name, outcome, and message. </summary>
''' <param name="a_testName">   [String] The name of the unit test to include in the report. </param
''' <param name="a_passed">         [String, ' passed.'] The passed part of the report. </param
''' <param name="a_failed">         [String] ' failed.'] The failed part of the outcome. </param
''' <param name="a_inconclusive">   [String] ' inconclusive.'] The inconclusive part of the report. </param
''' <returns>   [String]. </returns>
Public Function BuildReport(ByVal a_testName As String, _
        Optional ByVal a_passed As String = " passed.", _
        Optional ByVal a_failed As String = " failed.", _
        Optional ByVal a_inconclusive As String = " inconclusive.") As String

    BuildReport = a_testName & Me.Choose(a_passed, a_failed, a_inconclusive) & _
                    IIf(Me.AssertSuccessful, VBA.vbNullString, " " & Me.AssertMessage)
    
End Function

''' <summary>   Asserts that values are equal within the provided epsilon.
''' If not, this will return a failed assertion. </summary>
''' <param name="a_expected">   [Double] The value the caller expects to have. </param
''' <param name="a_actual">     [Double] The actual value. </param
''' <param name="a_epsilon">    [Double] The maximum difference for closeness. </param
''' <param name="a_message">    [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals True if values
'''             are equal; otherwise, False. </returns>
Public Function AreCloseDouble(ByVal a_expected As Double, ByVal a_actual As Double, _
    ByVal a_epsilon As Double, ByVal a_message As String) As Assert

    Set AreCloseDouble = Me.IsTrue(VBA.Abs(a_expected - a_actual) <= a_epsilon, a_message)
    
    If Not AreCloseDouble.AssertSuccessful Then
        AreCloseDouble.AssertMessage = AreCloseDouble.AssertMessage & vbCrLf & _
                "    Expected: " & VBA.CStr(a_expected) & vbCrLf & _
                "    Actual:   " & VBA.CStr(a_actual) & vbCrLf & _
                "    Epsilon:  " & VBA.CStr(a_epsilon)
    End If
End Function

''' <summary>   Asserts that values are not equal within the provided epsilon.
''' If not, this will return a failed assertion. </summary>
''' <param name="a_expected">   [Double] The value the caller expects to have. </param
''' <param name="a_actual">     [Double] The actual value. </param
''' <param name="a_epsilon">    [Double] The maximum difference for closeness. </param
''' <param name="a_message">    [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals True if
'''             values are not equal; otherwise, false. </returns>
Public Function AreNotCloseDouble(ByVal a_expected As Double, ByVal a_actual As Double, _
    ByVal a_epsilon As Double, ByVal a_message As String) As Assert
    
    Set AreNotCloseDouble = Me.IsTrue(VBA.Abs(a_expected - a_actual) > a_epsilon, a_message)
    
    If Not AreNotCloseDouble.AssertSuccessful Then
        AreNotCloseDouble.AssertMessage = AreNotCloseDouble.AssertMessage & vbCrLf & _
                "    Expected: " & VBA.CStr(a_expected) & vbCrLf & _
                "    Actual:   " & VBA.CStr(a_actual) & vbCrLf & _
                "    Epsilon:  " & VBA.CStr(a_epsilon)
    End If

End Function

''' <summary>   Asserts that values are equal within the provided epsilon.
''' If not, this will return a failed assertion. </summary>
''' <param name="a_expected">   [Single] The value the caller expects to have. </param
''' <param name="a_actual">     [Single] The actual value. </param
''' <param name="a_epsilon">    [Single] The maximum difference for closeness. </param
''' <param name="a_message">    [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals True if values
'''             are equal; otherwise, False. </returns>
Public Function AreCloseSingle(ByVal a_expected As Single, ByVal a_actual As Single, _
    ByVal a_epsilon As Single, ByVal a_message As String) As Assert

    Set AreCloseSingle = Me.IsTrue(VBA.Abs(a_expected - a_actual) <= a_epsilon, a_message)
    
    If Not AreCloseSingle.AssertSuccessful Then
        AreCloseSingle.AssertMessage = AreCloseSingle.AssertMessage & vbCrLf & _
                "    Expected: " & VBA.CStr(a_expected) & vbCrLf & _
                "    Actual:   " & VBA.CStr(a_actual) & vbCrLf & _
                "    Epsilon:  " & VBA.CStr(a_epsilon)
    End If
End Function

''' <summary>   Asserts that values are not equal within the provided epsilon.
''' If not, this will return a failed assertion. </summary>
''' <param name="a_expected">   [Single] The value the caller expects to have. </param
''' <param name="a_actual">     [Single] The actual value. </param
''' <param name="a_epsilon">    [Single] The maximum difference for closeness. </param
''' <param name="a_message">    [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals True if
'''             values are not equal; otherwise, false. </returns>
Public Function AreNotCloseSingle(ByVal a_expected As Single, ByVal a_actual As Single, _
    ByVal a_epsilon As Single, ByVal a_message As String) As Assert
    
    Set AreNotCloseSingle = Me.IsTrue(VBA.Abs(a_expected - a_actual) > a_epsilon, a_message)
    
    If Not AreNotCloseSingle.AssertSuccessful Then
        AreNotCloseSingle.AssertMessage = AreNotCloseSingle.AssertMessage & vbCrLf & _
                "    Expected: " & VBA.CStr(a_expected) & vbCrLf & _
                "    Actual:   " & VBA.CStr(a_actual) & vbCrLf & _
                "    Epsilon:  " & VBA.CStr(a_epsilon)
    End If

End Function


''' <summary>   Asserts that values are equal. If not, this will return a failed assertion. </summary>
''' <param name="a_expected">   [Variant] The value the caller expects to have. </param
''' <param name="a_actual">     [Variant] The actual value. </param
''' <param name="a_message">    [String] A message to show in case the assertion failed. </param
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

''' <summary>   Asserts that string values are equal. If not, this will return a failed assertion. </summary>
''' <param name="a_expected">   [String] The value the caller expects to have. </param
''' <param name="a_actual">     [String] The actual value. </param
''' <param name="a_compare">    [VBA.VbCompareMethod] The compare method. </param
''' <param name="a_message">    [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals True if values
'''             are equal; otherwise, False. </returns>
Public Function AreEqualString(ByVal a_expected As String, ByVal a_actual As String, _
    ByVal a_compare As VBA.VbCompareMethod, _
    ByVal a_message As String) As Assert

    Set AreEqualString = Me.IsTrue(1 = VBA.InStr(1, a_expected, a_actual, a_compare), a_message)
    
    If Not AreEqualString.AssertSuccessful Then
        AreEqualString.AssertMessage = AreEqualString.AssertMessage & vbCrLf & _
                "    Expected: " & a_expected & vbCrLf & _
                "    Actual:   " & a_actual
    End If
End Function

''' <summary>   Asserts that string values are not equal. If not, this will return a failed assertion. </summary>
''' <param name="a_expected">   [Variant] The value the caller expects to have. </param
''' <param name="a_actual">     [Variant] The actual value. </param
''' <param name="a_compare">    [VBA.VbCompareMethod] The compare method. </param
''' <param name="a_message">    [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals True if
'''             values are not equal; otherwise, false. </returns>
Public Function AreNotEqualString(ByVal a_expected As String, ByVal a_actual As String, _
    ByVal a_compare As VBA.VbCompareMethod, _
    ByVal a_message As String) As Assert
    
    Set AreNotEqualString = Me.IsTrue(1 <> VBA.InStr(1, a_expected, a_actual, a_compare), a_message)
    
    If Not AreNotEqualString.AssertSuccessful Then
        AreNotEqualString.AssertMessage = AreNotEqualString.AssertMessage & vbCrLf & _
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
    IsTrue.AssertSuccessful = a_truthy
    IsTrue.AssertMessage = a_message

End Function

''' <summary>   Asserts that the given value is False. </summary>
''' <param name="a_falsy">     [Variant] A value that can be seen as falsy. Very likely should be boolean,
'''                            but maybe you like to live dangerously. </param
''' <param name="a_message">   [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals False if
'''             <paramref name="falsy"/> is False; otherwise, True. </returns>
Public Function IsFalse(ByVal a_falsy As Variant, ByVal a_message As String) As Assert
    
    Set IsFalse = New Assert
    IsFalse.AssertSuccessful = Not a_falsy
    IsFalse.AssertMessage = a_message

End Function

''' <summary>   Asserts a failure with the specified message. </summary>
''' <param name="a_message">   [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals False. </returns>
Public Function Fail(ByVal a_message As String) As Assert
    
    Set Fail = New Assert
    Fail.AssertSuccessful = False
    Fail.AssertMessage = a_message

End Function

''' <summary>   Asserts a failure with the specified message. </summary>
''' <param name="a_message">   [String, vbNullString] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> equals False. </returns>
Public Function Pass(Optional ByVal a_message As String = VBA.vbNullString) As Assert
    
    Set Pass = New Assert
    Pass.AssertSuccessful = True
    Pass.AssertMessage = a_message

End Function

''' <summary>   Asserts that object are Same, in the sense that the actual object 'Is' the
'''             expected object. If not, this will return a failed assertion. </summary>
''' <param name="a_expected">   [Variant] The expected object. </param
''' <param name="a_actual">     [Variant] The actual object value. </param
''' <param name="a_message">    [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> Sames True if values
'''             are Same; otherwise, False. </returns>
Public Function AreSame(ByVal a_expected As Object, ByVal a_actual As Object, ByVal a_message As String) As Assert

    Set AreSame = Me.IsTrue(a_actual Is a_expected, a_message)
    
    If Not AreSame.AssertSuccessful Then
        AreSame.AssertMessage = AreSame.AssertMessage & vbCrLf & _
                "    Expected: " & TypeName(a_expected) & vbCrLf & _
                "    Actual:   " & TypeName(a_actual)
    End If
End Function

''' <summary>   Asserts that object are Not the Same, in the sense that the actual object
'''             'Is' not the expected object. If not, this will return a failed assertion. </summary>
''' <param name="a_expected">   [Variant] The value the caller expects to have. </param
''' <param name="a_actual">     [Variant] The actual value. </param
''' <param name="a_message">    [String] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertSuccessful"/> Sames True if
'''             values are not Same; otherwise, false. </returns>
Public Function AreNotSame(ByVal a_expected As Object, ByVal a_actual As Object, ByVal a_message As String) As Assert
    
    Set AreNotSame = Me.IsFalse(a_actual Is a_expected, a_message)
    
    If Not AreNotSame.AssertSuccessful Then
        AreNotSame.AssertMessage = AreNotSame.AssertMessage & vbCrLf & _
                "    Expected: " & TypeName(a_expected) & vbCrLf & _
                "    Actual:   " & TypeName(a_actual)
    End If
End Function

''' <summary>   Asserts an <see cref="AssertInconclusive"/> outcome with the specified message. </summary>
''' <param name="a_message">   [String, vbNullString] A message to show in case the assertion failed. </param
''' <returns>   [<see cref="Assert">] with <see cref="Assert.AssertInconclusive"/> equals True. </returns>
Public Function Inconclusive(Optional ByVal a_message As String = VBA.vbNullString) As Assert
    
    Set Inconclusive = New Assert
    Inconclusive.AssertInconclusive = True
    Inconclusive.AssertMessage = a_message

End Function



