VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GuardClauses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Encapsulates throwing methods upon meeting some error conditions. </summary>
''' <remarks>   Dependencies: UserDefinedErrors.cls. </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Raises a <see cref="UserDefinedError"/> error if the specified Boolean
'''             expression is True. </summary>
''' <param name="a_throw">     [Boolean] True if the error is to be thrown. </param>
''' <param name="a_source">    [String] The source of the error. </param>
''' <param name="a_message">   [String] The message to append to the standard error description. </param>
''' <param name="a_error">     [<see cref="UserDefinedError"/>] to raise. </param>
Public Sub GuardExpression(ByVal a_throw As Boolean, ByVal a_source As String, _
        ByVal a_message As String, _
        ByVal a_error As UserDefinedError)
    
    If a_throw Then UserDefinedErrors.RaiseError a_error, a_source, a_message

End Sub

''' <summary>   Raises a <see cref="UserDefinedErrors.InvalidOperationError"/> exceptuion
'''             if the specified instance isn't the default/predeclared class. </summary>
''' <remarks>   Use this guard clause in cases where a predeclared
'''             class is used both as predeclare or instantiated to prevent calling a
'''             method from the instantiated object. </remarks>
''' <param name="a_instance">           [Object] The instance. </param>
''' <param name="a_defaultInstance">    [Object] The default /predeclared instance. </param>
''' <param name="a_source">             [String] The source of the error. </param>
''' <param name="a_message">            [String, 'Methods should be ...'] The message to append to the standard
'''                                     error description. </param>
Public Sub GuardNonDefaultInstance(ByVal a_instance As Object, _
        ByVal a_defaultInstance As Object, _
        ByVal a_source As String, _
        Optional ByVal a_message As String = "Method should be invoked from the default/predeclared instance of this class.")
    
    Debug.Assert TypeName(a_instance) = TypeName(a_defaultInstance)
    
    Me.GuardExpression Not a_instance Is a_defaultInstance, _
                       a_source, _
                       a_message, UserDefinedErrors.InvalidOperationError

End Sub

''' <summary>   Raises a run-time error if the specified instance is the default instance. </summary>
''' <remarks>   Use this guard clause in cases where a predeclared
'''             class is used both as predeclare or instantiated to prevent calling a
'''             method from the predeclared object. </remarks>
''' <param name="a_instance">           [Object] The instance. </param>
''' <param name="a_defaultInstance">    [Object] The default /predeclared instance. </param>
''' <param name="a_source">             [String] The source of the error. </param>
''' <param name="a_message">            [String, 'Methods should be ...'] The message to append to the standard
'''                                     error description. </param>
Public Sub GuardDefaultInstance(ByVal a_instance As Object, _
        ByVal a_defaultInstance As Object, _
        ByVal a_source As String, _
        Optional ByVal a_message As String = "Method should be invoked from a new instance of this class.")

    Debug.Assert TypeName(a_instance) = TypeName(a_defaultInstance)
    
    GuardExpression a_instance Is a_defaultInstance, a_source, a_message, _
                    UserDefinedErrors.InvalidOperationError

End Sub

''' <summary>   Raises a run-time error if the specified object reference is already set. </summary>
''' <param name="a_value">     [Variant] The object reference that might already exist. </param>
''' <param name="a_source">    [String] The source of the error. </param>
''' <param name="a_message">   [String, 'Value is already initialized.'] The message to append to the standard
'''                            error description. </param>
Public Sub GuardDoubleInitialization(ByVal a_value As Variant, _
        ByVal a_source As String, _
        Optional ByVal a_message As String = " Value is already initialized.")
    
    Dim p_throw As Boolean
    If IsObject(a_value) Then
        p_throw = Not a_value Is Nothing
    Else
        p_throw = a_value <> cc_isr_Core_IO.CoreExtensions.GetDefaultValue(VBA.VarType(a_value))
    End If
    GuardExpression p_throw, a_source, a_message, UserDefinedErrors.InvalidOperationError

End Sub

''' <summary>   Returns the default value to the specified variable type. </summary>
''' <summary>   Raises a run-time error if the specified object reference is Nothing. </summary>
''' <param name="a_instance">   [Object] The instance. </param>
''' <param name="a_source">     [String] The source of the error. </param>
''' <param name="a_message">    [String, 'Object reference cannot be Nothing.'] The message to append to the standard
'''                             error description. </param>
Public Sub GuardNullReference(ByVal a_instance As Object, _
        ByVal a_source As String, _
        Optional ByVal a_message As String = "Object reference cannot be Nothing.")
    
    Me.GuardExpression a_instance Is Nothing, a_source, a_message, _
        cc_isr_Core_IO.UserDefinedErrors.NullArgumentError

End Sub

''' <summary>   Raises a run-time error if the specified string is empty. </summary>
''' <param name="a_instance">           [Object] The instance. </param>
''' <param name="a_defaultInstance">    [Object] The default /predeclared instance. </param>
''' <param name="a_source">             [String] The source of the error. </param>
''' <param name="a_message">            [String, 'String cannot be empty.'] The message to append to the standard
'''                                     error description. </param>
Public Sub GuardEmptyString(ByVal a_value As String, _
        ByVal a_source As String, _
        Optional ByVal a_message As String = "String cannot be empty.")
    
    Me.GuardExpression a_value = VBA.vbNullString, a_source, a_message, _
        cc_isr_Core_IO.UserDefinedErrors.InvalidArgumentError
                    
End Sub


