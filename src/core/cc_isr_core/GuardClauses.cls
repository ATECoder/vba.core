VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GuardClauses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' GuardClauses.cls
'
' Description: Encalsulate throw if error condition methods.
'
' Dependencies:
'
' UserDefiendErrors.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Raises a run-time error if the specified Boolean expression is True. </summary>
''' <param name="a_throw">     [Boolean] True if the error is to be thrown. </param>
''' <param name="a_source">    [String] The source of the error. </param>
''' <param name="a_message">   [String] The messsage to append to the standard error description. </param>
''' <param name="a_error">     [<see cref="UserDefinedError"/>] to raise. </param>
Public Sub GuardExpression(ByVal a_throw As Boolean, ByVal a_source As String, _
        ByVal a_message As String, _
        ByVal a_error As UserDefinedError)
    
    If a_throw Then UserDefinedErrors.RaiseError a_error, a_source, a_message

End Sub

''' <summary>   Raises a run-time error if the specified instance isn't the default /pre-declated. </summary>
''' <param name="a_instance">           [Object] The instance. </param>
''' <param name="a_defaultInstance">    [Object] The default /pre-declated instance. </param>
''' <param name="a_source">    [String, Optional] The source of the error. </param>
''' <param name="a_message">   [String, Optional] The messsage to append to the standard error description. </param>
Public Sub GuardNonDefaultInstance(ByVal a_instance As Object, ByVal a_defaultInstance As Object, _
    Optional ByVal a_source As String = VBA.vbNullString, _
    Optional ByVal a_message As String = "Method should be invoked from the default/predeclared instance of this class.")
    
    Debug.Assert TypeName(a_instance) = TypeName(a_defaultInstance)
    
    GuardExpression Not a_instance Is a_defaultInstance, _
                    IIf(IsMissing(a_source), TypeName(a_instance), a_source), _
                    a_message, UserDefinedErrors.InvalidOperationError

End Sub

''' <summary>   Raises a run-time error if the specified instance is the default instance. </summary>
''' <param name="a_instance">           [Object] The instance. </param>
''' <param name="a_defaultInstance">    [Object] The default /pre-declated instance. </param>
''' <param name="a_source">    [String, Optional] The source of the error. </param>
''' <param name="a_message">   [String, Optional] The messsage to append to the standard error description. </param>
Public Sub GuardDefaultInstance(ByVal a_instance As Object, ByVal a_defaultInstance As Object, _
    Optional ByVal a_source As String = VBA.vbNullString, _
    Optional ByVal a_message As String = "Method should be invoked from a new instance of this class.")

    Debug.Assert TypeName(a_instance) = TypeName(a_defaultInstance)
    
    GuardExpression a_instance Is a_defaultInstance, a_source, a_message, _
                    UserDefinedErrors.InvalidOperationError

End Sub

''' <summary>   Raises a run-time error if the specified object reference is already set. </summary>
''' <param name="a_value">     [Variant] The object reference that might already exist. </param>
''' <param name="a_source">    [String, Optional] The source of the error. </param>
''' <param name="a_message">   [String, Optional] The messsage to append to the standard error description. </param>
Public Sub GuardDoubleInitialization(ByVal a_value As Variant, _
    Optional ByVal a_source As String = VBA.vbNullString, _
    Optional ByVal a_message As String = "Value is already initialized.")
    
    Dim p_throw As Boolean
    If IsObject(a_value) Then
        p_throw = Not a_value Is Nothing
    Else
        p_throw = a_value <> GetDefaultValue(VBA.VarType(a_value))
    End If
    GuardExpression p_throw, a_source, a_message, UserDefinedErrors.InvalidOperationError

End Sub

''' <summary>   Returns the default value to the specified variable type. </summary>
''' <param name="a_variableType">   [VBA.VbVarType] subtype of a variable, or the type
'''                                 of an object's default property. </param>
''' <returns>   [Variant]. </returns>
Private Function GetDefaultValue(ByVal a_variableType As VBA.VbVarType) As Variant

    Select Case a_variableType
        
        Case VBA.VbVarType.vbString
            GetDefaultValue = VBA.vbNullString
        
        Case VBA.VbVarType.vbBoolean
            GetDefaultValue = False
        
        Case VBA.VbVarType.vbByte, VBA.VbVarType.vbCurrency, VBA.VbVarType.vbDate, _
             VBA.VbVarType.vbDecimal, VBA.VbVarType.vbDouble, VBA.VbVarType.vbInteger, _
             VBA.VbVarType.vbLong, VBA.VbVarType.vbSingle
            GetDefaultValue = 0
        
        Case VBA.VbVarType.vbArray, VBA.VbVarType.vbEmpty, VBA.VbVarType.vbVariant
            GetDefaultValue = Empty
        
        Case VBA.VbVarType.vbNull
            GetDefaultValue = Null
        
        Case VBA.VbVarType.vbObject
            Set GetDefaultValue = Nothing
#If VBA7 Then
#If Win64 Then
        Case VBA.VbVarType.vbLongLong 'VF: include longptr(?) if using GuardClauses with WinAPIs
            GetDefaultValue = 0
#End If
#End If
    End Select
End Function

''' <summary>   Raises a run-time error if the specified object reference is Nothing. </summary>
''' <param name="a_instance">   [Object] The instance. </param>
''' <param name="a_source">     [String, Optional] The source of the error. </param>
''' <param name="a_message">    [String, Optional] The messsage to append to the standard error description. </param>
Public Sub GuardNullReference(ByVal a_instance As Object, _
    Optional ByVal a_source As String = VBA.vbNullString, _
    Optional ByVal a_message As String = "Object reference cannot be Nothing.")
    
    GuardExpression a_instance Is Nothing, a_source, a_message, _
                    cc_isr_Core.UserDefinedErrors.NullArgumentError

End Sub

''' <summary>   Raises a run-time error if the specified string is empty. </summary>
''' <param name="a_instance">           [Object] The instance. </param>
''' <param name="a_defaultInstance">    [Object] The default /pre-declated instance. </param>
''' <param name="a_source">    [String, Optional] The source of the error. </param>
''' <param name="a_message">   [String, Optional] The messsage to append to the standard error description. </param>
Public Sub GuardEmptyString(ByVal a_value As String, _
    Optional ByVal a_source As String = VBA.vbNullString, _
    Optional ByVal a_message As String = "String cannot be empty.")
    
    GuardExpression a_value = VBA.vbNullString, a_source, a_message, _
                    cc_isr_Core.UserDefinedErrors.InvalidArgumentError
                    
End Sub


