VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoreExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Core extensions. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unbox a parameter array in case element zero is an array, which occurs when a parameter
''' array argument is passed to another function that expects a ParamArray argument.
''' </summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/20783170/pass-array-to-paramarray"/> <para>
''' When passing a parameter array that is received as an argument in function A to another function
''' B that expects a parameter array, the parameter array is received in function A appeats at element
''' zero of the function B. </para><para>
''' Generally, when each subsequent function receives the parameter array, and passes it on, it increases
''' the depth of the parameter array. </para><para>
''' This method undoes the depth levels added and returns a vector with arguments received as a Variant.
''' <code>
''' </code>
''' </para></remarks>
''' <param name="a_values">   [Variant, ParamArray] The parameter values to be unboxed. </param>
''' <returns>   [Variant, Array] The unboxed parameter values. </returns>
Public Function UnboxParameterArray(ParamArray a_values() As Variant) As Variant

    Dim p_params() As Variant, p_workArray() As Variant
    
    ' When a_values(0) is Array, it is assumed that the
    ' parameter arrayt was passed from another function
    p_params = a_values
    Do While VBA.VarType(p_params(0)) >= VBA.vbArray And UBound(p_params) < 1
        p_workArray = p_params(0)
        p_params = p_workArray
    Loop
    UnboxParameterArray = p_params
    
End Function

''' <param name="a_variableType">   [VBA.VbVarType] subtype of a variable, or the type
'''                                 of an object's default property. </param>
''' <returns>   [Variant]. </returns>
Public Function GetDefaultValue(ByVal a_variableType As VBA.VbVarType) As Variant

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
        Case VBA.VbVarType.vbLongLong ' VF: include longptr(?) if using GuardClauses with WinAPIs
            GetDefaultValue = 0
#End If
#End If
    End Select
End Function

''' <summary>   Queries whether code is running in the IDE (true), or compiled (false). </summary>
''' <returns>   [Boolean]. </returns>
Public Function RunningInIde() As Boolean

    Static p_counter As Variant
    If IsEmpty(p_counter) Then
        p_counter = 1
        Debug.Assert RunningInIde() Or True
        p_counter = p_counter - 1
    ElseIf p_counter = 1 Then
        p_counter = 0
    End If
    RunningInIde = p_counter
 
End Function

