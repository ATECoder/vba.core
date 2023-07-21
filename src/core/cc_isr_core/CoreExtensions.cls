VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoreExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' CoreExtensions.cls
'
' Dependencies:
'
' None
'
' 11
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

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

