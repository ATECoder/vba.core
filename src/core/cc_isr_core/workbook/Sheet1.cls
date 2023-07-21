VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' https://stackoverflow.com/questions/20783170/pass-array-to-paramarray

Public Sub RunTest()
    MethodA "a", "b", "C"
End Sub


Public Sub MethodA(ParamArray a_tokens() As Variant)
    Dim p_tokens() As Variant
    p_tokens = Sheet1.UnboxParameterArray(a_tokens)
    MethodB a_tokens
End Sub

Public Sub MethodB(ParamArray a_tokens() As Variant)
    Dim i As Integer, p_tokens() As Variant
    p_tokens = Sheet1.UnboxParameterArray(a_tokens)
    For i = 0 To UBound(p_tokens)
        Debug.Print StringExtensions.StringFormat("i: {0} prm: {1} ", i, p_tokens(i))
    Next i
    MethodC a_tokens
End Sub

Public Sub MethodC(ParamArray a_tokens() As Variant)
    Dim i As Integer, p_tokens() As Variant
    p_tokens = Sheet1.UnboxParameterArray(a_tokens)
    For i = 0 To UBound(p_tokens)
        Debug.Print StringExtensions.StringFormat("i: {0} prm: {1} ", i, p_tokens(i))
    Next i
End Sub

Public Function UnboxParameterArray(ParamArray a_tokens() As Variant) As Variant
    Dim p_tokens() As Variant, p_workArray() As Variant
    p_tokens = a_tokens
    Do While VarType(p_tokens(0)) >= vbArray And UBound(p_tokens) < 1
        p_workArray = p_tokens(0)
        p_tokens = p_workArray
    Loop
    UnboxParameterArray = p_tokens
End Function


