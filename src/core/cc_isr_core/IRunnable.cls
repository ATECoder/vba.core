VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRunnable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Represents an object that can execute a command with a parameter array arguments. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary> Call will call the passed function with param array. </summary>
''' <param name="">   [ParamArray Variant()]. </param>
'''<returns>   A Variant. </returns>
Public Function Run(ParamArray a_params() As Variant) As Variant: End Function

