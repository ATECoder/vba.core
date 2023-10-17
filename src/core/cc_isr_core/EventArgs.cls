VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Event arguments. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Initializes and returns this object. </summary>
''' <returns>   [<see cref="EventArgs"/>]. </returns>
Public Function Initialize() As EventArgs
    Set Initialize = Me
End Function
