VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Disposable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that encapsulates the necessary casting operation to invoke IDisposable.Dispose."
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' Disposable.cls
'
' Description: An object that encapsulates the necessary casting operation to invoke IDisposable.Dispose.
'
' Dependencies:
'
' IDisposable.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Tries to dispose a target object. </summary>
''' <param name="a_target">          [Object] The object which to dispose. </param>
''' <param name="a_tryLateBound">    [Boolean, Optional, False] True to attempt to force dispose on the object if
'''                                   it was not cast to Idisposable. </param>
'''< returns>   True if a Dispose method was successfully invoked (early-bound IDisposable.Dispose unless
''' specified otherwise). </returns>
Public Function TryDispose(ByVal a_target As Object, Optional ByVal a_tryLateBound As Boolean = False) As Boolean
Attribute TryDispose.VB_Description = "True if a Dispose method was successfully invoked (early-bound IDisposable.Dispose unless specified otherwise)."
    
    Dim p_disposableTarget As IDisposable
    
    If TryCastToDisposable(a_target, p_disposableTarget) Then
        
        ' if the object can be cast, then dispose
        p_disposableTarget.Dispose
        TryDispose = True
    
    ElseIf TryLateBound Then
        
        ' if allowing a late bound dispose, just assume this can be done.
        On Error Resume Next
        a_target.Dispose
        TryDispose = 0 = VBA.Err.Number
        On Error GoTo 0
        
    End If
    
End Function

''' <summary>   Tries to cast the object to <see cref="cc_isr_Core.IDisposable"/>. </summary>
''' <param name="a_target">    [Object] The object which to cast. </param>
''' <param name="a_result">    [IDisposable] The cast object if success. </param>
'''< returns>   [IDisposable] of null. </returns>
Private Function TryCastToDisposable(ByVal a_target As Object, ByRef a_result As IDisposable) As Boolean
    
    If TypeOf a_target Is IDisposable Then
        Set a_result = a_target
        TryCastToDisposable = True
    End If
    
End Function
