VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Type ThisData
    PauseRequested As Boolean
End Type

Private This As ThisData

Private WithEvents m_eventTimer As cc_isr_Core.EventTimer
Attribute m_eventTimer.VB_VarHelpID = -1

''' <summary>   Dispose of the worksheet objects. </summary>
Public Sub Dispose()

    Const p_procedureName As String = "Dispose"
    
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    If Not m_eventTimer Is Nothing Then
      m_eventTimer.StopTimer
      Set m_eventTimer = Nothing
    End If
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    ThisWorkbook.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    MsgBox ThisWorkbook.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

Sub ResetClock()
    Range("B4") = TimeValue("00:00:15")
End Sub

Sub StartTimer()
  This.PauseRequested = False
  If m_eventTimer Is Nothing Then
    Set m_eventTimer = cc_isr_Core.Factory.NewEventTimer().Initialize(100)
    m_eventTimer.StartTimer
  Else
    m_eventTimer.ResumeTimer
  End If
End Sub

Sub PauseTimer()
  This.PauseRequested = True
End Sub

Sub AbortTimer()

  If Not m_eventTimer Is Nothing Then
    m_eventTimer.StopTimer
    Set m_eventTimer = Nothing
  End If
  
End Sub

Private Sub m_eventTimer_Tick()
    On Error Resume Next
    If Range("B4").value <= 0 Or This.PauseRequested Then
        If Not m_eventTimer Is Nothing Then m_eventTimer.PauseTimer
    Else
        Range("B4") = Range("B4") - TimeValue("00:00:01")
    End If
    On Error GoTo 0
End Sub


