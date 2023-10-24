VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Timer Sheet. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m_lastErrorMessageCell As String = "C13"

Private Type this_
    PauseRequested As Boolean
    LastErrorMessage As String
End Type

Private This As this_

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
    cc_isr_Core_IO.ErrorMessageBuilder.AppendErrSource p_procedureName, Me.Name, ThisWorkbook
    
    ' display the error message
    Me.LastError = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage()
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

Sub ResetClock()
    
    Range("B4") = TimeValue("00:00:15")

    ' display an unreported errors.
    Me.TraceErrorQueue

End Sub

Sub StartTimer()
    This.PauseRequested = False
    If m_eventTimer Is Nothing Then
      Set m_eventTimer = cc_isr_Core.Factory.NewEventTimer().Initialize(100)
      m_eventTimer.StartTimer
    Else
      m_eventTimer.ResumeTimer
    End If
    
    ' display an unreported errors.
    Me.TraceErrorQueue

End Sub

Sub PauseTimer()
  
  This.PauseRequested = True

  ' display an unreported errors.
  Me.TraceErrorQueue

End Sub

Sub AbortTimer()

  If Not m_eventTimer Is Nothing Then
    m_eventTimer.StopTimer
    Set m_eventTimer = Nothing
  End If
  
  ' display an unreported errors.
  Me.TraceErrorQueue
  
End Sub

Private Sub m_eventTimer_Tick()
    On Error Resume Next
    If Range("B4").Value <= 0 Or This.PauseRequested Then
        If Not m_eventTimer Is Nothing Then m_eventTimer.PauseTimer
    Else
        Range("B4") = Range("B4") - TimeValue("00:00:01")
    End If
    On Error GoTo 0
End Sub

''' <summary>   Displays the last error. </summary>
Public Property Let LastError(ByVal a_value As String)
    
    Range(m_lastErrorMessageCell).Value = a_value

End Property

''' <summary>   Dequeues and displays the
''' <see cref="cc_isr_Core_IO.UserDefinedErrors.ErrorsQueue">errors queue</see>. </summary>
Public Sub TraceErrorQueue()
    
    If cc_isr_Core_IO.UserDefinedErrors.QueuedErrorCount > 0 Then
        Me.LastError = cc_isr_Core_IO.ErrorMessageBuilder.BuildStandardErrorMessage(VBA.vbCrLf)
    Else
        Me.LastError = VBA.vbNullString
    End If

End Sub

