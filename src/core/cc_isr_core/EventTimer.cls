VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   An event timer. </summary>
''' <remarks>   Dependencies: IRunnable.cls.
'''
''' This timer comes from <see href="https://github.com/sancarn/stdVBA/blob/master/src/stdTimer.cls"/>
''' Module greatly inspired by Cristian Buse's SafeTimers:
''' <see href="https://github.com/cristianbuse/Excel-VBA-SafeTimers"/>
'''
''' Inspirations used:
''' * Use of remote workbook
''' * Use of VBOM and simplified registry setting
'''
''' Comparison of SaveTimers-Native and EventTimer:
''' * EventTimer will not raise the Tick event while displaying MsgBox or InputBox, however SafeTimers will
''' * EventTimer will raise events while Edit mode is "Edit" or "Point", SafeTimers-Native do not.
''' * EventTimer uses COM events (specifically WorksheetChange), SafeTimers-Native uses direct COM calls as a communication mechanism
''' * EventTimer can be bound to <see cref="IRunnable"/> callback OR via event interop. SafeTimers-Native can only be hooked into with Module macro names
'
''' Limitations:
''' * Neither set of timers raise events while debugging code.
''' * Neither set of timers will be called exactly when it needs to. There will be a little bit of error which may have to be corrected.
'''
''' Architecture:
''' When a timer is initialized it:
'''   * Ensures VBOM is enabled
'''   * Creates a `new Application` object and adds a workbook to this new application.
'''   * Adds Timer code to this workbook and schedules this code to run with `Application.OnTime Now(), "MainLoop"`
''' Every time the timer loops
'''   * A change to the remote workbook's worksheet is made
'''   * In the original workbook we register to receive these events, and receive a `This.remoteWorkbook_SheetChange` event.
'''   * Here we raise the event, and call the callback if supplied.
''' When this class is terminated
'''   * It will also tell the remote workbook to close.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

'Set DEBUG_MODE to true if you want to see the remote workbook
#Const DEBUG_MODE = False

''' <summary>   The name of the timer module. </summary>
Const TIMER_MODULE As String = "TimerMain"

''' <summary>   The timer start method. </summary>
Const TIMER_START_CALL As String = "TimerMain.StartTimer"

''' <summary>   The timer stop method. </summary>
Const TIMER_STOP_CALL As String = "TimerMain.StopTimer"

''' <summary>   The timer pause method. </summary>
Const TIMER_PAUSE_CALL As String = "TimerMain.PauseTimer"

''' <summary>   The timer resume method. </summary>
Const TIMER_RESUME_CALL As String = "TimerMain.ResumeTimer"

''' <summary>   Called every time the timer loops. </summary>
Public Event Tick()

Private WithEvents m_remoteWorkbook As Workbook
Attribute m_remoteWorkbook.VB_VarHelpID = -1

Private Type this_
    
    Callback As IRunnable
    
    ' true when the workbook begins closing
    WorkbookClosing As Boolean
    
    ' the timer interval in milliseconds.
    Interval As Long

End Type

Private This As this_

''' <summary>   Instantiates and initializes a <see cref="EventTimer"/> object. </summary>
''' <remarks>   The frequency isn't guaranteed but it will be as close as humanly possible. </remarks>
''' <param name="a_interval">   [Long] Number of milliseconds between tick events. </param>
''' <param name="a_callback">   [IRunnable, Nothing] Callback object. Must contain a Run() method. </param>
''' <returns>   [<see cref="EventTimer"/>]. </returns>
Public Function Initialize(ByVal a_interval As Long, Optional ByVal a_callback As IRunnable = Nothing) As EventTimer

    This.Interval = a_interval
    
    Set This.Callback = a_callback
    
    If Not GetVBOM Then Call SetVBOM(True)
    
    Set m_remoteWorkbook = CreateRemoteWorkbook()
    
    If Not GetVBOM Then Call SetVBOM(False)
    
    ' inject the timer code into the timer module.
    
    Call InjectCode(m_remoteWorkbook, TIMER_MODULE, GetTimerCode(a_interval))
    
    Set Initialize = Me
    
End Function

''' <summary>   Creates and returns a remote workbook. This is a close equivalent of generating a new thread,
'''             however these threads will be in sync (due to usage of COM events). </summary>
''' <param name="a_visibleWorkbook">    [Boolean, True] to see the workbook if in debug mode. </param>
''' <returns>   [<see cref="Workbook"/>]. </returns>
Private Function CreateRemoteWorkbook(Optional ByVal a_visibleWorkbook As Boolean = True) As Workbook
    
    Dim p_application As Application: Set p_application = New Application
    #If DEBUG_MODE Then
        p_application.Visible = a_visibleWorkbook
    #Else
        p_application.Visible = False
        p_application.ScreenUpdating = False
    #End If
    p_application.PrintCommunication = False
    p_application.DisplayAlerts = False
    Set CreateRemoteWorkbook = p_application.Workbooks.Add

End Function

''' <summary>   Inject VBA code into a workbook. </summary>
''' <param name="a_workbook">      [Workbook] The workbook to inject code into. </param>
''' <param name="a_moduleName">    [String] The module name to create. </param>
''' <param name="a_code">          [String] a reference to the code to insert into the module. </param>
''' <param name="a_moduleType">    [Long, 1] The type of code module to create (Module=1, Class=2). </param>
Public Sub InjectCode(ByVal a_workbook As Workbook, ByVal a_moduleName As String, _
        ByRef a_code As String, Optional ByVal a_moduleType As Long = 1)
        
  With a_workbook.VBProject.VBComponents.Add(a_moduleType).CodeModule
    If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
    .AddFromString a_code
    .Parent.Name = a_moduleName
  End With
End Sub

''' <summary>   Check if VB Object model is extensible. </summary>
''' <returns>   [Boolean] True if VBE Is Extensible (i.e. checked within the IDE,
'''             rather than the registry). </returns>
Private Function GetVBOM() As Boolean
  On Error Resume Next
  GetVBOM = Not Application.VBE.ActiveVBProject Is Nothing
End Function

''' <summary>   Check whether a registry key exists.  </summary>
''' <remarks>   Author <see href="https://github.com/cristianbuse">Cristian Buse</see>. </remarks>
''' <param name="a_visibleWorkbook">    [Boolean} True to enable VBOM, False otherwise. </param>
''' <returns>   [Boolean] True if key exists, False otherwise. </returns>
Private Function SetVBOM(ByVal a_newValue As Boolean) As Boolean
  
  Dim i As Long: i = IIf(a_newValue, 1, 0)
  #If Mac Then
    Shell "defaults write com.Microsoft.Excel AccessVBOM -int " & i
    SetVBOM = True
  #Else
    Dim p_registeryKey As String
    p_registeryKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version _
          & "\" & Replace(Application.Name, "Microsoft ", vbNullString) _
          & "\Security\AccessVBOM"
    On Error Resume Next
    CreateObject("WScript.Shell").RegWrite p_registeryKey, i, "REG_DWORD"
    SetVBOM = (Err.Number = 0)
    On Error GoTo 0
  #End If

End Function

''' <summary>   Obtains the code for a simple timer. </summary>
''' <param name="a_interval">   [Long] Number of milliseconds between tick events. </param>
''' <returns>   [String]. </returns>
Private Function GetTimerCode(ByVal a_interval As Long) As String
  
  Dim s As String: s = "' Auto-generated by EventTimer"
  s = s & vbCrLf & "#If Mac Then"
  s = s & vbCrLf & "  #If VBA7 Then"
  s = s & vbCrLf & "    Private Declare PtrSafe Sub USleep Lib ""/usr/lib/libc.dylib"" Alias ""usleep"" (ByVal a_microSecs As Long)"
  s = s & vbCrLf & "  #Else"
  s = s & vbCrLf & "    Private Declare Sub USleep Lib ""/usr/lib/libc.dylib"" Alias ""usleep"" (ByVal a_microSecs As Long)"
  s = s & vbCrLf & "  #End If"
  s = s & vbCrLf & "#Else 'Windows"
  s = s & vbCrLf & "  #If VBA7 Then"
  s = s & vbCrLf & "    Public Declare PtrSafe Sub Sleep Lib ""kernel32"" (ByVal a_milliSecs As Long)"
  s = s & vbCrLf & "  #Else"
  s = s & vbCrLf & "    Public Declare  Sub Sleep Lib ""kernel32"" Alias ""Sleep"" (ByVal a_milliSecs As Long)"
  s = s & vbCrLf & "  #End If"
  s = s & vbCrLf & "#End If"
  s = s & vbCrLf & "Private Type this_"
  s = s & vbCrLf & "  Interval as Long"
  s = s & vbCrLf & "  Running as Boolean"
  s = s & vbCrLf & "  Enabled as Boolean"
  s = s & vbCrLf & "End Type"
  s = s & vbCrLf & "Private This As this_"
  s = s & vbCrLf & "Sub MainLoop()"
  s = s & vbCrLf & "  While This.Running"
  s = s & vbCrLf & "    If This.Enabled Then"
  s = s & vbCrLf & "      Sheet1.Range(""A1"").Value = ""Tick"""
  s = s & vbCrLf & "    End If"
  s = s & vbCrLf & "    sleep This.Interval"
  s = s & vbCrLf & "    DoEvents"
  s = s & vbCrLf & "  Wend"
  s = s & vbCrLf & "  ThisWorkbook.Close false" 'cleanup
  s = s & vbCrLf & "End Sub"
  s = s & vbCrLf & "Sub StartTimer()"
  s = s & vbCrLf & "  This.Interval = " & a_interval
  s = s & vbCrLf & "  This.Enabled = true"
  s = s & vbCrLf & "  This.Running = true"
  s = s & vbCrLf & "  Application.OnTime now(), ""MainLoop""" 'initialize main loop asynchronously
  s = s & vbCrLf & "End Sub"
  s = s & vbCrLf & "Sub PauseTimer()"
  s = s & vbCrLf & "  This.Enabled = false"
  s = s & vbCrLf & "End Sub"
  s = s & vbCrLf & "Sub ResumeTimer()"
  s = s & vbCrLf & "  This.Enabled = true"
  s = s & vbCrLf & "End Sub"
  s = s & vbCrLf & "Sub StopTimer()"
  s = s & vbCrLf & "  This.Running = false"
  s = s & vbCrLf & "End Sub"
  GetTimerCode = s

End Function

''' <summary>   Handles the construction of this class. </summary>
''' <remarks>   On initialize set reference to myself. </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Handles the termination of this class. </summary>
''' <remarks>   On terminate, stop timer. This will cause workbook to close. </remarks>
Private Sub Class_Terminate()

    ' errors are trapped in line here becuase the remote workbook may not be available
    ' even if not disposed.
    
    On Error Resume Next
    
    ' 20230704: this worked well.
    
    This.WorkbookClosing = False

    If Not m_remoteWorkbook Is Nothing Then
        
        ' this should also close the workbook
        
        ' this caused an exception; the remote workbook was said to be unavailable.
       
        Call m_remoteWorkbook.Application.Run(TIMER_STOP_CALL)
        
        ' wait for one timer interval to complete to ensure the timer
        ' events no longer trigger.
        Dim p_isDone As Boolean: p_isDone = False
        Dim p_waitMilliseconds As Long: p_waitMilliseconds = 2 * This.Interval
        Dim p_seconds As Double: p_seconds = 0.001 * p_waitMilliseconds
        ' wait at least two computer clock cycles (54 ms)
        If p_seconds < 0.11 Then p_seconds = 0.11
        Dim p_endTime As Double
        p_endTime = cc_isr_Core_IO.CoreExtensions.DaysNow() + _
            (p_seconds / cc_isr_Core_IO.CoreExtensions.SecondsPerDay)
        Dim stopper As cc_isr_Core_IO.Stopwatch: stopper = cc_isr_Core_IO.Factory.NewStopwatch()
        Do Until p_isDone
            ' previous code failed to detect error as the workbook closed
            ' while this loop was still running.
            If Err.Number <> 0 Then
                p_isDone = True
            ElseIf m_remoteWorkbook Is Nothing Then
                p_isDone = True
            ElseIf This.WorkbookClosing Then
                p_isDone = True
            ElseIf stopper Is Nothing Then
                p_isDone = True
            ElseIf stopper.ElapsedMilliseconds > p_waitMilliseconds Then
                p_isDone = True
            ElseIf cc_isr_Core_IO.CoreExtensions.DaysNow() > p_endTime Then
                p_isDone = True
            End If
            VBA.DoEvents
        Loop
        
        If Not This.WorkbookClosing Then
          m_remoteWorkbook.Close False
        End If
        
        ' terminating the workbook does not seem to close the excel
        ' instance of this remote workbook.
        Set m_remoteWorkbook = Nothing
        
    End If
    
    On Error GoTo 0
    
End Sub

''' <summary>   Handles the workbook before close event. </summary>
Private Sub m_remoteWorkbook_BeforeClose(Cancel As Boolean)
    This.WorkbookClosing = True
End Sub

''' <summary>   Handles the workbook sheet change event, which raises the <see cref="Tick"/>
''' event. </summary>
Private Sub m_remoteWorkbook_SheetChange(ByVal a_sender As Object, ByVal a_target As Range)
  RaiseEvent Tick
  If Not This.Callback Is Nothing Then This.Callback.Run
End Sub

''' <summary>   Starts the timer. </summary>
Public Sub StartTimer()
  Call m_remoteWorkbook.Application.Run(TIMER_START_CALL)
End Sub

''' <summary>   Stops the timer. </summary>
''' <remarks>   This also closes the workbook. </remarks>
Public Sub StopTimer()
  Call m_remoteWorkbook.Application.Run(TIMER_STOP_CALL)
End Sub

''' <summary>   Halts the timer. </summary>
Public Sub PauseTimer()
  Call m_remoteWorkbook.Application.Run(TIMER_PAUSE_CALL)
End Sub

''' <summary>   Resumes the timer. </summary>
Public Sub ResumeTimer()
  Call m_remoteWorkbook.Application.Run(TIMER_RESUME_CALL)
End Sub

