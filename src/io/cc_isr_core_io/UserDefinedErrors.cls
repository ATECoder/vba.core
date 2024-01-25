VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserDefinedErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   User-defined errors. </summary>
''' <remarks>   Dependencies: UserDefinedError.cls. </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    ErrorsDix As Dictionary
    WinsockCleanupError As UserDefinedError
    SocketCreationError As UserDefinedError
    SocketConnectionError As UserDefinedError
    SocketDisconnectionError As UserDefinedError
    SocketOptionSettingError As UserDefinedError
    SocketSendError As UserDefinedError
    SocketReceiveError As UserDefinedError
    SocketBindError As UserDefinedError
    SocketListenError As UserDefinedError
    SocketAcceptError As UserDefinedError
    SocketSetStatusError As UserDefinedError
    WinsockStartingError As UserDefinedError
    NullArgumentError As UserDefinedError
    InvalidArgumentError As UserDefinedError
    InvalidFormatStringError As UserDefinedError
    InvalidOperationError As UserDefinedError
    TimeoutError As UserDefinedError
    FormatError As UserDefinedError
    Ieee488DeviceError As UserDefinedError
    FirstUserDefinedErrorNumber As Long
    NextUserDefinedErrorNumber As Long
    ErrorsQueue As Queue
    ErrorsArchiveStack As Stack
    LastQueuedError As UserDefinedError
End Type

Private This As this_

''' <summary>   Initializes the user defined errors. </summary>
Private Sub Class_Initialize()

    Set This.ErrorsQueue = Factory.CreateQueue(100)
    Set This.ErrorsArchiveStack = Factory.CreateStack(100)
    Set This.ErrorsDix = New Dictionary
    Dim p_dummyError As UserDefinedError
    Set p_dummyError = New UserDefinedError
    This.FirstUserDefinedErrorNumber = VBA.vbObjectError + p_dummyError.UserErrorNumberBase + 1
    This.NextUserDefinedErrorNumber = This.FirstUserDefinedErrorNumber
    
    Set This.WinsockStartingError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "WinsockStartingError", _
                                        "Failed starting Winsock using Winsock WSAStartup().")
    AddUserDefinedError Me.WinsockStartingError
    
    Set This.WinsockCleanupError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "WinsockCleanupError", _
                                        "Failed disposing Winsock using Winsock WSACleanup().")
    AddUserDefinedError Me.WinsockCleanupError
    
    Set This.SocketCreationError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketCreationError", _
                                        "Failed creating socket using Winsock socket().")
    AddUserDefinedError Me.SocketCreationError
    
    Set This.SocketConnectionError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketConnectionError", _
                                        "Failed connecting socket using Winsock connect().")
    AddUserDefinedError Me.SocketConnectionError
    
    Set This.SocketDisconnectionError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketDisconnectionError", _
                                            "Failed disconnecting socket using Winsock closesocket().")
    AddUserDefinedError Me.SocketDisconnectionError
    
    Set This.SocketOptionSettingError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketOptionSettingError", _
                                            "Failed setting socket options using Winsock setsockopt().")
    AddUserDefinedError Me.SocketOptionSettingError
    
    Set This.SocketSendError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketSendError", _
                                    "Failed send a socket message using Winsock send().")
    AddUserDefinedError Me.SocketSendError

    Set This.SocketReceiveError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketReceiveError", _
                                    "Failed receiving a socket message using Winsock recv().")
    AddUserDefinedError Me.SocketReceiveError

    Set This.SocketBindError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketBindError", _
                                    "Failed binding a socket using Winsock bind().")
    AddUserDefinedError Me.SocketBindError

    Set This.SocketListenError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketListenError", _
                                    "Failed stating a socket listening using Winsock listen().")
    AddUserDefinedError Me.SocketListenError

    Set This.SocketAcceptError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketAcceptError", _
                                    "Failed permitting an incoming connection attempt on a socket using Winsock accept().")
    AddUserDefinedError Me.SocketAcceptError

    Set This.SocketSetStatusError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketSetStatusError", _
                                        "Failed determining the socket status using Winsock select().")
    AddUserDefinedError Me.SocketSetStatusError

    Set This.NullArgumentError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "nullArgumentError", _
                                    "Null argument error.")
    AddUserDefinedError Me.NullArgumentError

    Set This.InvalidArgumentError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "InvalidArgumentError", _
                                        "Invalid argument error.")
    AddUserDefinedError Me.InvalidArgumentError

    Set This.InvalidFormatStringError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "InvalidFormatStringError", _
                                            "Invalid format string.")
    AddUserDefinedError Me.InvalidFormatStringError

    Set This.InvalidOperationError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "InvalidOperationError", _
                                        "Invalid operation error.")
    AddUserDefinedError Me.InvalidOperationError
    
    Set This.TimeoutError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "TimeoutError", _
                                "Timeout error.")
    AddUserDefinedError Me.TimeoutError

    Set This.FormatError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "FormatError", _
        "The number indicating an argument to format is less than zero or greater than or equal to the length of the array of arguments.")
    AddUserDefinedError Me.FormatError

    Set This.Ieee488DeviceError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "Ieee488DeviceError", _
        "IEEE488 Device error(s) occurred.")
    AddUserDefinedError Me.Ieee488DeviceError

End Sub

''' <summary>   Initializes the error Numbers and descriptions. </summary>
''' <returns>   [<see cref="UserDefinedErrors"/>]. </returns>
Public Function Initialize() As UserDefinedErrors
    Class_Initialize
    Set Initialize = Me
End Function

''' <summary>   Adds a user defined error to the error dictionary and increments the
''' <see cref="NextUserDefinedErrorNumber"/>. </summary>
''' <para name"a_userError">   The error of type <see cref="UserDefinedError"/>   </param>
Private Sub AddUserDefinedError(ByVal a_userError As UserDefinedError)
    
    ' the error source and timestamp must be provided when raising the error.
    
    a_userError.Source = VBA.vbNullString
    a_userError.Timestamp = 0
    
    This.ErrorsDix.Add a_userError.Number, a_userError
    
    ' set the next error number.
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

End Sub

''' <summary>   Returns true if the specified error number is a user defined error. </summary>
''' <para name"a_errorNumber">   [Long] The error number. </param>
''' <returns>   [Boolean]. </returns>
Public Function IsUserDefinedError(ByVal a_errorNumber As Long) As Boolean
    
    IsUserDefinedError = This.FirstUserDefinedErrorNumber <= a_errorNumber And _
                         This.NextUserDefinedErrorNumber > a_errorNumber

End Function

''' <summary>   Returns true if the specified error exists. </summary>
''' <para name"a_userError">   [<see cref="UserDefinedError"/>] The error. </param>
''' <returns>   [Boolean]. </returns>
Public Function UserDefinedErrorExists(ByVal a_userError As UserDefinedError) As Boolean
    
    UserDefinedErrorExists = This.ErrorsDix.Exists(a_userError.Number)

End Function

''' <summary>   Returns true if the specified error exists. </summary>
''' <para name"a_userErrorNumber">   [Long] The error Number </param>
''' <returns>   [Boolean]. </returns>
Public Function UserDefinedErrorNumberExists(ByVal a_userErrorNumber As Long) As Boolean
    
    UserDefinedErrorNumberExists = This.ErrorsDix.Exists(a_userErrorNumber)

End Function

''' <summary>   Returns a user defined error. </summary>
''' <param name="a_key">   [Long] The user defined error Number. </param>
''' <returns>   [<see cref="UserDefinedError"/>] An existing or a new 'unknown' user defined error.. </returns>
Public Function GetUserDefinedError(ByVal a_key As Long) As UserDefinedError
    
    If This.ErrorsDix.Exists(a_key) Then
        Set GetUserDefinedError = This.ErrorsDix.Item(a_key)
    Else
        Set GetUserDefinedError = Factory.NewUserDefinedError.Initialize(a_key, _
                                        "UnknownUserDefinedError", "Unknown error.")
    End If

End Function

''' <summary>   Get the error that occurred starting Winsock. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get WinsockStartingError() As UserDefinedError
    
    Set WinsockStartingError = This.WinsockStartingError

End Property

''' <summary>   Get the error that occurred ending Winsock. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get WinsockCleanupError() As UserDefinedError
    
    Set WinsockCleanupError = This.WinsockCleanupError

End Property

''' <summary>   Get the error that occurred creating a socket. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get SocketCreationError() As UserDefinedError
    
    Set SocketCreationError = This.SocketCreationError

End Property

''' <summary>   Get the error that occurred connecting a socket. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get SocketConnectionError() As UserDefinedError
    
    Set SocketConnectionError = This.SocketConnectionError

End Property

''' <summary>   Get the error that occurred disconnecting a socket. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get SocketDisconnectionError() As UserDefinedError
    
    Set SocketDisconnectionError = This.SocketDisconnectionError

End Property

''' <summary>   Get the error that occurred setting a socket option. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get SocketOptionSettingError() As UserDefinedError
    
    Set SocketOptionSettingError = This.SocketOptionSettingError

End Property

''' <summary>   Get the error that occurred sending a socket message. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get SocketSendError() As UserDefinedError
    
    Set SocketSendError = This.SocketSendError

End Property

''' <summary>   Get the error that occurred receiving a socket message. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get SocketReceiveError() As UserDefinedError
    
    Set SocketReceiveError = This.SocketReceiveError

End Property

''' <summary>   Get the error that occurred binding a socket to an address. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get SocketBindError() As UserDefinedError
    
    Set SocketBindError = This.SocketBindError

End Property

''' <summary>   Get the error that occurred starting to listen on a socket. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get SocketListenError() As UserDefinedError
    
    Set SocketListenError = This.SocketListenError

End Property

''' <summary>   Get the error that occurred accepting an incoming connection. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get SocketAcceptError() As UserDefinedError
    
    Set SocketAcceptError = This.SocketAcceptError

End Property

''' <summary>   Get the error that occurred determining the status of a socket set. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get SocketSetStatusError() As UserDefinedError
    
    Set SocketSetStatusError = This.SocketSetStatusError

End Property

''' <summary>   Get the error that occurred when the method argument is null. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get NullArgumentError() As UserDefinedError
    
    Set NullArgumentError = This.NullArgumentError

End Property

''' <summary>   Get the error that occurred when the method argument is invalid. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get InvalidArgumentError() As UserDefinedError
    
    Set InvalidArgumentError = This.InvalidArgumentError

End Property

''' <summary>   Get the error that occurred when the format string is invalid. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get InvalidFormatStringError() As UserDefinedError
    
    Set InvalidFormatStringError = This.InvalidFormatStringError

End Property

''' <summary>   Get the error that occurred when the method operation is invalid. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get InvalidOperationError() As UserDefinedError
    
    Set InvalidOperationError = This.InvalidOperationError

End Property

''' <summary>   Get the error that occurred when the method operation times out. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get TimeoutError() As UserDefinedError
    
    Set TimeoutError = This.TimeoutError

End Property

''' <summary>   Get the error that occurred when a format error occurred. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get FormatError() As UserDefinedError
    
    Set FormatError = This.FormatError

End Property

''' <summary>   Get the error that occurred when a IEEE488 device error(s) occurred. </summary>
''' <value>   [<see cref="UserDefinedError"/>]. </value>
Public Property Get Ieee488DeviceError() As UserDefinedError
    
    Set Ieee488DeviceError = This.Ieee488DeviceError

End Property

''' <summary>   gets the error Number of the next user defined error that can
''' be added to the collection of user defined errors. </summary>
''' <value>   [Long]. </value>
Public Property Get NextUserDefinedErrorNumber() As Long
    NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber
End Property

''' <summary>   Sets the error Number of the next user defined error that can
''' be added to the collection of user defined errors. </summary>
''' <param name="a_value">    [Long] the next error Number. </param>
Public Property Let NextUserDefinedErrorNumber(ByVal a_value As Long)
    This.NextUserDefinedErrorNumber = a_value
End Property

''' <summary>   Gets the stack of enqueued errors. </summary>
''' <remarks>   Pending removal. </remarks>
''' <value>   [<see cref="cc_isr_Core_IO.Queue"/>]. </value>
Private Property Get ErrorsQueue() As Queue
    Set ErrorsQueue = This.ErrorsQueue
End Property

''' <summary>   Peeks at an error from the queue. </summary>
''' <param name="a_position">   [Optional, Integer, 1] the position of the item
'''                             where the first queued item is at 1. </param>
''' <returns>   [<see cref="cc_isr_Core_IO.UserDefinedError"/>]. </value>
Public Function PeekQueue(Optional ByVal a_position As Integer = 1) As UserDefinedError

    If Me.QueuedErrorCount = 0 Then
        Set PeekQueue = Nothing
    Else
        Set PeekQueue = This.ErrorsQueue.Peek(a_position)
    End If
    
End Function

''' <summary>   Dequeue an error from the queue. </summary>
''' <returns>   [<see cref="cc_isr_Core_IO.UserDefinedError"/>]. </value>
Public Function Dequeue() As UserDefinedError

    If Me.QueuedErrorCount = 0 Then
        Set Dequeue = Nothing
    Else
        Set Dequeue = This.ErrorsQueue.Dequeue()
    End If
    
    If Me.QueuedErrorCount = 0 Then
        Set This.LastQueuedError = Nothing
    End If
    
End Function

''' <summary>   Gets the number of queued errors. </summary>
''' <value>   [Integer]. </value>
Public Property Get QueuedErrorCount() As Integer
    QueuedErrorCount = This.ErrorsQueue.Count
End Property

''' <summary>   Gets the stack of archived errors. </summary>
''' <remarks>   Pending removal. </remarks>
''' <value>   [<see cref="cc_isr_Core_IO.Stack"/>]. </value>
Private Property Get ErrorsArchiveStack() As Stack
    Set ErrorsArchiveStack = This.ErrorsArchiveStack
End Property

''' <summary>   Push an error onto the error archive. </summary>
''' <returns>   [<see cref="cc_isr_Core_IO.UserDefinedError"/>]. </value>
Public Function Push(ByVal a_error As UserDefinedError) As UserDefinedError

    This.ErrorsArchiveStack.Push a_error
    Set Push = a_error
    
End Function

''' <summary>   Peeks at an error from the archive stack. </summary>
''' <param name="a_position">   [Optional, Integer, 1] the position of the item
'''                             where the top of the stack is at 1. </param>
''' <returns>   [<see cref="cc_isr_Core_IO.UserDefinedError"/>]. </value>
Public Function PeekArchive(Optional ByVal a_position As Integer = 1) As UserDefinedError

    If Me.ArchivedErrorCount = 0 Then
        Set PeekArchive = Nothing
    Else
        Set PeekArchive = This.ErrorsArchiveStack.Peek(a_position)
    End If
    
End Function

''' <summary>   Gets the number of archived errors. </summary>
''' <value>   [Integer]. </value>
Public Property Get ArchivedErrorCount() As Integer
    ArchivedErrorCount = This.ErrorsArchiveStack.Count
End Property

''' <summary>   Archives the queued errors and returns the errors as a delimited string. </summary>
''' <param name="a_delimiter">   [String, VBCrLf] delimiter between errors. </param>
''' <returns>   [String]. </returns>
Public Function ArchiveQueuedErrors(Optional ByVal a_delimiter As String = VBA.vbCrLf) As String

    Dim p_builder As String: p_builder = VBA.vbNullString
    
    Dim p_error As UserDefinedError
    While This.ErrorsQueue.Count > 0
        
        DoEvents
        
        ' dequeue errors from the queue to the string builder and move them to the archive.
        Set p_error = This.ErrorsQueue.Dequeue
        
        If VBA.Len(p_builder) > 0 Then p_builder = p_builder & a_delimiter
        p_builder = p_builder & p_error.ToString()
        
        ' push the error to the archive, which is circular.
        This.ErrorsArchiveStack.Push p_error
        
    Wend
    
    ' clear the last error
    Set This.LastQueuedError = Nothing
    
    ArchiveQueuedErrors = p_builder

End Function

''' <summary>   Move all archived errors to a delimited string. </summary>
''' <param name="a_delimiter">   [String, VBCrLf] delimiter between errors. </param>
''' <returns>   [String]. </returns>
Public Function PopArchivedErrors(Optional ByVal a_delimiter As String = VBA.vbCrLf) As String

    Dim p_builder As String: p_builder = VBA.vbNullString
    
    Dim p_error As UserDefinedError
    While This.ErrorsArchiveStack.Count > 0
        
        DoEvents
        
        ' pop errors from the archive to the string builder.
        Set p_error = This.ErrorsArchiveStack.Pop
        
        If VBA.Len(p_builder) > 0 Then p_builder = p_builder & a_delimiter
        p_builder = p_builder & p_error.ToString()
        
    Wend
    
    PopArchivedErrors = p_builder

End Function


''' <summary>   Enqueues a new user defined error or append the source to the last user defined error. </summary>
''' <param name="a_error">       [<see cref="UserDefinedError"/>]. </param>
''' <returns>    [<see cref="UserDefinedError"/>]. </returns>
Public Function EnqueueUserDefinedError(ByVal a_error As UserDefinedError) As UserDefinedError

    If This.LastQueuedError Is Nothing Then
        ' if no last error, make this the last error and enqueue it.
        Set This.LastQueuedError = a_error
        This.ErrorsQueue.Enqueue This.LastQueuedError
    ElseIf This.LastQueuedError.Same(a_error) Then
        If 1 <> VBA.InStr(1, This.LastQueuedError.Source, a_error.Source, VBA.VbCompareMethod.vbTextCompare) Then
            ' append the error source if the error is the same as the last error.
            ' but has a different error source as it is possible to trape an error from the
            ' same procedure that raised that error.
            This.LastQueuedError.Source = This.LastQueuedError.Source & " in " & a_error.Source
        End If
    Else
        ' if this is a new error, make this the last error and enqueue it.
        Set This.LastQueuedError = a_error
        This.ErrorsQueue.Enqueue This.LastQueuedError
    End If
    
    ' return the last error.
    Set EnqueueUserDefinedError = This.LastQueuedError
    
End Function

''' <summary>   Enqueues a user defined error optionally updating the error source and timeout
''' and adding a suffix to the description. </summary>
''' <param name="a_error">       [<see cref="UserDefinedError"/>]. </param>
''' <param name="a_source">      [Optional, Variant, String] The error source. </param>
''' <param name="a_suffix">      [Optional, Variant, String] A suffix to append to the message description. </param>
''' <param name="a_timestamp">   [Optional, Variant, Double] The error timestamp. </param>
''' <returns>    [<see cref="UserDefinedError"/>]. </returns>
Public Function EnqueueError(ByVal a_error As UserDefinedError, _
        Optional ByVal a_source As Variant, _
        Optional ByVal a_suffix As Variant, _
        Optional ByVal a_timestamp As Variant) As UserDefinedError

    Dim p_error As New UserDefinedError
    
    If VBA.IsMissing(a_timestamp) Then
        If VBA.IsMissing(a_suffix) Then
            If VBA.IsMissing(a_source) Then
                p_error.FromAppendedUserDefinedError a_error
            Else
                p_error.FromAppendedUserDefinedError a_error, a_source
            End If
        ElseIf VBA.IsMissing(a_source) Then
            p_error.FromAppendedUserDefinedError a_error, a_suffix:=a_suffix
        Else
            p_error.FromAppendedUserDefinedError a_error, a_source, a_suffix
        End If
    ElseIf VBA.IsMissing(a_suffix) Then
        If VBA.IsMissing(a_source) Then
            p_error.FromAppendedUserDefinedError a_error, a_timestamp:=a_timestamp
        Else
            p_error.FromAppendedUserDefinedError a_error, a_source:=a_source, a_timestamp:=a_timestamp
        End If
    ElseIf VBA.IsMissing(a_source) Then
        p_error.FromAppendedUserDefinedError a_error, a_suffix:=a_suffix, a_timestamp:=a_timestamp
    Else
        p_error.FromAppendedUserDefinedError a_error, a_source, a_suffix, a_timestamp
    End If
    
    Set EnqueueError = Me.EnqueueUserDefinedError(p_error)

End Function

''' <summary>   Enqueues a VBA Error object optionally updating the error source and timeout
''' and adding a suffix to the description. . </summary>
''' <param name="a_source">      [Optional, Variant, String] The error source. </param>
''' <param name="a_suffix">      [Optional, Variant, String] A suffix to append to the message description. </param>
''' <param name="a_timestamp">   [Optional, Variant, double] The error timestamp. </param>
''' <returns>    [<see cref="UserDefinedError"/>]. </returns>
Public Function EnqueueErrorObject( _
        Optional ByVal a_source As Variant, _
        Optional ByVal a_suffix As Variant, _
        Optional ByVal a_timestamp As Variant) As UserDefinedError

    Dim p_error As New UserDefinedError
    
    If 0 <> Err.Number Then
    
        If VBA.IsMissing(a_timestamp) Then
            If VBA.IsMissing(a_suffix) Then
                If VBA.IsMissing(a_source) Then
                    p_error.FromAppendedErrorObject
                Else
                    p_error.FromAppendedErrorObject a_source
                End If
            ElseIf VBA.IsMissing(a_source) Then
                p_error.FromAppendedErrorObject a_suffix:=a_suffix
            Else
                p_error.FromAppendedErrorObject a_source, a_suffix
            End If
        ElseIf VBA.IsMissing(a_suffix) Then
            If VBA.IsMissing(a_source) Then
                p_error.FromAppendedErrorObject a_timestamp:=a_timestamp
            Else
                p_error.FromAppendedErrorObject a_source:=a_source, a_timestamp:=a_timestamp
            End If
        ElseIf VBA.IsMissing(a_source) Then
            p_error.FromAppendedErrorObject a_suffix:=a_suffix, a_timestamp:=a_timestamp
        Else
            p_error.FromAppendedErrorObject a_source, a_suffix, a_timestamp
        End If
        
        Set EnqueueErrorObject = Me.EnqueueUserDefinedError(p_error)

    Else
    
        Set EnqueueErrorObject = p_error
    
    End If

End Function

''' <summary>   Raises a user defined error optionally updating the error source and timeout
''' and adding a suffix to the description. </summary>
''' <param name="a_error">       [<see cref="UserDefinedError"/>]. </param>
''' <param name="a_source">      [String] The error source. </param>
''' <param name="a_suffix">      [Optional, Variant, String] A suffix to append to the message description. </param>
''' <param name="a_timestamp">   [Optional, Variant, Double] The error timestamp. </param>
Public Sub RaiseError(ByVal a_error As UserDefinedError, ByVal a_source As String, _
        Optional ByVal a_suffix As Variant, _
        Optional ByVal a_timestamp As Variant)

    Dim p_error As UserDefinedError
    
    If VBA.IsMissing(a_timestamp) Then
        If VBA.IsMissing(a_suffix) Then
            Set p_error = Me.EnqueueError(a_error, a_source)
        Else
            Set p_error = Me.EnqueueError(a_error, a_source, a_suffix)
        End If
    ElseIf VBA.IsMissing(a_suffix) Then
        Set p_error = Me.EnqueueError(a_error, a_source, a_timestamp:=a_timestamp)
    Else
        Set p_error = Me.EnqueueError(a_error, a_source, a_suffix, a_timestamp)
    End If
    
    VBA.Err.Raise p_error.Number, p_error.Source, p_error.Description
    
End Sub

''' <summary>   Clears the VBA Error object, and user defined error archive and queue. </summary>
Public Sub ClearErrorState()

    ' remove the last queued error.
    Set This.LastQueuedError = Nothing

    ' clear the error archive.
    This.ErrorsArchiveStack.Clear

    ' clear the error queue.
    This.ErrorsQueue.Clear

    ' clear the error object
    Err.Clear
    
End Sub

''' <summary>   Tests raising a user defined error message for checking the automation error. </summary>
Public Sub TestRaisingNullArgumentError(a_object As Object)

    If a_object Is Nothing Then
    
        UserDefinedErrors.RaiseError cc_isr_Core_IO.UserDefinedErrors.NullArgumentError, _
                    ThisWorkbook.VBProject.Name & "." & TypeName(Me) & "TestRaisingNullArgumentError", _
                    " Argument name: 'a_value'.", cc_isr_Core_IO.CoreExtensions.DaysNow()
                    
        Debug.Print ErrorMessageBuilder.BuildStandardErrorMessage()
    
    End If

End Sub
