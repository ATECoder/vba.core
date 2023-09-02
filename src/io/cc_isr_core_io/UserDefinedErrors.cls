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
    NextUserDefinedErrorNumber As Long
    ErrorsQueue As Queue
    ErrorsArchiveStack As Stack
End Type

Private This As this_

''' <summary>   Initializes the user defined errors. </summary>
Private Sub Class_Initialize()

    Set This.ErrorsQueue = Factory.CreateQueue(100)
    Set This.ErrorsArchiveStack = Factory.CreateStack(100)
    Set This.ErrorsDix = New Dictionary
    Dim p_dummyError As UserDefinedError
    Set p_dummyError = New UserDefinedError
    This.NextUserDefinedErrorNumber = VBA.vbObjectError + p_dummyError.UserErrorNumberBase + 1
    
    Set This.WinsockStartingError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "WinsockStartingError", _
                                        "Failed starting Winsock using Winsock WSAStartup(). ")
    AddUserDefinedError Me.WinsockStartingError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1
    
    Set This.WinsockCleanupError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "WinsockCleanupError", _
                                        "Failed disposing Winsock using Winsock WSACleanup(). ")
    AddUserDefinedError Me.WinsockCleanupError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1
    
    Set This.SocketCreationError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketCreationError", _
                                        "Failed creating socket using Winsock socket(). ")
    AddUserDefinedError Me.SocketCreationError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1
    
    Set This.SocketConnectionError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketConnectionError", _
                                        "Failed connecting socket using Winsock connect(). ")
    AddUserDefinedError Me.SocketConnectionError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1
    
    Set This.SocketDisconnectionError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketDisconnectionError", _
                                            "Failed disconnecting socket using Winsock closesocket(). ")
    AddUserDefinedError Me.SocketDisconnectionError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1
    
    Set This.SocketOptionSettingError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketOptionSettingError", _
                                            "Failed setting socket options using Winsock setsockopt(). ")
    AddUserDefinedError Me.SocketOptionSettingError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1
    
    Set This.SocketSendError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketSendError", _
                                    "Failed send a socket message using Winsock send(). ")
    AddUserDefinedError Me.SocketSendError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

    Set This.SocketReceiveError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketReceiveError", _
                                    "Failed receiving a socket message using Winsock recv(). ")
    AddUserDefinedError Me.SocketReceiveError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

    Set This.SocketBindError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketBindError", _
                                    "Failed binding a socket using Winsock bind(). ")
    AddUserDefinedError Me.SocketBindError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

    Set This.SocketListenError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketListenError", _
                                    "Failed stating a socket listening using Winsock listen(). ")
    AddUserDefinedError Me.SocketListenError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

    Set This.SocketAcceptError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketAcceptError", _
                                    "Failed permitting an incoming connection attempt on a socket using Winsock accept(). ")
    AddUserDefinedError Me.SocketAcceptError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

    Set This.SocketSetStatusError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "SocketSetStatusError", _
                                        "Failed determining the socket status using Winsock select(). ")
    AddUserDefinedError Me.SocketSetStatusError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

    Set This.NullArgumentError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "nullArgumentError", _
                                    "Null argument error. ")
    AddUserDefinedError Me.NullArgumentError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

    Set This.InvalidArgumentError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "InvalidArgumentError", _
                                        "Invalid argument error. ")
    AddUserDefinedError Me.InvalidArgumentError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

    Set This.InvalidFormatStringError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "InvalidFormatStringError", _
                                            "Invalid format string. ")
    AddUserDefinedError Me.InvalidFormatStringError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

    Set This.InvalidOperationError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "InvalidOperationError", _
                                        "Invalid operation error. ")
    AddUserDefinedError Me.InvalidOperationError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1
    
    Set This.TimeoutError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "TimeoutError", _
                                "Timeout error. ")
    AddUserDefinedError Me.TimeoutError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

    Set This.FormatError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "FormatError", _
        "The number indicating an argument to format is less than zero or greater than or equal to the length of the array of arguments. ")
    AddUserDefinedError Me.FormatError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

    Set This.Ieee488DeviceError = Factory.NewUserDefinedError.Initialize(This.NextUserDefinedErrorNumber, "Ieee488DeviceError", _
        "IEEE488 Device error(s) occurred. ")
    AddUserDefinedError Me.Ieee488DeviceError
    This.NextUserDefinedErrorNumber = This.NextUserDefinedErrorNumber + 1

End Sub

''' <summary>   Initializes the error Numbers and descriptions. </summary>
''' <returns>   [<see cref="UserDefinedErrors"/>]. </returns>
Public Function Initialize() As UserDefinedErrors
    Class_Initialize
    Set Initialize = Me
End Function

''' <summary>   Adds a user defined error to the error dictionary. </summary>
''' <para name"a_userError">   The error of type <see cref="UserDefinedError"/>   </param>
Private Sub AddUserDefinedError(ByVal a_userError As UserDefinedError)
    
    ' the error source and timestamp must be provided when raising the error.
    
    a_userError.Source = VBA.vbNullString
    a_userError.Timestamp = 0
    
    This.ErrorsDix.Add a_userError.Number, a_userError

End Sub

''' <summary>   Returns true if the specified error exists. </summary>
''' <para name"a_userError">   The error of type <see cref="UserDefinedError"/>   </param>
Public Function UserDefinedErrorExists(ByVal a_userError As UserDefinedError) As Boolean
    
    UserDefinedErrorExists = This.ErrorsDix.Exists(a_userError.Number)

End Function

''' <summary>   Returns true if the specified error exists. </summary>
''' <para name"a_userErrorNumber">   [Long] The error Number </param>
Public Function UserDefinedErrorNumberExists(ByVal a_userErrorNumber As Long) As Boolean
    
    UserDefinedErrorNumberExists = This.ErrorsDix.Exists(a_userErrorNumber)

End Function

''' <summary>   Returns a user defined error. </summary>
''' <param name="a_key">   [Long] The user defined error Number. </param>
''' <returns>   An existing or a new 'unknown' <see cref="UserDefinedError"/>. </returns>
Public Function GetUserDefinedError(ByVal a_key As Long) As UserDefinedError
    
    If This.ErrorsDix.Exists(a_key) Then
        Set GetUserDefinedError = This.ErrorsDix.Item(a_key)
    Else
        Set GetUserDefinedError = Factory.NewUserDefinedError.Initialize(a_key, _
                                        "UnknownUserDefinedError", "Unknown error. ")
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
''' <value>   [<see cref="cc_isr_Core_IO.Queue"/>]. </value>
Public Property Get ErrorsQueue() As Queue
    Set ErrorsQueue = This.ErrorsQueue
End Property

''' <summary>   Gets the stack of archived errors. </summary>
''' <value>   [<see cref="cc_isr_Core_IO.Stack"/>]. </value>
Public Property Get ErrorsArchiveStack() As Stack
    Set ErrorsArchiveStack = This.ErrorsArchiveStack
End Property

''' <summary>   Clones a user defined error optionally updating the error source and timeout
''' and adding a suffix to the description. </summary>
''' <param name="a_error">       [<see cref="UserDefinedError"/>]. </param>
''' <param name="a_source">      [String] The error source. </param>
''' <param name="a_suffix">      [String, Optional, Null] A suffix to append to the message description. </param>
''' <param name="a_timestamp">   [Variant double] The error timestamp. </param>
''' <returns>    [<see cref="UserDefinedError"/>]. </returns>
Public Function CloneError(ByVal a_error As UserDefinedError, _
        Optional ByVal a_source As String, _
        Optional ByVal a_suffix As String, _
        Optional ByVal a_timestamp As Variant) As UserDefinedError

    ' Append new information tot he standard user defined error
    
    Dim p_description As String
    p_description = a_error.Description
    If Not VBA.IsMissing(a_suffix) Then _
        If a_suffix <> VBA.vbNullString Then _
            p_description = p_description & a_suffix
    
    Dim p_error As UserDefinedError
    Set p_error = Factory.NewUserDefinedError.Initialize(a_error.Number, a_error.Name, _
        p_description)
    
    If Not VBA.IsMissing(a_source) Then _
        If a_source <> VBA.vbNullString Then _
            p_error.Source = a_source
    
    If Not VBA.IsMissing(a_timestamp) Then _
        p_error.Timestamp = a_timestamp
    
    Set CloneError = p_error

End Function


''' <summary>   Enqueues a user defined error optionally updating the error source and timeout
''' and adding a suffix to the description. </summary>
''' <param name="a_error">       [<see cref="UserDefinedError"/>]. </param>
''' <param name="a_source">      [String] The error source. </param>
''' <param name="a_suffix">      [String, Optional, Null] A suffix to append to the message description. </param>
''' <param name="a_timestamp">   [Variant double] The error timestamp. </param>
''' <returns>    [<see cref="UserDefinedError"/>]. </returns>
Public Function EnqueueError(ByVal a_error As UserDefinedError, _
        Optional ByVal a_source As String, _
        Optional ByVal a_suffix As String, _
        Optional ByVal a_timestamp As Variant) As UserDefinedError

    Dim p_error As UserDefinedError
    
    If VBA.IsMissing(a_timestamp) Then
        If VBA.IsMissing(a_suffix) Then
            If VBA.IsMissing(a_source) Then
                Set p_error = Me.CloneError(a_error)
            Else
                Set p_error = Me.CloneError(a_error, a_source)
            End If
        ElseIf VBA.IsMissing(a_source) Then
            Set p_error = Me.CloneError(a_error, a_suffix:=a_suffix)
        Else
            Set p_error = Me.CloneError(a_error, a_source, a_suffix)
        End If
    ElseIf VBA.IsMissing(a_suffix) Then
        If VBA.IsMissing(a_source) Then
            Set p_error = Me.CloneError(a_error, a_timestamp:=a_timestamp)
        Else
            Set p_error = Me.CloneError(a_error, a_source:=a_source, a_timestamp:=a_timestamp)
        End If
    ElseIf VBA.IsMissing(a_source) Then
        Set p_error = Me.CloneError(a_error, a_suffix:=a_suffix, a_timestamp:=a_timestamp)
    Else
        Set p_error = Me.CloneError(a_error, a_source, a_suffix, a_timestamp)
    End If
    
    This.ErrorsQueue.Enqueue p_error
    
    Set EnqueueError = p_error

End Function

''' <summary>   Raises a user defined error optionally updating the error source and timeout
''' and adding a suffix to the description. </summary>
''' <param name="a_error">       [<see cref="UserDefinedError"/>]. </param>
''' <param name="a_source">      [String] The error source. </param>
''' <param name="a_suffix">      [String, Optional, Null] A suffix to append to the message description. </param>
''' <param name="a_timestamp">   [Variant double] The error timestamp. </param>
Public Sub RaiseError(ByVal a_error As UserDefinedError, ByVal a_source As String, _
        Optional ByVal a_suffix As String, _
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

    ' clear the error archive.
    Me.ErrorsArchiveStack.Clear

    ' clear the error queue.
    Me.ErrorsQueue.Clear

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
