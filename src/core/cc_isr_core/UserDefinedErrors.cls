VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserDefinedErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' UserDefinedErrors.cls
'
' Dependencies:
'
' UserDefinedError.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type MyType
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
    InvalidOperationError As UserDefinedError
    TimeoutError As UserDefinedError
End Type

Private This As MyType

''' <summary>   Initializes the user defined errors. </summary>
Private Sub Class_Initialize()

    Set This.ErrorsDix = New Dictionary
    Dim p_errorCode As Long
    Dim p_dummyError As UserDefinedError
    Set p_dummyError = New UserDefinedError
    p_errorCode = VBA.vbObjectError + p_dummyError.UserErrorNumberBase
    
    p_errorCode = p_errorCode + 1
    Set This.WinsockStartingError = New UserDefinedError
    This.WinsockStartingError.InitiateProperties p_errorCode, "WinsockStartingError", _
        "Failed starting Winsock using Winsock WSAStartup()."
    AddUserDefinedError Me.WinsockStartingError
    
    p_errorCode = p_errorCode + 1
    Set This.WinsockCleanupError = New UserDefinedError
    This.WinsockCleanupError.InitiateProperties p_errorCode, "WinsockCleanupError", _
        "Failed disposing Winsock using Winsock WSACleanup()."
    AddUserDefinedError Me.WinsockCleanupError
    
    p_errorCode = p_errorCode + 1
    Set This.SocketCreationError = New UserDefinedError
    This.SocketCreationError.InitiateProperties p_errorCode, "SocketCreationError", _
        "Failed creating socket using Winsock socket()."
    AddUserDefinedError Me.SocketCreationError
    
    p_errorCode = p_errorCode + 1
    Set This.SocketConnectionError = New UserDefinedError
    This.SocketConnectionError.InitiateProperties p_errorCode, "SocketConnectionError", _
        "Failed connecting socket using Winsock connect()."
    AddUserDefinedError Me.SocketConnectionError
    
    p_errorCode = p_errorCode + 1
    Set This.SocketDisconnectionError = New UserDefinedError
    This.SocketDisconnectionError.InitiateProperties p_errorCode, "SocketDisconnectionError", _
        "Failed disconnecting socket using Winsock closesocket()."
    AddUserDefinedError Me.SocketDisconnectionError
    
    p_errorCode = p_errorCode + 1
    Set This.SocketOptionSettingError = New UserDefinedError
    This.SocketOptionSettingError.InitiateProperties p_errorCode, "SocketOptionSettingError", _
        "Failed setting socket options using Winsock setsockopt()."
    AddUserDefinedError Me.SocketOptionSettingError
    
    p_errorCode = p_errorCode + 1
    Set This.SocketSendError = New UserDefinedError
    This.SocketSendError.InitiateProperties p_errorCode, "SocketSendError", _
        "Failed send a socket message using Winsock send()."
    AddUserDefinedError Me.SocketSendError

    p_errorCode = p_errorCode + 1
    Set This.SocketReceiveError = New UserDefinedError
    This.SocketReceiveError.InitiateProperties p_errorCode, "SocketReceiveError", _
        "Failed receiving a socket message using Winsock recv()."
    AddUserDefinedError Me.SocketReceiveError

    p_errorCode = p_errorCode + 1
    Set This.SocketBindError = New UserDefinedError
    This.SocketBindError.InitiateProperties p_errorCode, "SocketBindError", _
        "Failed binding a socket using Winsock bind()."
    AddUserDefinedError Me.SocketBindError

    p_errorCode = p_errorCode + 1
    Set This.SocketListenError = New UserDefinedError
    This.SocketListenError.InitiateProperties p_errorCode, "SocketListenError", _
        "Failed stating a socket listening using Winsock listen()."
    AddUserDefinedError Me.SocketListenError

    p_errorCode = p_errorCode + 1
    Set This.SocketAcceptError = New UserDefinedError
    This.SocketAcceptError.InitiateProperties p_errorCode, "SocketAcceptError", _
        "Failed permitting an incoming connection attempt on a socket using Winsock accept()."
    AddUserDefinedError Me.SocketAcceptError

    p_errorCode = p_errorCode + 1
    Set This.SocketSetStatusError = New UserDefinedError
    This.SocketSetStatusError.InitiateProperties p_errorCode, "SocketSetStatusError", _
        "Failed determining the socket status using Winsock select()."
    AddUserDefinedError Me.SocketSetStatusError

    p_errorCode = p_errorCode + 1
    Set This.NullArgumentError = New UserDefinedError
    This.NullArgumentError.InitiateProperties p_errorCode, "nullArgumentError", _
        "null argument error."
    AddUserDefinedError Me.NullArgumentError

    p_errorCode = p_errorCode + 1
    Set This.InvalidArgumentError = New UserDefinedError
    This.InvalidArgumentError.InitiateProperties p_errorCode, "InvalidArgumentError", _
        "Invalid argument error."
    AddUserDefinedError Me.InvalidArgumentError

    p_errorCode = p_errorCode + 1
    Set This.InvalidOperationError = New UserDefinedError
    This.InvalidOperationError.InitiateProperties p_errorCode, "InvalidOperationError", _
        "Invalid operation error."
    AddUserDefinedError Me.InvalidOperationError
    
    p_errorCode = p_errorCode + 1
    Set This.TimeoutError = New UserDefinedError
    This.TimeoutError.InitiateProperties p_errorCode, "TimeoutError", _
        "Timeout error."
    AddUserDefinedError Me.TimeoutError

End Sub

''' <summary>   Initializes the error codes and descriptions. </summary>
Public Sub Initialize()
    Class_Initialize
End Sub

''' <summary>   Adds a user defined error to the error dictionary. </summary>
''' <para name"a_userError">   The error of type <see cref="UserDefinedError"/>   </param>
Private Sub AddUserDefinedError(ByVal a_userError As UserDefinedError)
    
    This.ErrorsDix.Add a_userError.Code, a_userError

End Sub

''' <summary>   Returns true if the specified error exists. </summary>
''' <para name"a_userError">   The error of type <see cref="UserDefinedError"/>   </param>
Public Function UserDefinedErrorExists(ByVal a_userError As UserDefinedError) As Boolean
    
    UserDefinedErrorExists = This.ErrorsDix.Exists(a_userError.Code)

End Function

''' <summary>   Returns true if the specified error exists. </summary>
''' <para name"a_userErrorCode">   [Long] The error code </param>
Public Function UserDefinedErrorCodeExists(ByVal a_userErrorCode As Long) As Boolean
    
    UserDefinedErrorCodeExists = This.ErrorsDix.Exists(a_userErrorCode)

End Function

''' <summary>   Returns a user defined error. </summary>
''' <param name="a_key">   The user defined error code. </param>
Public Function GetUserDefinedError(ByVal a_key As Long) As UserDefinedError
    
    If This.ErrorsDix.Exists(a_key) Then
        Set GetUserDefinedError = This.ErrorsDix.item(a_key)
    Else
        Set GetUserDefinedError = New UserDefinedError
        GetUserDefinedError.InitiateProperties a_key, "UnknwonUserDefinedError", "Unknown error."
    End If

End Function

''' <summary>   Get the error that occurred starting Winsock. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get WinsockStartingError() As UserDefinedError
    
    Set WinsockStartingError = This.WinsockStartingError

End Property

''' <summary>   Get the error that occurred ending Winsock. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get WinsockCleanupError() As UserDefinedError
    
    Set WinsockCleanupError = This.WinsockCleanupError

End Property

''' <summary>   Get the error that occurred creating a socket. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketCreationError() As UserDefinedError
    
    Set SocketCreationError = This.SocketCreationError

End Property

''' <summary>   Get the error that occurred connecting a socket. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketConnectionError() As UserDefinedError
    
    Set SocketConnectionError = This.SocketConnectionError

End Property

''' <summary>   Get the error that occurred disconnecting a socket. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketDisconnectionError() As UserDefinedError
    
    Set SocketDisconnectionError = This.SocketDisconnectionError

End Property

''' <summary>   Get the error that occurred setting a socket option. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketOptionSettingError() As UserDefinedError
    
    Set SocketOptionSettingError = This.SocketOptionSettingError

End Property

''' <summary>   Get the error that occurred sending a socket message. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketSendError() As UserDefinedError
    
    Set SocketSendError = This.SocketSendError

End Property

''' <summary>   Get the error that occurred receiving a socket message. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketReceiveError() As UserDefinedError
    
    Set SocketReceiveError = This.SocketReceiveError

End Property

''' <summary>   Get the error that occurred binding a socket to an address. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketBindError() As UserDefinedError
    
    Set SocketBindError = This.SocketBindError

End Property

''' <summary>   Get the error that occurred starting to listen on a socket. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketListenError() As UserDefinedError
    
    Set SocketListenError = This.SocketListenError

End Property

''' <summary>   Get the error that occurred accepting an incoming connection. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketAcceptError() As UserDefinedError
    
    Set SocketAcceptError = This.SocketAcceptError

End Property

''' <summary>   Get the error that occurred determining the status of a socket set. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketSetStatusError() As UserDefinedError
    
    Set SocketSetStatusError = This.SocketSetStatusError

End Property

''' <summary>   Get the error that occurred when the method argument is null. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get NullArgumentError() As UserDefinedError
    
    Set NullArgumentError = This.NullArgumentError

End Property

''' <summary>   Get the error that occurred when the method argument is invalid. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get InvalidArgumentError() As UserDefinedError
    
    Set InvalidArgumentError = This.InvalidArgumentError

End Property

''' <summary>   Get the error that occurred when the method operation is invalid. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get InvalidOperationError() As UserDefinedError
    
    Set InvalidOperationError = This.InvalidOperationError

End Property

''' <summary>   Get the error that occurred when the method operation times out. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get TimeoutError() As UserDefinedError
    
    Set TimeoutError = This.TimeoutError

End Property

''' <summary>   Raises a user defined error. </summary>
''' <param name="a_error">    [<see cref="UserDefinedError"/>]. </param>
''' <param name="a_source">   [String] The error source. </param>
''' <param name="a_suffix">   [String, Optional, Null] A suffix to append to the message description. </param>
Public Sub RaiseError(ByVal a_error As UserDefinedError, ByVal a_source As String, _
        Optional a_suffix As String = VBA.vbNullString)

    VBA.Err.Raise a_error.Code, a_source, a_error.Description & a_suffix

End Sub
