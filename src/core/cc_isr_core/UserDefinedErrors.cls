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

Private m_errorsDix As Dictionary

Private m_winsockStartingError As UserDefinedError

Private m_winsockCleanupError As UserDefinedError

Private m_socketCreationError As UserDefinedError

Private m_socketConnectionError As UserDefinedError

Private m_socketDisconnectionError As UserDefinedError

Private m_socketOptionSettingError As UserDefinedError

Private m_socketSendError As UserDefinedError

Private m_socketReceiveError As UserDefinedError

Private m_socketBindError As UserDefinedError

Private m_socketListenError As UserDefinedError

Private m_socketAcceptError As UserDefinedError

Private m_socketSetStatusError As UserDefinedError

Private m_invalidArgumentError As UserDefinedError

Private m_invalidOperationError As UserDefinedError

Private m_timeoutError As UserDefinedError


''' <summary>   Initializes the user defined errors. </summary>
Private Sub Class_Initialize()

    Set m_errorsDix = New Dictionary
    Dim p_errorCode As Long
    Dim p_dummyError As UserDefinedError
    Set p_dummyError = New UserDefinedError
    p_errorCode = VBA.vbObjectError + p_dummyError.UserErrorNumberBase
    
    p_errorCode = p_errorCode + 1
    Set m_winsockStartingError = New UserDefinedError
    m_winsockStartingError.InitiateProperties p_errorCode, "WinsockStartingError", _
        "Failed starting Winsock using Winsock WSAStartup()."
    AddUserDefinedError Me.WinsockStartingError
    
    p_errorCode = p_errorCode + 1
    Set m_winsockCleanupError = New UserDefinedError
    m_winsockCleanupError.InitiateProperties p_errorCode, "WinsockCleanupError", _
        "Failed disposing Winsock using Winsock WSACleanup()."
    AddUserDefinedError Me.WinsockCleanupError
    
    p_errorCode = p_errorCode + 1
    Set m_socketCreationError = New UserDefinedError
    m_socketCreationError.InitiateProperties p_errorCode, "SocketCreationError", _
        "Failed creating socket using Winsock socket()."
    AddUserDefinedError Me.SocketCreationError
    
    p_errorCode = p_errorCode + 1
    Set m_socketConnectionError = New UserDefinedError
    m_socketConnectionError.InitiateProperties p_errorCode, "SocketConnectionError", _
        "Failed connecting socket using Winsock connect()."
    AddUserDefinedError Me.SocketConnectionError
    
    p_errorCode = p_errorCode + 1
    Set m_socketDisconnectionError = New UserDefinedError
    m_socketDisconnectionError.InitiateProperties p_errorCode, "SocketDisconnectionError", _
        "Failed disconnecting socket using Winsock closesocket()."
    AddUserDefinedError Me.SocketDisconnectionError
    
    p_errorCode = p_errorCode + 1
    Set m_socketOptionSettingError = New UserDefinedError
    m_socketOptionSettingError.InitiateProperties p_errorCode, "SocketOptionSettingError", _
        "Failed setting socket options using Winsock setsockopt()."
    AddUserDefinedError Me.SocketOptionSettingError
    
    p_errorCode = p_errorCode + 1
    Set m_socketSendError = New UserDefinedError
    m_socketSendError.InitiateProperties p_errorCode, "SocketSendError", _
        "Failed send a socket message using Winsock send()."
    AddUserDefinedError Me.SocketSendError

    p_errorCode = p_errorCode + 1
    Set m_socketReceiveError = New UserDefinedError
    m_socketReceiveError.InitiateProperties p_errorCode, "SocketReceiveError", _
        "Failed receiving a socket message using Winsock recv()."
    AddUserDefinedError Me.SocketReceiveError

    p_errorCode = p_errorCode + 1
    Set m_socketBindError = New UserDefinedError
    m_socketBindError.InitiateProperties p_errorCode, "SocketBindError", _
        "Failed binding a socket using Winsock bind()."
    AddUserDefinedError Me.SocketBindError

    p_errorCode = p_errorCode + 1
    Set m_socketListenError = New UserDefinedError
    m_socketListenError.InitiateProperties p_errorCode, "SocketListenError", _
        "Failed stating a socket listening using Winsock listen()."
    AddUserDefinedError Me.SocketListenError

    p_errorCode = p_errorCode + 1
    Set m_socketAcceptError = New UserDefinedError
    m_socketAcceptError.InitiateProperties p_errorCode, "SocketAcceptError", _
        "Failed permitting an incoming connection attempt on a socket using Winsock accept()."
    AddUserDefinedError Me.SocketAcceptError

    p_errorCode = p_errorCode + 1
    Set m_socketSetStatusError = New UserDefinedError
    m_socketSetStatusError.InitiateProperties p_errorCode, "SocketSetStatusError", _
        "Failed determining the socket status using Winsock select()."
    AddUserDefinedError Me.SocketSetStatusError

    p_errorCode = p_errorCode + 1
    Set m_invalidArgumentError = New UserDefinedError
    m_invalidArgumentError.InitiateProperties p_errorCode, "InvalidArgumentError", _
        "Invalid argument error."
    AddUserDefinedError Me.InvalidArgumentError

    p_errorCode = p_errorCode + 1
    Set m_invalidOperationError = New UserDefinedError
    m_invalidOperationError.InitiateProperties p_errorCode, "InvalidOperationError", _
        "Invalid operation error."
    AddUserDefinedError Me.InvalidOperationError
    
    p_errorCode = p_errorCode + 1
    Set m_timeoutError = New UserDefinedError
    m_timeoutError.InitiateProperties p_errorCode, "TimeoutError", _
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
    
    m_errorsDix.Add a_userError.Code, a_userError

End Sub

''' <summary>   Returns true if the specified error exists. </summary>
''' <para name"a_userError">   The error of type <see cref="UserDefinedError"/>   </param>
Public Function UserDefinedErrorExists(ByVal a_userError As UserDefinedError) As Boolean
    
    UserDefinedErrorExists = m_errorsDix.Exists(a_userError.Code)

End Function

''' <summary>   Returns true if the specified error exists. </summary>
''' <para name"a_userErrorCode">   [Long] The error code </param>
Public Function UserDefinedErrorCodeExists(ByVal a_userErrorCode As Long) As Boolean
    
    UserDefinedErrorCodeExists = m_errorsDix.Exists(a_userErrorCode)

End Function

''' <summary>   Returns a user defined error. </summary>
''' <param name="a_key">   The user defined error code. </param>
Public Function GetUserDefinedError(ByVal a_key As Long) As UserDefinedError
    
    If m_errorsDix.Exists(a_key) Then
        Set GetUserDefinedError = m_errorsDix.item(a_key)
    Else
        Set GetUserDefinedError = New UserDefinedError
        GetUserDefinedError.InitiateProperties a_key, "UnknwonUserDefinedError", "Unknown error."
    End If

End Function

''' <summary>   Get the error that occurred starting Winsock. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get WinsockStartingError() As UserDefinedError
    
    Set WinsockStartingError = m_winsockStartingError

End Property

''' <summary>   Get the error that occurred ending Winsock. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get WinsockCleanupError() As UserDefinedError
    
    Set WinsockCleanupError = m_winsockCleanupError

End Property

''' <summary>   Get the error that occurred creating a socket. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketCreationError() As UserDefinedError
    
    Set SocketCreationError = m_socketCreationError

End Property

''' <summary>   Get the error that occurred connecting a socket. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketConnectionError() As UserDefinedError
    
    Set SocketConnectionError = m_socketConnectionError

End Property

''' <summary>   Get the error that occurred disconnecting a socket. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketDisconnectionError() As UserDefinedError
    
    Set SocketDisconnectionError = m_socketDisconnectionError

End Property

''' <summary>   Get the error that occurred setting a socket option. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketOptionSettingError() As UserDefinedError
    
    Set SocketOptionSettingError = m_socketOptionSettingError

End Property

''' <summary>   Get the error that occurred sending a socket message. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketSendError() As UserDefinedError
    
    Set SocketSendError = m_socketSendError

End Property

''' <summary>   Get the error that occurred receiving a socket message. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketReceiveError() As UserDefinedError
    
    Set SocketReceiveError = m_socketReceiveError

End Property

''' <summary>   Get the error that occurred binding a socket to an address. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketBindError() As UserDefinedError
    
    Set SocketBindError = m_socketBindError

End Property

''' <summary>   Get the error that occurred starting to listen on a socket. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketListenError() As UserDefinedError
    
    Set SocketListenError = m_socketListenError

End Property

''' <summary>   Get the error that occurred accepting an incoming connection. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketAcceptError() As UserDefinedError
    
    Set SocketAcceptError = m_socketAcceptError

End Property

''' <summary>   Get the error that occurred determining the status of a socket set. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get SocketSetStatusError() As UserDefinedError
    
    Set SocketSetStatusError = m_socketSetStatusError

End Property

''' <summary>   Get the error that occurred when the method argument is invalid. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get InvalidArgumentError() As UserDefinedError
    
    Set InvalidArgumentError = m_invalidArgumentError

End Property

''' <summary>   Get the error that occurred when the method operation is invalid. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get InvalidOperationError() As UserDefinedError
    
    Set InvalidOperationError = m_invalidOperationError

End Property

''' <summary>   Get the error that occurred when the method operation times out. </summary>
''' <value>   <see cref="UserDefinedError"/>. </value>
Public Property Get TimeoutError() As UserDefinedError
    
    Set TimeoutError = m_timeoutError

End Property

