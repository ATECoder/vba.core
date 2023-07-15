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

''' <summary>   Builds a an error message from the VBA <see cref="Err"/>. </summary>
Public Function FormattedErr() As String
    
    Dim p_errNoDetails As String
    p_errNoDetails = " (" & (Err.Number - VBA.vbObjectError) & " / 0x" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")"
    FormattedErr = "[Error Number=" & Err.Number & p_errNoDetails & ", Source=" & Err.Source & ", Description=" & Err.Description & "]"

End Function

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

''' <summary>   Sets the Err object source string to project.module.procedure. </summary>
''' <param name="a_errProcedureName">   [String] Specifies the name of the procedure. </param>
''' <param name="a_errModuleName">      [String] Specifies the module name. </param>
''' <param name="a_errProjectName">     [Optional, String] Specifies the project name; otherwise the project
'''                                     name of the active workbook is used. </param>
Public Sub SetErrSource(ByVal a_errProcedureName As String, ByVal a_errModuleName As String, _
        Optional ByVal a_errProjectName As String = vbNullString)

    ' this procedure must not trap errors because it must
    ' not alter the error object.
      
    ' thus we assume that this code is robust and will
    ' not cause errors.
    
    If vbNullString = a_errProjectName Then
        a_errProjectName = ActiveWorkbook.VBProject.name
    End If
    
    ' get the current source string
    
    Dim p_errorSource As String: p_errorSource = Err.Source
    
    ' build the error source.
    
    p_errorSource = a_errProjectName & "." & a_errModuleName & "." & a_errProcedureName
  
    ' Update the Err.Source
    
    Err.Source = p_errorSource
 
End Sub

''' <summary>   Builds a standard error message. </summary>
''' <param name="a_displayWarning">         [Optional, Boolean, false] True
'''                                         to display a warning rather than
'''                                         an error message. </param>
''' <param name="a_descriptionDelimiter">   [Optional, String, ': '] Specify
'''                                         the delimiter preceding the description. </param>
''' <returns>   A Standard error message string in the form: <para>
''' Error # (0x#) occurred in <c>Source</c>: Description </para><para>
''' or  </para><para>
''' Warning # (0x#) occurred in <c>Source</c>: Description  </para>
''' </returns>
Public Function BuildStandardErrorMessage(Optional ByVal a_displayWarning As Boolean = False, _
                Optional a_descriptionDelimiter As String = ": ") As String

    Dim p_builder As StringBuilder: Set p_builder = New StringBuilder
  
    ' check if we have an error
    If Err.Number <> 0 Then
    
        p_builder.Append IIf(a_displayWarning, "Warning ", "Error ")
        p_builder.Append Format$(Err.Number)
        
        Dim p_errNumber As Long: p_errNumber = Err.Number - vbObjectError
        
        If Abs(p_errNumber) < &HFFFF& Then
            p_errNumber = p_errNumber - &H200
            p_builder.Append " (+0x"
        Else
            p_errNumber = Err.Number
            p_builder.Append " (0x"
        End If
        p_builder.Append Hex$(p_errNumber)
        p_builder.Append ") "
        p_builder.Append "occurred in "
        p_builder.Append Err.Source
        p_builder.Append a_descriptionDelimiter
        p_builder.Append Err.Description
        
    End If

    BuildStandardErrorMessage = p_builder.ToString

End Function


