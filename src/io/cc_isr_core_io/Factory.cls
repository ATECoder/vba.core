VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A public creatable singleton class for creating instances of this workbook
'''             classes to share with external workbooks. </summary>
''' <remarks>   Dependencies: Stopwatch.cls, UserDefinedError.cls. <remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   A public creatable singleton class for creating instances of this workbook
''' classes to share with external workbooks. </summary>
''' <remarks>
''' <see href="https://learn.Microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/set-up-vb-project-using-class"/>
''' </remarks>
Private Sub Class_Initialize()
End Sub

''' <summary>   Instantiates a new <see cref="Stopwatch"/> object. </summary>
''' <returns>   [<see cref="Stopwatch"/>]. </returns>
Public Function NewStopwatch() As cc_isr_Core_IO.Stopwatch
    Set NewStopwatch = New cc_isr_Core_IO.Stopwatch
End Function

''' <summary>   Instantiates a new <see cref="UserDefinedError"/> object. </summary>
''' <returns>   [<see cref="UserDefinedError"/>]. </returns>
Public Function NewUserDefinedError() As UserDefinedError
    Set NewUserDefinedError = New UserDefinedError
End Function

''' <summary>   Instantiates a new <see cref="Queue"/> object. </summary>
''' <returns>   [<see cref="Queue"/>]. </returns>
Public Function NewQueue() As Queue
    Set NewQueue = New Queue
End Function

''' <summary>   Instantiates a new <see cref="Queue"/> object with the specified capacity. </summary>
''' <param name="a_capacity">   [Integer] The capacity of the queue. </param>
''' <returns>   [<see cref="Queue"/>]. </returns>
Public Function CreateQueue(ByVal a_capacity As Integer) As Queue
    Set CreateQueue = Me.NewQueue.Initialize(a_capacity)
End Function

''' <summary>   Instantiates a new <see cref="Stack"/> object. </summary>
''' <returns>   [<see cref="Stack"/>]. </returns>
Public Function NewStack() As Stack
    Set NewStack = New Stack
End Function

''' <summary>   Instantiates a new <see cref="Stack"/> object with the specified capacity. </summary>
''' <param name="a_capacity">   [Integer] The capacity of the Stack. </param>
''' <returns>   [<see cref="Stack"/>]. </returns>
Public Function CreateStack(ByVal a_capacity As Integer) As Stack
    Set CreateStack = Me.NewStack.Initialize(a_capacity)
End Function



