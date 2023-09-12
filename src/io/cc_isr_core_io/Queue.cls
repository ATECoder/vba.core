VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A circular queue class. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Items As Collection
    Capacity As Integer
    IsObject As Boolean
End Type

Private This As this_

''' <summary>   Initializes this object. </summary>
''' <param name="a_capacity">   [Integer] The capacity of the queue. </param>
Public Function Initialize(ByVal a_capacity As Integer) As Queue
    Me.Clear
    This.Capacity = a_capacity
    Set Initialize = Me
End Function

''' <summary>   Gets the number of elements in the Queue. </summary>
''' <value>   [Integer]. </value>
Public Property Get Count() As Integer
    Count = This.Items.Count
End Property

''' <summary>   Gets the maximum number of elements in the Queue. </summary>
''' <value>   [Integer]. </value>
Public Property Get Capacity() As Integer
    Capacity = This.Capacity
End Property

''' <summary>   Gets the maximum number of elements in the Queue. </summary>
''' <param name="a_value">   [Integer] The new capacity of the queue. </param>
Public Property Let Capacity(ByVal a_value As Integer)
    This.Capacity = a_value
    While Me.Count > This.Capacity
        DoEvents
        Me.Dequeue
    Wend
End Property

''' <summary>   Adds an item to the begging of the queue. </summary>
''' <param name="a_item">   [Variant] the item to add to the queue. </param>
''' <value>   [Integer]. </value>
Public Sub Enqueue(ByVal a_item As Variant)
    If Me.Count = 0 Then
        This.IsObject = VBA.IsObject(a_item)
    ElseIf This.IsObject <> VBA.IsObject(a_item) Then
        UserDefinedErrors.RaiseError UserDefinedErrors.InvalidOperationError, _
            ThisWorkbook.VBProject.Name & "." & TypeName(Me) & ".Push", _
            " A stack must contain either objects or values."
    End If
    This.Items.Add a_item
    While Me.Count > This.Capacity
        DoEvents
        Me.Dequeue
    Wend
End Sub

''' <summary>   Removes and returns the first item in the queue. </summary>
''' <value>   [Variant]. </value>
Public Function Dequeue() As Variant
    If Me.Count > 0 Then
        If This.IsObject Then
            Set Dequeue = This.Items(1)
        Else
            Dequeue = This.Items(1)
        End If
        This.Items.Remove 1
    Else
        If This.IsObject Then
            Set Dequeue = Nothing
        Else
            Dequeue = Nothing
        End If
    End If
End Function

''' <summary>   Returns the indexed item in the queue. </summary>
''' <param name="a_position">   [Optional, Integer, 1] the position of the item
'''                             where the first queued item is at 1. </param>
''' <value>   [Variant]. </value>
Public Function Peek(Optional ByVal a_position As Integer = 1) As Variant
    If Me.Count > 0 And a_position > 0 And a_position <= Me.Count Then
        If This.IsObject Then
            Set Peek = This.Items(a_position)
        Else
            Peek = This.Items(a_position)
        End If
    Else
        If This.IsObject Then
            Set Peek = Nothing
        Else
            Peek = Nothing
        End If
    End If
End Function

''' <summary>   Clears teh queue. </summary>
Public Sub Clear()
    Set This.Items = New Collection
End Sub

''' <summary>   Initializes this object. </summary>
Private Sub Class_Initialize()
    Me.Clear
    This.Capacity = 100
End Sub
