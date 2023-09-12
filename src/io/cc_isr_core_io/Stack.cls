VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   A circular stack class. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Type this_
    Items As Collection
    Capacity As Integer
    IsObject As Boolean
End Type

Private This As this_

''' <summary>   Initializes this object. </summary>
''' <param name="a_capacity">   [Integer] The capacity of the Stack. </param>
Public Function Initialize(ByVal a_capacity As Integer) As Stack
    Me.Clear
    This.Capacity = a_capacity
    Set Initialize = Me
End Function

''' <summary>   Gets the number of elements in the Stack. </summary>
''' <value>   [Integer]. </value>
Public Property Get Count() As Integer
    Count = This.Items.Count
End Property

''' <summary>   Gets the maximum number of elements in the Stack. </summary>
''' <value>   [Integer]. </value>
Public Property Get Capacity() As Integer
    Capacity = This.Capacity
End Property

''' <summary>   Gets the maximum number of elements in the Stack. </summary>
''' <param name="a_value">   [Integer] The new capacity of the Stack. </param>
Public Property Let Capacity(ByVal a_value As Integer)
    This.Capacity = a_value
    While Me.Count > This.Capacity
        DoEvents
        This.Items.Remove 1
    Wend
End Property

''' <summary>   Adds an item to the top of the Stack. </summary>
''' <param name="a_item">   [Variant] the item to add to the top of the Stack. </param>
''' <value>   [Integer]. </value>
Public Sub Push(ByVal a_item As Variant)
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
        This.Items.Remove 1
    Wend
End Sub

''' <summary>   Removes and returns the item from the top of the Stack. </summary>
''' <value>   [Variant]. </value>
Public Function Pop() As Variant
    If Me.Count > 0 Then
        If This.IsObject Then
            Set Pop = This.Items(Me.Count)
        Else
            Pop = This.Items(Me.Count)
        End If
        This.Items.Remove Me.Count
    Else
        If This.IsObject Then
            Set Pop = Nothing
        Else
            Pop = Nothing
        End If
    End If
End Function

''' <summary>   Returns the indexed item in the Stack. </summary>
''' <param name="a_position">   [Optional, Integer, 1] the position of the item
'''                             where the top of the stack is at 1. </param>
''' <value>   [Variant]. </value>
Public Function Peek(Optional ByVal a_position As Integer = 1) As Variant
    If Me.Count > 0 And a_position > 0 And a_position <= Me.Count Then
        Dim p_index As Integer
        p_index = Me.Count - a_position + 1
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

''' <summary>   Clears teh Stack. </summary>
Public Sub Clear()
    Set This.Items = New Collection
End Sub

''' <summary>   Initializes this object. </summary>
Private Sub Class_Initialize()
    Me.Clear
    This.Capacity = 100
End Sub

