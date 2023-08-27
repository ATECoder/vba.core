VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Collection extensions. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Returns true if the collections are equal. </summary>
''' <param name="a_left">    [Collection] The left hand side collection. </param>
''' <param name="a_right">   [Collection] The right hand side collection. </param>
''' <returns>   [Bollean] True if the collections are equal. Requires that the collection items
'''             implement Equals. </returns>
Public Function AreEqual(ByVal a_left As VBA.Collection, ByVal a_right As VBA.Collection) As Boolean
    
    Dim p_areEqual As Boolean
    If a_left Is Nothing Then
        p_areEqual = False
    ElseIf a_right Is Nothing Then
        p_areEqual = False
    ElseIf a_left.Count <> a_right.Count Then
        p_areEqual = False
    Else
        p_areEqual = True
        Dim p_count As Integer: p_count = a_left.Count
        Dim i As Integer
        For i = 1 To p_count
            If a_left(i) <> a_right(i) Then
                p_areEqual = False
                Exit For
            End If
        Next i
    End If
    
    AreEqual = p_areEqual
    
End Function



''' <summary>   Returns true if the collection contains the specified key. </summary>
''' <param name="a_col">     [Collection] The subject collection. </param>
''' <param name="a_key">     [Variant] The key to check for in the collection. </param>
''' <returns>   True if the key is contained in the collection. </returns>
Public Function ContainsKey(ByVal a_col As VBA.Collection, ByVal a_key As Variant) As Boolean
    
    Dim p_found As Boolean
    p_found = False
    Dim colItem As Variant
    For Each colItem In a_col
        DoEvents
        If colItem = a_key Then
            p_found = True
            Exit For
        End If
    Next colItem
    ContainsKey = p_found

#If False Then
    ' this causes error 5 event if doing this:
    Print a_col.Item(1)
    StringExtensionsTests
    Print a_col.Item("StringExtensionsTests")
    ' or even this!
    Print a_col.Item(a_col.Item(1))
    On Error Resume Next
    a_col a_key ' Just try it. If it fails, Err.Number will be nonzero.
    ContainsKey = (Err.Number = 0)
#End If

End Function

''' <summary>   Returns true if the object is contained in the collection. </summary>
''' <param name="a_col">         [Collection] The subject collection. </param>
''' <param name="a_contained">   [Collection] The collection which to check for being contained in the
'''                              subject collection. </param>
''' <returns>   True if the contained collection is fully contained in the collection. </returns>
Public Function ContainsAll(ByVal a_col As VBA.Collection, ByVal a_contained As VBA.Collection) As Boolean
    
    Dim p_result As Boolean: p_result = True
    Dim p_key As Variant
    For Each p_key In a_contained
        DoEvents
        If Not ContainsKey(a_col, p_key) Then
            p_result = False
            Exit For
        End If
    Next p_key
    ContainsAll = p_result
    
End Function

''' <summary>   Returns the first item that exists in <paramref name="a_contained"/>
''' not existing in <paramref name="a_col"/>. </summary>
''' <param name="a_col">         [Collection] The subject collection. </param>
''' <param name="a_contained">   [Collection] The collection which to check for being contained in the
'''                              subject collection. </param>
Public Function FindMissingItem(ByVal a_col As VBA.Collection, ByVal a_contained As VBA.Collection) As Variant
    
    Dim p_result As Variant: Set p_result = Nothing
    Dim p_key As Variant
    For Each p_key In a_contained
        DoEvents
        If Not ContainsKey(a_col, p_key) Then
            p_result = p_key
            Exit For
        End If
    Next p_key
    Set FindMissingItem = p_result
    
End Function

''' <summary>   Adds items from the source to the destination collection. </summary>
''' <param name="a_source">        [Collection] The source collection. </param>
''' <param name="a_destination">   [Collection] The destination collection which receives the added items. </param>
''' <returns>   The number of items that were added. </returns>
Public Function AddItems(ByVal a_source As VBA.Collection, ByVal a_destination As VBA.Collection) As Integer
    
    Dim p_count As Integer: p_count = 0
    Dim p_item As Variant
    For Each p_item In a_source
        DoEvents
        a_destination.Add p_item
        p_count = p_count + 1
    Next p_item
    AddItems = p_count
    
End Function

''' <summary>   Returns true if the collection is nothing or empty. </summary>
''' <param name="a_collection">   [Collection] the collection. </param>
''' <returns>   [Boolean]. </value>
Public Function CollectionIsNothingOrEmpty(ByVal a_collection As Collection)
    If a_collection Is Nothing Then
        CollectionIsNothingOrEmpty = True
    ElseIf 0 = a_collection.Count Then
        CollectionIsNothingOrEmpty = True
    Else
        CollectionIsNothingOrEmpty = False
    End If
End Function

