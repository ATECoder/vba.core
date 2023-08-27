Attribute VB_Name = "CollectionExtensionsTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Unit tests. Collection extension methods. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Asserts that the collection contains an expected value. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldContain() As Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_col As VBA.Collection
    Set p_col = New VBA.Collection
    Dim p_expected As Variant: p_expected = "a"
    p_col.Add p_expected
    Set p_outcome = Assert.IsTrue(CollectionExtensions.ContainsKey(p_col, p_expected), "The collection should contain the value")

    Debug.Print p_outcome.BuildReport("TestCollectionShouldContain")
    
    Set TestCollectionShouldContain = p_outcome

End Function

''' <summary>   Unit test. Asserts that the collection does not contain a value. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldNotContain() As Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_col As VBA.Collection
    Set p_col = New VBA.Collection
    Dim p_expected As Variant: p_expected = "a"
    Dim notExpected As Variant: notExpected = "b"
    p_col.Add p_expected
    Set p_outcome = Assert.IsFalse(CollectionExtensions.ContainsKey(p_col, notExpected), "The collection should not contain a value")

    Debug.Print p_outcome.BuildReport("TestCollectionShouldNotContain")
    
    Set TestCollectionShouldNotContain = p_outcome

End Function

''' <summary>   Unit test. Asserts that the collection contains an expected value. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldContainItself() As Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_col As New VBA.Collection
    p_col.Add "a"
    p_col.Add "b"
    Set p_outcome = Assert.IsTrue(CollectionExtensions.ContainsAll(p_col, p_col), _
                                    "The collection should contain itself")

    Debug.Print p_outcome.BuildReport("TestCollectionShouldContainItself ")
    
    Set TestCollectionShouldContainItself = p_outcome

End Function

''' <summary>   Unit test. Asserts that collection are equal. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldBeEqual() As Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_left As VBA.Collection
    Set p_left = New VBA.Collection
    Dim p_right As VBA.Collection
    Set p_right = New VBA.Collection
    Dim p_item As String
    p_item = "a": p_left.Add p_item: p_right.Add p_item
    p_item = "b": p_left.Add p_item: p_right.Add p_item
    
    Set p_outcome = Assert.IsTrue(CollectionExtensions.AreEqual(p_left, p_right), _
        "The collection should be equal")

    Debug.Print p_outcome.BuildReport("TestCollectionShouldBeEqual")
    
    Set TestCollectionShouldBeEqual = p_outcome

End Function

''' <summary>   Unit test. Asserts that collection are not equal. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/> instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldNotBeEqual() As Assert
    
    Dim p_outcome As cc_isr_Test_Fx.Assert
    Dim p_left As VBA.Collection
    Set p_left = New VBA.Collection
    Dim p_right As VBA.Collection
    Set p_right = New VBA.Collection
    Dim p_item As String
    p_item = "a": p_left.Add p_item: p_right.Add p_item
    p_item = "b": p_left.Add p_item: p_right.Add p_item
    p_item = "c": p_left.Add p_item
    
    Set p_outcome = Assert.IsFalse(CollectionExtensions.AreEqual(p_left, p_right), _
        "The collection should not be equal because they are of difference length")

    If p_outcome.AssertSuccessful Then
        p_right.Add p_item & "d"
        Set p_outcome = Assert.IsFalse(CollectionExtensions.AreEqual(p_left, p_right), _
            "The collection should not be equal because they have difference items")
    End If
    Debug.Print p_outcome.BuildReport("TestCollectionShouldNotBeEqual")
    
    Set TestCollectionShouldNotBeEqual = p_outcome

End Function



