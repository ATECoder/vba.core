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

