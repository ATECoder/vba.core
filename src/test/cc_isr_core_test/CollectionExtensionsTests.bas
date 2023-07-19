Attribute VB_Name = "CollectionExtensionsTests"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
' CollectionExtensionsTests.bas
'
' Dependencies:
'
' Assert.cls
' CollectionExtensions.cls
'
' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unit test. Asserts that the collection contains an expected value. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldContain() As Assert
    
    Dim p_col As VBA.Collection
    Set p_col = New VBA.Collection
    Dim p_expected As Variant: p_expected = "a"
    p_col.Add p_expected
    Set TestCollectionShouldContain = Assert.IsTrue(CollectionExtensions.ContainsKey(p_col, p_expected), "The collection should contain the value")

End Function

''' <summary>   Unit test. Asserts that the collection does not contain a value. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldNotContain() As Assert
    
    Dim p_col As VBA.Collection
    Set p_col = New VBA.Collection
    Dim p_expected As Variant: p_expected = "a"
    Dim notExpected As Variant: notExpected = "b"
    p_col.Add p_expected
    Set TestCollectionShouldNotContain = Assert.IsFalse(CollectionExtensions.ContainsKey(p_col, notExpected), "The collection should not contain a value")

End Function

''' <summary>   Unit test. Asserts that the collection contains an expected value. </summary>
''' <returns>   An <see cref="cc_isr_Test_Fx.Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestCollectionShouldContainItself() As Assert
    
    Dim p_col As New VBA.Collection
    p_col.Add "a"
    p_col.Add "b"
    Set TestCollectionShouldContainItself = Assert.IsTrue(CollectionExtensions.ContainsAll(p_col, p_col), _
                                    "The collection should contain itself")

End Function

