Attribute VB_Name = "AssetTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Test assertion tests.  </summary>
''' <remarks>   Dependencies: Assert.cls.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Unit test. Asserts creating a list of test modules. </summary>
''' <returns>   An <see cref="Assert"/>   instance of <see cref="Assert.AssertSuccessful"/>   True if the test passed. </returns>
Public Function TestNothingAssertion() As Assert

    Dim p_object As Object
    Set p_object = Nothing
    
    Dim p_outcome As Assert
    
    Set p_outcome = Assert.IsNothing(p_object, "Object should be noting")
    
    Debug.Print "TestNothingAssertion " & _
        IIf(p_outcome.AssertSuccessful, "passed.", "failed: " & p_outcome.AssertMessage)
    
    Set TestNothingAssertion = p_outcome
    
End Function



