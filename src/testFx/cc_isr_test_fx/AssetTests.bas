Attribute VB_Name = "AssetTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Test assertion tests.  </summary>
''' <remarks>   Dependencies: Assert.cls.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Unit test. Asserting <see cref="Assert.Fail"/> should report failure. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestAssertingFailShouldReportFailure() As Assert

    Dim p_outcome As Assert
    
    Set p_outcome = Assert.Fail("Asserting Fail to test failure reporting.")
    
    Set p_outcome = Assert.IsFalse(p_outcome.AssertSuccessful, "Asserting failure should report AssertSuccessful as false.")
    
    Debug.Print "TestAssertingFailShouldReportFailure " & _
        IIf(p_outcome.AssertSuccessful, "passed.", "failed: " & p_outcome.AssertMessage)
    
    Set TestAssertingFailShouldReportFailure = p_outcome
    
End Function

''' <summary>   Unit test. Asserting <see cref="Assert.Pass"/> should report pass. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestAssertingPassShouldReportPass() As Assert

    Dim p_outcome As Assert
    
    Set p_outcome = Assert.Pass("Asserting Pass to test Pass reporting.")
    
    Set p_outcome = Assert.IsTrue(p_outcome.AssertSuccessful, "Asserting Pass should report AssertSuccessful as True.")
    
    Debug.Print "TestAssertingPassShouldReportPass " & _
        IIf(p_outcome.AssertSuccessful, "passed.", "Passed: " & p_outcome.AssertMessage)
    
    Set TestAssertingPassShouldReportPass = p_outcome
    
End Function

''' <summary>   Unit test. Asserting nothing should assert nothing. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestNothingShouldBeAsserted() As Assert

    Dim p_object As Object
    Set p_object = Nothing
    
    Dim p_outcome As Assert
    
    Set p_outcome = Assert.IsNothing(p_object, "Object should be noting.")
    
    Debug.Print "TestNothingShouldBeAsserted " & _
        IIf(p_outcome.AssertSuccessful, "passed.", "failed: " & p_outcome.AssertMessage)
    
    Set TestNothingShouldBeAsserted = p_outcome
    
End Function

''' <summary>   Unit test. Asserting not nothing should not assert nothing. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestNothingShouldNotBeAsserted() As Assert

    Dim p_object As Object
    Set p_object = Assert
    
    Dim p_outcome As Assert
    
    Set p_outcome = Assert.IsNotNothing(p_object, "Object should be not be noting.")
    
    Debug.Print "TestNothingShouldNotBeAsserted " & _
        IIf(p_outcome.AssertSuccessful, "passed.", "failed: " & p_outcome.AssertMessage)
    
    Set TestNothingShouldNotBeAsserted = p_outcome
    
End Function

''' <summary>   Unit test. Asserting Null should assert Null. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestNullShouldBeAsserted() As Assert

    Dim p_object As Object
    Dim p_value As Integer
    Dim p_variant As Variant
    
    Dim p_outcome As Assert
    
    Set p_outcome = Assert.IsNull(p_value, "Integer value should be Null (Not IsObject()).")
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNull(p_variant, "Unset Variant should be Null (Not IsObject()).")
    
    End If
        
    p_variant = CInt(0)
     
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNull(p_variant, "Variant set to integer should be Null (Not IsObject()).")
    
    End If
     
    p_variant = "a"
     
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNull(p_variant, "Variant set to a string should be Null (Not IsObject()).")
    
    End If
     
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNull(Empty, "'Empty' should be Null (Not IsObject()).")
    
    End If
        
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNull(Null, "'Null' should be Null (Not IsObject()).")
    
    End If
    
    Debug.Print "TestNullShouldBeAsserted " & _
        IIf(p_outcome.AssertSuccessful, "passed.", "failed: " & p_outcome.AssertMessage)
    
    Set TestNullShouldBeAsserted = p_outcome
    
End Function

''' <summary>   Unit test. Asserting not Null should not assert Null. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestNullShouldNotBeAsserted() As Assert

    Dim p_object As Object
    Dim p_variant As Variant
    
    Dim p_outcome As Assert
    
    Set p_outcome = Assert.IsNotNull(p_object, "Object should be not be Null (IsObject()).")
    
    Debug.Print "TestNullShouldNotBeAsserted " & _
        IIf(p_outcome.AssertSuccessful, "passed.", "failed: " & p_outcome.AssertMessage)
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsTrue(VBA.IsObject(Nothing), "IsObject(Nothing) should be true.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsFalse(VBA.IsNull(Nothing), "IsNull(Nothing) should be false.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsTrue(VBA.IsObject(Nothing), "VBA.IsObject(Nothing) should be true.")
    
    End If
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNotNull(Nothing, "'Nothing' should not be Null (IsObject()).")
    
    End If
    
    Set p_variant = p_object
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsNotNull(Nothing, "Variant set to an object should not be Null (IsObject()).")
    
    End If
    
    Debug.Print "TestNullShouldNotBeAsserted " & _
        IIf(p_outcome.AssertSuccessful, "passed.", "failed: " & p_outcome.AssertMessage)
    
    Set TestNullShouldNotBeAsserted = p_outcome
    
End Function

''' <summary>   Unit test. Asserting <see cref="Assert.Same"/> should report success if
'''             objects are the same and failure if objects are not the same. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestAssertingSamenessShouldReportSameness() As Assert

    Dim p_outcome As Assert
    
    Dim p_object1 As Object
    Dim p_object2 As Variant
    
    ' set the object tothe testing sheet.
    Set p_object1 = cc_isr_Test_Fx.Testing
    Set p_object2 = p_object1
    
    Set p_outcome = Assert.AreSame(p_object1, p_object2, "The objects should be the same.")
    Set p_outcome = Assert.IsTrue(p_outcome.AssertSuccessful, _
            "Asserting sameness on the same objects should report AssertSuccessful as True.")
    
    If p_outcome.AssertSuccessful Then
    
        Set p_object2 = Nothing
        Set p_outcome = Assert.AreSame(p_object1, p_object2, "The objects should not be the same.")
        Set p_outcome = Assert.IsFalse(p_outcome.AssertSuccessful, _
                "Asserting sameness on different objects should report AssertSuccessful as False.")
    
    End If
    
    Debug.Print "TestAssertingSamenessShouldReportSameness " & _
        IIf(p_outcome.AssertSuccessful, "passed.", "failed: " & p_outcome.AssertMessage)
    
    Set TestAssertingSamenessShouldReportSameness = p_outcome
    
End Function

''' <summary>   Unit test. Asserting <see cref="Assert.NotSame"/> should report success if
'''             objects are not the same and failure if objects are the same. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestAssertingNonSamenessShouldReportNonSameness() As Assert

    Dim p_outcome As Assert
    
    Dim p_object1 As Object
    Dim p_object2 As Variant
    
    ' set the object tothe testing sheet.
    Set p_object1 = cc_isr_Test_Fx.Testing
    Set p_object2 = p_object1
    
    Set p_outcome = Assert.AreNotSame(p_object1, p_object2, "The objects should be the same.")
    Set p_outcome = Assert.IsFalse(p_outcome.AssertSuccessful, _
            "Asserting non sameness on the same objects should report AssertSuccessful as False.")
    
    If p_outcome.AssertSuccessful Then
    
        Set p_object2 = Nothing
        Set p_outcome = Assert.AreNotSame(p_object1, p_object2, "The objects should not be the same.")
        Set p_outcome = Assert.IsTrue(p_outcome.AssertSuccessful, _
                "Asserting non sameness on different objects should report AssertSuccessful as True.")
    
    End If
   
    Debug.Print "TestAssertingNonSamenessShouldReportNonSameness " & _
        IIf(p_outcome.AssertSuccessful, "passed.", "Passed: " & p_outcome.AssertMessage)
    
    Set TestAssertingNonSamenessShouldReportNonSameness = p_outcome
    
End Function





