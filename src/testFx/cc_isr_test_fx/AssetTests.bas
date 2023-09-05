Attribute VB_Name = "AssetTests"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Test assertion tests.  </summary>
''' <remarks>   Dependencies: Assert.cls.
''' </remarks>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit

''' <summary>   Unit test. Asserting <see cref="Assert.Inconclusive"/> should report Inconclusive. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertInconclusive"/> True. </returns>
Public Function TestAssertingInconclusiveShouldReportInconclusive() As Assert

    Dim p_assert As Assert

    Dim p_outcome As Assert
    
    Set p_assert = Assert.Inconclusive("Asserting Inconclusive to test inconclusive outcome.")
    
    Set p_outcome = Assert.IsTrue(p_assert.AssertInconclusive, "Asserting inconclusive should report AssertInconclusive as True.")
    
    If p_outcome.AssertSuccessful Then
    
        Set p_outcome = Assert.IsFalse(p_assert.AssertSuccessful, "Asserting inconclusive should report AssertSuccessful as False.")
    
    End If
    
    Debug.Print p_outcome.BuildReport("TestAssertingInconclusiveShouldReportInconclusive")
    
    Set TestAssertingInconclusiveShouldReportInconclusive = p_outcome
    
End Function

''' <summary>   Unit test. Asserting <see cref="Assert.Fail"/> should report failure. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestAssertingFailShouldReportFailure() As Assert

    Dim p_assert As Assert
    Dim p_outcome As Assert
    
    Set p_assert = Assert.Fail("Asserting Fail to test failure outcome.")
    
    Set p_outcome = Assert.IsFalse(p_assert.AssertSuccessful, "Asserting failure should report AssertSuccessful as false.")
    
    Debug.Print p_outcome.BuildReport("TestAssertingFailShouldReportFailure")
    
    Set TestAssertingFailShouldReportFailure = p_outcome
    
End Function

''' <summary>   Unit test. Asserting <see cref="Assert.Pass"/> should report pass. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestAssertingPassShouldReportPass() As Assert

    Dim p_assert As Assert
    Dim p_outcome As Assert
    
    Set p_assert = Assert.Pass("Asserting Pass to test Pass outcome.")
    
    Set p_outcome = Assert.IsTrue(p_assert.AssertSuccessful, "Asserting Pass should report AssertSuccessful as True.")
    
    Debug.Print p_outcome.BuildReport("TestAssertingPassShouldReportPass")
    
    Set TestAssertingPassShouldReportPass = p_outcome
    
End Function

''' <summary>   Unit test. Asserting nothing should assert nothing. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestNothingShouldBeAsserted() As Assert

    Dim p_object As Object
    Set p_object = Nothing
    
    Dim p_outcome As Assert
    
    Set p_outcome = Assert.IsNothing(p_object, "Object should be noting.")
    
    Debug.Print p_outcome.BuildReport("TestNothingShouldBeAsserted")
    
    Set TestNothingShouldBeAsserted = p_outcome
    
End Function

''' <summary>   Unit test. Asserting not nothing should not assert nothing. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestNothingShouldNotBeAsserted() As Assert

    Dim p_object As Object
    Set p_object = Assert
    
    Dim p_outcome As Assert
    
    Set p_outcome = Assert.IsNotNothing(p_object, "Object should be not be noting.")
    
    Debug.Print p_outcome.BuildReport("TestNothingShouldNotBeAsserted")
    
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
    
    Debug.Print p_outcome.BuildReport("TestNullShouldBeAsserted")
    
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
    
    Debug.Print p_outcome.BuildReport("TestNullShouldNotBeAsserted")
    
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
    
    Debug.Print p_outcome.BuildReport("TestAssertingSamenessShouldReportSameness")
    
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
   
    Debug.Print p_outcome.BuildReport("TestAssertingNonSamenessShouldReportNonSameness")
    
    Set TestAssertingNonSamenessShouldReportNonSameness = p_outcome
    
End Function

''' <summary>   Unit test. String equality should work. </summary>
''' <returns>   [<see cref="Assert"/>] with <see cref="Assert.AssertSuccessful"/> True if the test passed. </returns>
Public Function TestStringEqualityShouldWork() As Assert

    Dim p_outcome As Assert
    
    Dim p_expected As String
    Dim p_actual As String
    
    p_expected = "ALL CAPS"
    p_actual = "ALL CAPS"
    Set p_outcome = Assert.AreEqualString(p_expected, p_actual, VBA.VbCompareMethod.vbBinaryCompare, _
        "The two strings should equal using binary compare.")
        
    If p_outcome.AssertSuccessful Then
    
        p_expected = "ALL CAPS"
        p_actual = "all caps"
        Set p_outcome = Assert.AreNotEqualString(p_expected, p_actual, VBA.VbCompareMethod.vbBinaryCompare, _
            "The two string should not equal using binary compare.")
    
    End If
   
    If p_outcome.AssertSuccessful Then
    
        p_expected = "ALL CAPS"
        p_actual = "all caps"
        Set p_outcome = Assert.AreEqualString(p_expected, p_actual, VBA.VbCompareMethod.vbTextCompare, _
            "The two string should equal using text compare.")
    
    End If
   
    Debug.Print p_outcome.BuildReport("TestStringEqualityShouldWork")
    
    Set TestStringEqualityShouldWork = p_outcome
    
End Function



