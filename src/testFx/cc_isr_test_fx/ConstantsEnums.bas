Attribute VB_Name = "ConstantsEnums"
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Constants and Enums.  </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Public Const TestMethodPrefix As String = "Test"
Public Const BeforeAllMethodName As String = "BeforeAll"
Public Const BeforeEachMethodName As String = "BeforeEach"
Public Const AfterAllMethodName As String = "AfterAll"
Public Const AfterEachMethodName As String = "AfterEach"

''' <summary>   Enum types that lists the test methods flags. </summary>
Public Enum TestMethodFlags
    None = 0
    BeforeAll = 1
    AfterAll = 2
    BeforeEach = 4
    AfterEach = 8
End Enum


