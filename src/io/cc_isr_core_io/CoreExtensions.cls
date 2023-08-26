VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoreExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''' <summary>   Core extensions. </summary>
''' - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

''' <summary>   Unbox a parameter array in case element zero is an array, which occurs when a parameter
''' array argument is passed to another function that expects a ParamArray argument.
''' </summary>
''' <remarks>
''' <see href="https://stackoverflow.com/questions/20783170/pass-array-to-paramarray"/> <para>
''' When passing a parameter array that is received as an argument in function A to another function
''' B that expects a parameter array, the parameter array is received in function A appears at element
''' zero of the function B. </para><para>
''' Generally, when each subsequent function receives the parameter array, and passes it on, it increases
''' the depth of the parameter array. </para><para>
''' This method undoes the depth levels added and returns a vector with arguments received as a Variant.
''' <code>
''' </code>
''' </para></remarks>
''' <param name="a_values">   [Variant, ParamArray] The parameter values to be unboxed. </param>
''' <returns>   [Variant, Array] The unboxed parameter values. </returns>
Public Function UnboxParameterArray(ParamArray a_values() As Variant) As Variant

    Dim p_params() As Variant, p_workArray() As Variant
    
    ' When a_values(0) is Array, it is assumed that the
    ' parameter array was passed from another function
    p_params = a_values
    Do While VBA.VarType(p_params(0)) >= VBA.vbArray And UBound(p_params) < 1
        p_workArray = p_params(0)
        p_params = p_workArray
    Loop
    UnboxParameterArray = p_params
    
End Function

''' <param name="a_variableType">   [VBA.VbVarType] subtype of a variable, or the type
'''                                 of an object's default property. </param>
''' <returns>   [Variant]. </returns>
Public Function GetDefaultValue(ByVal a_variableType As VBA.VbVarType) As Variant

    Select Case a_variableType
        
        Case VBA.VbVarType.vbString
            GetDefaultValue = VBA.vbNullString
        
        Case VBA.VbVarType.vbBoolean
            GetDefaultValue = False
        
        Case VBA.VbVarType.vbByte, VBA.VbVarType.vbCurrency, VBA.VbVarType.vbDate, _
             VBA.VbVarType.vbDecimal, VBA.VbVarType.vbDouble, VBA.VbVarType.vbInteger, _
             VBA.VbVarType.vbLong, VBA.VbVarType.vbSingle
            GetDefaultValue = 0
        
        Case VBA.VbVarType.vbArray, VBA.VbVarType.vbEmpty, VBA.VbVarType.vbVariant
            GetDefaultValue = Empty
        
        Case VBA.VbVarType.vbNull
            GetDefaultValue = Null
        
        Case VBA.VbVarType.vbObject
            Set GetDefaultValue = Nothing
#If VBA7 Then
#If Win64 Then
        Case VBA.VbVarType.vbLongLong
            GetDefaultValue = 0
#End If
#End If
    End Select
End Function

''' <summary>   Queries whether code is running in the IDE (true), or compiled (false). </summary>
''' <returns>   [Boolean]. </returns>
Public Function RunningInIDE() As Boolean

    Static p_counter As Variant
    If IsEmpty(p_counter) Then
        p_counter = 1
        Debug.Assert RunningInIDE() Or True
        p_counter = p_counter - 1
    ElseIf p_counter = 1 Then
        p_counter = 0
    End If
    RunningInIDE = p_counter
 
End Function

''' <summary>   Returns the VBA.Timer resolution. </summary>
''' <remarks>
''' To see the SystemClockRate value in clock ticks, use the following formula: (value in seconds) × 1,000 × 10,000 For example, if SystemClockRate is 0.0156250 seconds, the value that the equation uses is 156,250 clock ticks.
''' <see href="https://answers.microsoft.com/en-us/windows/forum/all/did-default-vba-timer-resolution-system-tick/d5d24e18-3e47-43e0-a4f5-c2791388e5fe"/>
''' <see href="https://learn.microsoft.com/en-us/windows-server/networking/windows-time-service/windows-time-service-tools-and-settings?tabs=config"/>
''' </remarks>
Public Function TimerResolution() As Double
    TimerResolution = 1# / 64#
End Function

''' <summary>   Returns the number of seconds per day. </summary>
Public Function SecondsPerDay() As Double

    SecondsPerDay = 60# * 60# * 24
    
End Function

''' <summary>   Returns the date and time with resolution of the VBA.Timer. </summary>
Public Function DateTimeNow() As Double
    
    DateTimeNow = Fix(VBA.Now) + VBA.Timer / Me.SecondsPerDay()
    
End Function

''' <summary>   Waits for the specified duration in seconds. </summary>
''' <param name="a_seconds">   [Double] wait time in seconds. </param>
''' <returns>   [Double] wait time in seconds. </returns>
Public Function Wait(ByVal a_seconds As Double) As Double

    Dim p_startTime As Double: p_startTime = Me.DateTimeNow()
    Dim p_endTime As Double: p_endTime = p_startTime + (a_seconds# / Me.SecondsPerDay)
    Dim p_now As Double: p_now = Me.DateTimeNow()
    While p_now < p_endTime
        DoEvents
        p_now = Me.DateTimeNow()
    Wend
    Wait = Me.SecondsPerDay * (p_now - p_startTime)

End Function
