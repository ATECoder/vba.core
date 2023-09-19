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

''' <summary>   Clamp value between (including) minimum and maximum. </summary>
''' <param name="a_value">   [Variant] The value to clamp. </param>
''' <param name="a_min">     [Variant] The minimum value. </param>
''' <param name="a_max">     [Variant] The maximum value. </param>
''' <returns>   [Variant]. </retursn>
Public Function Clamp(ByVal a_value As Variant, ByVal a_min As Variant, ByVal a_max As Variant) As Variant
    a_value = IIf(a_value < a_min, a_min, a_value)
    Clamp = IIf(a_value > a_max, a_max, a_value)
End Function

''' <summary>   Clamp value between (including) minimum and maximum. </summary>
''' <param name="a_value">   [Byte] The value to clamp. </param>
''' <param name="a_min">     [Byte] The minimum value. </param>
''' <param name="a_max">     [Byte] The maximum value. </param>
''' <returns>   [Byte]. </retursn>
Public Function ClampByte(ByVal a_value As Byte, ByVal a_min As Byte, ByVal a_max As Byte) As Byte
    a_value = IIf(a_value < a_min, a_min, a_value)
    ClampByte = IIf(a_value > a_max, a_max, a_value)
End Function

''' <summary>   Clamp value between (including) minimum and maximum. </summary>
''' <param name="a_value">   [Double] The value to clamp. </param>
''' <param name="a_min">     [Double] The minimum value. </param>
''' <param name="a_max">     [Double] The maximum value. </param>
''' <returns>   [Double]. </retursn>
Public Function ClampDouble(ByVal a_value As Double, ByVal a_min As Double, ByVal a_max As Double) As Double
    a_value = IIf(a_value < a_min, a_min, a_value)
    ClampDouble = IIf(a_value > a_max, a_max, a_value)
End Function

''' <summary>   Clamp value between (including) minimum and maximum. </summary>
''' <param name="a_value">   [Integer] The value to clamp. </param>
''' <param name="a_min">     [Integer] The minimum value. </param>
''' <param name="a_max">     [Integer] The maximum value. </param>
''' <returns>   [Integer]. </retursn>
Public Function ClampInteger(ByVal a_value As Integer, ByVal a_min As Integer, ByVal a_max As Integer) As Integer
    a_value = IIf(a_value < a_min, a_min, a_value)
    ClampInteger = IIf(a_value > a_max, a_max, a_value)
End Function

''' <summary>   Clamp value between (including) minimum and maximum. </summary>
''' <param name="a_value">   [Long] The value to clamp. </param>
''' <param name="a_min">     [Long] The minimum value. </param>
''' <param name="a_max">     [Long] The maximum value. </param>
''' <returns>   [Long]. </retursn>
Public Function ClampLong(ByVal a_value As Long, ByVal a_min As Long, ByVal a_max As Long) As Long
    a_value = IIf(a_value < a_min, a_min, a_value)
    ClampLong = IIf(a_value > a_max, a_max, a_value)
End Function

''' <summary>   Clamp value betweeen minimum and maximum. </summary>
''' <param name="a_value">   [Single] The value to limit. </param>
Public Function ClampSingle(ByVal a_value As Single, ByVal a_min As Single, ByVal a_max As Single) As Single
    a_value = IIf(a_value < a_min, a_min, a_value)
    ClampSingle = IIf(a_value > a_max, a_max, a_value)
End Function


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

''' <summary>   Returns the time since December 30, 1899 in days. </summary>
''' <returns>   [Double]. Days since December 30, 1899. </returns>
Public Function DaysNow() As Double
    
    DaysNow = Fix(VBA.Now) + VBA.Timer / Me.SecondsPerDay()
    
End Function

''' <summary>   Parses the days since December 30, 1899. </summary>
''' <param name="a_days">           [Double] the number of days since December 30, 1899. </param>
''' <param name="a_day">            [Long] the number of days since December 30, 1899. </param>
''' <param name="a_seconds">        [Long] the seconds part. </param>
''' <param name="a_milliseconds">   [Integer] the milliseconds part. </param>
Public Sub ParseDays(ByVal a_days As Double, ByRef a_day As Long, _
    ByRef a_seconds As Long, ByRef a_milliseconds As Integer)
    
    a_day = Fix(a_days)
    Dim p_seconds As Double: p_seconds = Me.SecondsPerDay * (a_days - a_day)
    Me.ParseTime p_seconds, a_seconds, a_milliseconds

End Sub

''' <summary>   Parses the time of dday since midnight. </summary>
''' <param name="a_time">           [Double] the time since midnight. </param>
''' <param name="a_seconds">        [Long] the seconds part. </param>
''' <param name="a_milliseconds">   [Integer] the milliseconds part. </param>
Public Sub ParseTime(ByVal a_time As Double, ByRef a_seconds As Long, _
    ByRef a_milliseconds As Integer)
    
    a_seconds = Fix(a_time)
    a_milliseconds = 1000 * (a_time - a_seconds)

End Sub

''' <summary>   Formats the days since December 30, 1899. </summary>
''' <param name="a_days">                 [Double] the number of days since December 30, 1899. </param>
''' <param name="a_dateTimeFormat">       [String] the day time format. </param>
''' <param name="a_millisecondsFormat">   [String] the milliseconds format. </param>
''' <returns>   [String] such as 20230901113312.700. </returns>
Public Function FormatDays(ByVal a_days As Double, _
    Optional ByVal a_dateTimeFormat As String = "YYYYMMDDhhmmss", _
    Optional ByVal a_millisecondsFormat As String = "0") As String

    Dim p_day As Long
    Dim p_seconds As Long
    Dim p_milliseconds As Integer
    Me.ParseDays a_days, p_day, p_seconds, p_milliseconds
    
    FormatDays = VBA.Format(a_days, a_dateTimeFormat) & "." & _
                 VBA.Format(p_milliseconds, a_millisecondsFormat)
    
End Function


''' <summary>   Waits for the specified duration in seconds. </summary>
''' <param name="a_seconds">   [Double] wait time in seconds. </param>
''' <returns>   [Double] wait time in seconds. </returns>
Public Function Wait(ByVal a_seconds As Double) As Double

    Dim p_startTime As Double: p_startTime = Me.DaysNow()
    Dim p_endTime As Double: p_endTime = p_startTime + (a_seconds# / Me.SecondsPerDay)
    Dim p_now As Double: p_now = Me.DaysNow()
    While p_now < p_endTime
        DoEvents
        p_now = Me.DaysNow()
    Wend
    Wait = Me.SecondsPerDay * (p_now - p_startTime)

End Function
