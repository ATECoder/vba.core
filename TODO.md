# TODO

IO: Save components only if not saved.


## Code

## Tests

## Fixes

## Updates

https://stackoverflow.com/questions/40625618/automation-error-2146232576-80131700-on-creating-an-array?rq=1
https://learn.microsoft.com/en-us/dotnet/framework/install/dotnet-35-windows

Option Explicit

Public Sub TestStringBuilderNet()

    Dim sb As StringBuilderNet
    Set sb = New StringBuilderNet
    sb.Append "1234"
    Debug.Print sb.ToString
   
    
End Sub


Implement inheritance
https://stackoverflow.com/questions/3669270/vba-inheritance-analog-of-super

IBase							IErrorTracer							IDisposable								IConnectable
	GO								TraceError								Dispose									Open; Close
					
Class B							Class ErrorTracer						Class Disposable						Class Socket
	Implements IBase				Implements IErrorTracer					Implement IDisposable					Implements IConnectable
	Public Sub Go: End Sub			Public Sub TraceError					Public Sub Dispose						Public Sub Open Close
	Pub Prop Get Super As IBase		Pub Prop Get Super As IErrorTracer		Pub Prop Get Super As IDisposable		Pub Prop Get Super As IConnectable
	Private SUb IBase_GO()			Private SUb IErrorTracer_TraceError()	Private SUb IDisposable_Dispose()		Private Sub IConnectable_Open  IConnectable_Close
	
	'
	'Note that the methods are accessible through the IBase interface
	'
	Private Sub IBase_go()
		Debug.Print "B: super.go()"
	End Sub

	Private Sub IBase_gogo()
		Debug.Print "B: super.gogo()"
	End Sub	
	
	
Class A							Class GpibLan							Class GpibLan							Class TcpClient
	Pri Type MyType					Pri Type MyType							Pri Type MyType							Pri Type MyType
	  B_ As B                         B_ As ErrorTracer			              B_ As Disposable                        B_ As Socket
	  IBase_ as IBase                 IBase_ as IErrorTracer                  IBase_ as IDisposable                   IBase_ as IConnectable
	End Type                        End Type                                End Type                                End Type

	' VBA version of 'this'			' VBA version of 'this'					' VBA version of 'this'					' VBA version of 'this'
	Private this As myType          Private this As myType                  Private this As myType                  Private this As myType
	
	'
	'Every class that implements 'B' (abstract class)
	'you must initialize in your constructor some variables
	'of instance.
	'
	Private Sub Class_Initialize()

		With this

			'we create an instance of object B, ErrorTracer, Disposable, Socket
			Set .B_ = New B; New ErrorTracer, New Disposable, New Socket

			'the variable 'IBase_' refers to the IBase interface, 
			' implemented by class B, ErrorTraceer, Disposable, Socket
			Set .IBase_ = .B_

		End With

	End Sub
		
	'Visible only for those who reference interface B
	Private Property Get B_super() As IBase

		'returns the methods implemented by 'B', through the interface IBase
		Set B_super = this.IBase_

	End Property

	Private Sub B_go()
		Debug.Print "A: go()"
	End Sub
	'==================================================

	'Class 'A' local method
	Sub localMethod1()
		Debug.Print "A: Local method 1"
	End Sub	
	
	And finally, let's create the 'main' module.

	Sub testA()

		'reference to class 'A'
		Dim objA As A

		'reference to interface 'B'
		Dim objIA As B

		'we create an instance of 'A'
		Set objA = New A

		'we access the local methods of instance 'A'
		objA.localMethod1

		'we obtain the reference to interface B (abstract class) implemented by 'A'
		Set objIA = objA

		'we access the 'go' method, implemented by interface 'B'
		objIA.go

		'we go to the 'go' method of the super class
		objIA.super.go

		'we access the 'gogo' method of the super class
		objIA.super.gogo

	End Sub
	And the output, in the verification window, will be:

	A: Local method 1
	A: go()
	B: super.go()
	B: super.gogo()

    If VBA.vbNullString = a_errProjectName Then a_errProjectName = ActiveWorkbook.VBProject.Name

    ' turn on code exporting if not deploy or read only
    Me.DisableSavingCode cc_isr_Core_io.ThisWorkbook, Me.Deployed
    Me.DisableSavingCode cc_isr_Core.ThisWorkbook, Me.Deployed
    Me.DisableSavingCode cc_isr_Test_Fx.ThisWorkbook, Me.Deployed
    Me.DisableSavingCode cc_isr_Core_Test.ThisWorkbook, Me.Deployed
    Me.DisableSavingCode cc_isr_Core_Demo.ThisWorkbook, Me.Deployed
    Me.DisableSavingCode ThisWorkbook, Me.Deployed

    ' hide the referenced workbooks
    Application.Windows(cc_isr_Core_io.ThisWorkbook.Name).Visible = False
    Application.Windows(cc_isr_Core.ThisWorkbook.Name).Visible = False
    Application.Windows(cc_isr_Test_Fx.ThisWorkbook.Name).Visible = False
    
    ' show this work book
    Application.Windows(ThisWorkbook.Name).Visible = True


''' <summary>   Handles the workbook before close event. </summary>
''' <remarks>   Disables the save dialog for deployed or read-only workbooks.
'''             Disposes any disposable worksheets. </remarks>
''' <para name="a_cancel">   [Boolean] Set to true to cancel closing. </param>
Private Sub Workbook_BeforeClose(ByRef a_cancel As Boolean)

    Const p_procedureName As String = "Workbook_BeforeClose"
   
    ' Trap errors to the error handler
    On Error GoTo err_Handler
    
    ' this disables the save dialog on read only workbooks.
    Me.MarkAsSaved cc_isr_Core_IO.ThisWorkbook, Me.Deployed
    Me.MarkAsSaved cc_isr_Core.ThisWorkbook, Me.Deployed
    Me.MarkAsSaved cc_isr_Test_Fx.ThisWorkbook, Me.Deployed
    Me.MarkAsSaved cc_isr_Core_Test.ThisWorkbook, Me.Deployed
    Me.MarkAsSaved cc_isr_Core_Demo.ThisWorkbook, Me.Deployed
    Me.MarkAsSaved ThisWorkbook, Me.Deployed
    
' . . . . . . . . . . . . . . . . . . . . . . . . . . .
exit_Handler:

    On Error GoTo 0
    Exit Sub

' . . . . . . . . . . . . . . . . . . . . . . . . . . .
err_Handler:
  
    ' append the error source
    Me.SetErrSource p_procedureName, Me.name
    
    ' display the error message
    MsgBox Me.BuildStandardErrorMessage(), vbExclamation
    
    ' exit this procedure (not an active handler)
    On Error Resume Next
    GoTo exit_Handler
    
End Sub

Private Type ThisData
    ExportCodeAfterSave As Boolean
    Deployed As Boolean
End Type

Private This As ThisData

''' <summary>   Gets the option for exporting code files after save. </summary>
''' <value>   [Boolean]. </value>
Public Property Get ExportCodeAfterSave() As Boolean
    ExportCodeAfterSave = This.ExportCodeAfterSave
End Property

''' <summary>   Enables or disables exporting code files after save. </summary>
''' <para name="value">   True to enable exporting code files after save. </param>
Public Property Let ExportCodeAfterSave(ByVal a_value As Boolean)
    This.ExportCodeAfterSave = a_value
End Property

''' <summary>   Gets the deployed status. </summary>
''' <remarks>   Code is not saved if the workbook is deployed. </remarks>
''' <value>   [Boolean]. </value>
Public Property Get Deployed() As Boolean
    Deployed = This.Deployed
End Property

''' <summary>   Sets the deployed status. </summary>
''' <para name="value">   True to set the deployed status. </param>
Public Property Let Deployed(ByVal a_value As Boolean)

    This.Deployed = a_value

    On Error Resume Next

    ' toggle deploy mode on referenced workbooks.

    cc_isr_Core_IO.ThisWorkbook.Deployed = a_value
    cc_isr_Core.ThisWorkbook.Deployed = a_value
    cc_isr_Test_Fx.ThisWorkbook.Deployed = a_value
    cc_isr_Core_Test.ThisWorkbook.Deployed = a_value
    cc_isr_Core_Demo.ThisWorkbook.Deployed = a_value
	
End Property

''' <summary>   Sets the Err object source string to project.module.procedure. </summary>
''' <param name="a_errProcedureName">   [String] Specifies the name of the procedure. </param>
''' <param name="a_errModuleName">      [String] Specifies the module name. </param>
''' <param name="a_errProjectName">     [Optional, String] Specifies the project name; otherwise the project
'''                                     name of the active workbook is used. </param>
Public Sub SetErrSource(ByVal a_errProcedureName As String, ByVal a_errModuleName As String, _
        Optional ByVal a_errProjectName As String = vbNullString)

    ' this procedure must not trap errors because it must
    ' not alter the error object.
      
    ' thus we assume that this code is robust and will
    ' not cause errors.
    
    If VBA.vbNullString = a_errProjectName Then a_errProjectName = ActiveWorkbook.VBProject.Name
    
    ' get the current source string
    
    Dim p_errorSource As String: p_errorSource = Err.Source
    
    ' build the error source.
    
    p_errorSource = a_errProjectName & "." & a_errModuleName & "." & a_errProcedureName
  
    ' Update the Err.Source
    
    Err.Source = p_errorSource
 
End Sub

''' <summary>   Builds a standard error message. </summary>
''' <param name="a_displayWarning">         [Optional, Boolean, false] True
'''                                         to display a warning rather than
'''                                         an error message. </param>
''' <param name="a_descriptionDelimiter">   [Optional, String, ': '] Specify
'''                                         the delimiter preceding the description. </param>
''' <returns>   A Standard error message string in the form: <para>
''' Error # (0x#) occurred in <c>Source</c>: Description </para><para>
''' or  </para><para>
''' Warning # (0x#) occurred in <c>Source</c>: Description  </para>
''' </returns>
Public Function BuildStandardErrorMessage(Optional ByVal a_displayWarning As Boolean = False, _
                Optional a_descriptionDelimiter As String = ": ") As String

    Dim p_builder As String
  
    ' check if we have an error
    If Err.Number <> 0 Then
    
        p_builder = p_builder & IIf(a_displayWarning, "Warning ", "Error ")
        p_builder = p_builder & Format$(Err.Number)
        
        Dim p_errNumber As Long: p_errNumber = Err.Number - vbObjectError
        
        If Abs(p_errNumber) < &HFFFF& Then
            p_errNumber = p_errNumber - &H200
            p_builder = p_builder & " (+0x"
        Else
            p_errNumber = Err.Number
            p_builder = p_builder & " (0x"
        End If
        p_builder = p_builder & Hex$(p_errNumber)
        p_builder = p_builder & ") "
        p_builder = p_builder & "occurred in "
        p_builder = p_builder & Err.Source
        p_builder = p_builder & a_descriptionDelimiter
        p_builder = p_builder & Err.Description
        
    End If

    BuildStandardErrorMessage = p_builder

End Function

''' <summary>   Marks a workbook as saved if not saved and deployed or read only. </summary>
''' <param name="a_workbook">   [Excel.Workbook] The workbook. </param>
''' <param name="a_deployed">   [Boolean] True if the workbook was deployed. </param>
Public Sub MarkAsSaved(ByVal a_workbook As Excel.Workbook, ByVal a_deployed As Boolean)
    
    If a_deployed Or a_workbook.ReadOnly Then _
        a_workbook.Saved = True

End Sub

''' <summary>   Disables saving code if deployed or read only. </summary>
''' <param name="a_workbook">   [Excel.Workbook] The workbook. </param>
''' <param name="a_deployed">   [Boolean] True if the workbook was deployed. </param>
Public Sub DisableSavingCode(ByVal a_workbook As Excel.Workbook, ByVal a_deployed As Boolean)
    
    a_workbook.ExportCodeAfterSave = Not (a_deployed Or a_workbook.ReadOnly)

End Sub

