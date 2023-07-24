# TODO

    Dim p_escape As EscapeSequence
    Dim p_escapes As Collection: Set p_escapes = Me.EscapeSequences
    
    For Each p_escape In p_escapes
        
        If Me.StringContains(p_returnValue, p_escape.Value) Then _
            p_returnValue = Replace(p_returnValue, p_escape.Value, p_escape.Literal)

        If Not Me.StringContains(p_returnValue, "\") Then GoTo normalExit
    
    Next

    If Me.StringContains(a_original, "\") Then
    
        ' replace "ASCII (oct)" escape sequence
    
        Set p_regex = New VBScript_RegExp_55.RegExp
        p_regex.Pattern = "\\(\d{3})"
        p_regex.IgnoreCase = True
        p_regex.Global = True
        Set p_matches = p_regex.Execute(a_format)
    
        Dim char As Long
        If p_matches.count <> 0 Then
            For Each p_match In p_matches
                p = p_match.SubMatches(0)
                '"p" contains the octal number representing the ASCII code we're after:
                p = "&O" & p 'prepend octal prefix
                char = CLng(p)
                p_returnValue = Replace(p_returnValue, p_match.Value, Chr$(char))
            Next
        End If
    
    End If

    If Me.StringContains(a_original, "\") Then
    
        ' replace "ASCII (hex)" escape sequence
        Set p_regex = New VBScript_RegExp_55.RegExp
        p_regex.Pattern = "\\x(\w{2})"
        p_regex.IgnoreCase = True
        p_regex.Global = True
        Set p_matches = p_regex.Execute(a_format)
        
        If p_matches.count <> 0 Then
            For Each p_match In p_matches
                p = p_match.SubMatches(0)
                '"p" contains the hex value representing the ASCII code we're after:
                p = "&H" & p 'prepend hex prefix
                char = CLng(p)
                p_returnValue = Replace(p_returnValue, p_match.Value, Chr$(char))
            Next
        End If
    
    End If



## Code

## Tests
* string format
* string contains any

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
