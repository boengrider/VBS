Class CliArgsParser

	Private mandatoryArgsCount,providedArgsCount
	Private dictArgs
	Private dictProvidedArgs
	Private dictMissingArgs
	
	Private Sub Class_Initialize()
		providedArgsCount = 0 ' Initially 0
		mandatoryArgsCount = 0 ' Initially 0
		Set dictProvidedArgs = CreateObject("Scripting.Dictionary")
		Set dictMissingArgs = CreateObject("Scripting.Dictionary")
		Set dictArgs = CreateObject("Scripting.Dictionary")
	End Sub 
	
	Public Function ParseArguments
		Dim arg,i
		
		For i = 0 To WScript.Arguments.Count - 1 ' if count 3 then 0, 1 , 2
		
			If dictArgs.Exists(WScript.Arguments.Item(i)) And i < (WScript.Arguments.Count - 1) Then
				dictProvidedArgs.Add WScript.Arguments.Item(i),WScript.Arguments.Item(i + 1)
			End If
			
			If dictArgs.Item(WScript.Arguments.Item(i)) = "1" Then ' This argument is marked as mandatory
				debug.WriteLine "Argument " & WScript.Arguments.Item(i)
				providedArgsCount = providedArgsCount + 1
			End If 
			
		Next	
	End Function 
	
	Public Function InsertArgs(strArgs,strMask,intCount) ' Comma delimited string of arguments and bit mask indicating if an argument is required or optional 
		Dim i,arg
		i = 1
		For Each arg In Split(strArgs,",")
			If Mid(strMask,i,1) = "1" Then
				dictArgs.Add arg,True
				mandatoryArgsCount = mandatoryArgsCount + 1
				debug.WriteLine "Mandatory arguments count: " & mandatoryArgsCount
			Else 
				dictArgs.Add arg,False
			End If
			i = i + 1
		Next
	End Function
	
	Public Property Get IsMandatory(strName)
		IsMandatory = dictArgs.Item(strName)
	End Property 
	
	Public Property Get Errors
		debug.WriteLine "Total arguments count: " & dictArgs.Count
		debug.WriteLine "Mandatory arguments count: " & mandatoryArgsCount
		debug.WriteLine "Provided arguments count: " & providedArgsCount
		debug.WriteLine "Provided arguments count: " & dictProvidedArgs.Count
		If mandatoryArgsCount = providedArgsCount Then
			Errors = False
		Else 
			Errors = True
		End If
	End Property 
	
	Public Property Get GetArgValue(strName)
		GetArgValue = dictProvidedArgs.Item(strName)
	End Property 
	
End Class	

Dim cap
Set cap = New CliArgsParser
cap.InsertArgs "--system,--client,--directory,--period","1110",4
cap.ParseArguments
debug.WriteLine CStr(cap.Errors)
