' ------------ Logger----------------

Class Logger
	
	Private boolRemoteLogExists		' Set to True if log file exists
	Private boolLocalLogExists		' Set to True if log file exists
	Private strUser					' String representing user
	Private strComputer				' String represnting computer
	Private strDate					' Current date
	Private strSource				' Script name 
	Private strLocalLog				' Path to the local log file
	Private strRemoteLog			' UNC path to the remote log file
	Private oFSO					' File system object
	Private oNET					' Network object. Used to obtain the user name
	Private dictSeverity			' Dictionary holding severity codes
	Private oLocalLog				' Local log file descriptor
	Private oRemoteLog				' Remote log file descriptor
	
	Sub Class_Initialize
	
		boolRemoteLogExists = False	' Initially set to False
		boolLocalLogExists = False	' Initially set to False
		Set oNET = CreateObject("Wscript.Network")
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set dictSeverity = CreateObject("Scripting.Dictionary")
		dictSeverity.Add "ERR","E"
		dictSeverity.Add "INF","I"
		dictSeverity.Add "WRN","W"
		strUser = oNET.UserName
		strComputer = oNET.ComputerName
		strDate = Date()
		strSource = WScript.ScriptName
		strLocalLog = Null
		strRemoteLog = Null
		oLocalLog = Null
		oRemoteLog = Null
		
	End Sub 
	
	Sub Class_Terminate
		
		If Not IsNull(strLocalLog) Then
		
			oLocalLog.Close
		
		End If 
		
		If Not IsNull(strRemoteLog) Then
		
			oRemoteLog.Close
		
		End If 
		
	End Sub 
	
	'LogEvent method
	' strMessage -> Message to log
	' strSeverity -> Severity code, e.g "ERROR"
	' boolLogRemote -> If false log locally only. If true log remotely aswell
	Public Function LogEvent(strMessage,strSeverity,boolLogRemote)
		
		Select Case boolLogRemote
		
			Case True
			
				If Not IsNull(oLocalLog) And boolLocalLogExists Then
					oLocalLog.WriteLine strSource & vbTab & strDate & vbTab & Time() & vbTab & strUser & "@" & strComputer & vbTab & strMessage & vbTab & dictSeverity(strSeverity)
				End If 
				
				If Not IsNull(oRemoteLog) And boolRemoteLogExists Then
					oRemoteLog.WriteLine strSource & vbTab & strDate & vbTab & Time() & vbTab & strUser & "@" & strComputer & vbTab & strMessage & vbTab & dictSeverity(strSeverity)
				End If 
				
			Case False 
			
				If Not IsNull(oLocalLog) And boolLocalLogExists Then 
					oLocalLog.WriteLine strSource & vbTab & strDate & vbTab & Time() & vbTab & strUser & "@" & strComputer & vbTab & strMessage & vbTab & dictSeverity(strSeverity)
				End If 
				
		End Select 
	
	End Function 
	
	Public Function Header
	
		oLocalLog.WriteLine vbCrLf & "_____________________ " & WScript.ScriptName & " _____________________" & vbCrLf
		oRemoteLog.WriteLine vbCrLf &"_____________________ " & WScript.ScriptName & " _____________________" & vbCrLf 
		
	End Function 
	
	Public Function ReleaseLogs
		
		If Not IsNull(strLocalLog) Or Not strLocalLog = "" Then
		
			If IsObject(oLocalLog) Then 
				debug.WriteLine "Local log released"
				oLocalLog.Close
			End If 
		
		End If 
		
		If Not IsNull(strRemoteLog) Or Not strRemoteLog = "" Then
		
			If IsObject(oRemoteLog) Then 
				debug.WriteLine "Remote log released"
				oRemoteLog.Close
			End If 
		
		End If 
		
	End Function 
	
	
	

	' LocalLogFile property. 
	' Opens a log file for appending. If it doesn't exist, creates it
	Public Property Let LocalLogFile(strPath)
		
		If IsNull(strPath) Then
			boolLocalLogExists = False
			Exit Property
		End If 
		
		If oFSO.FolderExists(strPath) Then 
			Set oLocalLog = oFSO.OpenTextFile(strPath & "\log.txt",8,True)
			If oFSO.FileExists(strPath & "\log.txt") Then
				boolLocalLogExists = True
			End If 
		End If 
		
	End Property 
	
	' RemoteLogFile property. 
	' Opens a log file for appending. If it doesn't exist, creates it
	Public Property Let RemoteLogFile(strPath)
		
		If IsNull(strPath) Then
			boolRemoteLogExists = False
			Exit Property
		End If
		
		If oFSO.FolderExists(strPath) Then
			Set oRemoteLog = oFSO.OpenTextFile(strPath & "\log.txt",8,True)
			If oFSO.FileExists(strPath & "\log.txt") Then
				boolRemoteLogExists = True
			End If 
		End If 
		
	End Property 
	
End Class 