Class ProcessTerminator

	Private oNET
	Private oWMI
	Private oWSH
	Private oAPP
	Private oFSO
	Private oTOOLKIT ' Toolkit is a vbsedit dll (vbsedit64.dll). VBSedit must me installed on the system, or install library prior execution
					 ' To load it run as administrator  "regsvr32 C:\Users\a293793\AppData\Local\Adersoft\VbsEdit\x64\vbsedit64.dll"
	
	
	Private Sub Class_Initialize
	
		Set oNET = CreateObject("Wscript.Network")
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set oTOOLKIT = CreateObject("Vbsedit.Toolkit")
		Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
		Set oWSH = CreateObject("wscript.shell")
		Set oAPP = CreateObject("Shell.Application")
		
		
		
	End Sub 
	
	Private Sub Class_Terminate
	
	
	End Sub 
	
	
	Public Sub TakeScreenshot
	
		Dim screenshot
		screenshot = "C:\Users\" & oNET.UserName & "\Pictures\" & Year(Date()) & Month(Date()) & Day(Date()) & "_" & Hour(Time()) & Minute(Time()) & Second(Time()) & ".jpg" 
		oAPP.WindowSwitcher
		WScript.Sleep 1000
		oTOOLKIT.DesktopWindow.Screenshot screenshot
		WScript.Sleep 500
		oWSH.SendKeys "{ESC}"
		
		
	End Sub 
	
	Public Function Terminate
	
		Dim colProcesses,proc
		Set colProcesses = oWMI.ExecQuery("SELECT * FROM Win32_Process WHERE (Name = 'wscript.exe' OR Name = 'saplogon.exe')")
		
		If colProcesses.count <> 0 Then
		
			If colProcesses.count > 1 Then  ' Possible hanging processes. 
		 		TakeScreenshot
		 	End If 
		 	
			For Each proc In colProcesses 
				
				If InStr(proc.CommandLine,WScript.ScriptName) <> 0 Then
					debug.WriteLine "Skipping this process with ID: " & proc.ProcessId
				Else 
					debug.WriteLine "Process " & proc.Name & " with ID: " & proc.ProcessId & " terminated"
					proc.Terminate
				End If 
				
			Next
			
		End If 
		
	End Function 



End Class 

