Class RateUpload

	Private oXML
	Private oHTTP
	Private oFSO
	Private oWSH
	Private oNET
	Private oSAPGUI
	Private oAPP
	Private oCON
	Private oSES
	Private boolSAPRunning 
	Private strSSN ' Sap System Name e.g FQ2
	Private strSCN  ' Sap Client Name e.g. 105
	Private strSSD ' Sap System Description. This string is found in the local landscape xml and used to connect to the sap system
	Private strGlobalURL ' Global Landscape URL
	Private strLocalLandscapePATH ' Absolute PATH to the local Landscape xml file
	Private strUserName ' System user name e.g. a293793
	Private strComputerName ' System name e.g. SKSENEW128
	
	' ============== Constructor & Destructor ==================
	Private Sub Class_Initialize
	
		Set oXML = CreateObject("MSXML2.DOMDocument")   
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oFSO = CreateObject("scripting.filesystemobject")
		Set oWSH = CreateObject("wscript.shell")
		Set oNET = CreateObject("wscript.network")
		oSAPGUI = Null
		oAPP = Null
		oCON = Null
		oSES = Null 
		strSSN = Null
		strSCN = Null
		strGlobalURL = Null
		strLocalLandscapePATH = Null
		strSSD = Null
		boolSAPRunning = False
		strUserName = oNET.UserName
		strComputerName = oNET.ComputerName

	End Sub
	
	Private Sub Class_Terminate
	
	End Sub
	
	' ============ P U B L I C  &  P R I V A T E   M E T H O D S  &   S U B R O U T I N E S ===========
	
	Public Function Init(str_global_url,str_local_path,str_sap_systemName,str_sap_clientName)
	
		strGlobalURL = str_global_url
		strSSN = str_sap_systemName
		strSCN = str_sap_clientName
		strLocalLandscapePATH = str_local_path 
		FindSAPSystemDescription
		
	End Function 
	
	' ---------- CheckSAPLogon
	Private Sub CheckSAPLogon
	
		Dim oWmi,colProc,proc,oSAP,waitfor
		Set oWmi = GetObject("winmgmts:\\.\root\cimv2")
		Set colProc = oWmi.ExecQuery("SELECT Name, ProcessId FROM Win32_Process")
		
		On Error Resume Next 
		For Each proc In colProc
			If InStr(LCase(proc.Name),"saplogon") > 0 Then 
				Do While True 
					Set oSAPGUI = GetObject("SAPGUI") ' Wait until the object is instantiated
						If IsObject(oSAPGUI) Then
							boolSAPRunning = True
							FindSAPSession 
							Exit Sub  ' At this point we can safely assume that SAPLogon is running and SAPGUI object is available
					End If 
				Loop
			End If 
		Next 
		
		On Error GoTo 0 ' Reenable error handling
			
		'Start SAPLogon and open system passed in the command line parameter
		Set proc = oWSH.Exec("C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
		'WScript.Sleep 5000
		Set colProc = oWmi.ExecQuery("SELECT Name, ProcessId FROM Win32_Process")
		On Error Resume Next 
		For Each proc In colProc
			If InStr(LCase(proc.Name),"saplogon") > 0 Then 
				Do While True 
					Set oSAPGUI = GetObject("SAPGUI") ' Wait until the object is instantiated
						If IsObject(oSAPGUI) Then
							boolSAPRunning = True
							FindSAPSession 
							Exit Sub  ' At this point we can safely assume that SAPLogon is running and SAPGUI object is available
					End If 
				Loop
			End If 
		Next 
		
		On Error GoTo 0 ' Reenable error handling
		

	End Sub 
	
	
	
	' ---------- FindSAPSession
	Private Sub FindSAPSession
		
		If Not boolSAPRunning Then
			oSES = Null
			Exit Sub 
		End If  
		
		If IsNull(strSSD) Then
			oSES = Null
			Exit Sub
		End If
		 
		Set oAPP = oSAPGUI.GetScriptingEngine
		Select Case oAPP.Children.Count
	
			Case 0 ' No open connections exist
				Set oCON = oAPP.OpenConnection(strSSD,True,False) ' Open a new connection asynchronously
				Set oSES = oCON.Children(0) ' Attach to the first session
			
		
			Case Else ' Atleast one connection exists
				For Each oCON In oAPP.Children ' connections
					For Each oSES In oCON.Children ' sessions	
						If LCase(oSES.Info.SystemName) = LCase(strSSN) Then
							Exit Sub ' Stop here. We found our desired system. oCON and oSES objects hold our target system
						End If 
					Next
				Next
			
				Set oCON = oAPP.OpenConnection(strSSD,True,False) ' Open a new connection asynchronously
				Set oSES = oCON.Children(0) ' Attach to the first session
			
			End Select 
	
	End Sub 



	' --------- UploadRates
	Public Function UploadRates(strFiles,strExRateType) ' strFiles is comma delimited list of files to upload
	
		Dim validfrom,SAPfile,i,ratetype,filename
		i = 0
		ratetype = UCase(strExRateType)
	
		For Each SAPfile In Split(strFiles,",")
			If oFSO.FileExists(SAPfile) Then 
				filename = oFSO.GetFileName(SAPfile) ' Returns 20200630.txt 
				validfrom = "" ' Clear
				validfrom = Mid(filename,7,2) & "." & Mid(filename,5,2) & "." & Mid(filename,1,4) ' SAP compatible date format DD.MM.YYYY
				KillPopups(oSES)
				oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NZTC_ZCURR_UPLOAD"
				oSES.findById("wnd[0]").sendVKey 0 ' ENTER
				KillPopups(oSES)
				oSES.findById("wnd[0]/usr/txtP_FILE").text = SAPfile
				oSES.findById("wnd[0]/usr/txtP_KURST").text = ratetype
				oSES.findById("wnd[0]/usr/ctxtP_GDATU").text = validfrom
				oSES.findById("wnd[0]").sendVKey 8
				KillPopups(oSES)
				oSES.findById("wnd[0]").sendVKey 0
				KillPopups(oSES)
		
				Do While oSES.Children.Count > 1
					oSES.findById("wnd[0]").sendVKey 0
				Loop
				i = i + 1
				WScript.Sleep 2000 ' Wait a bit
			End If 	
		Next
		
		oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NEX" ' Close transaction window
		oSES.findById("wnd[0]").sendVKey 0
	
		UploadRates = i ' Return the number of uploaded files or 0 if error occured   

	End Function 
	
	' --------- FindSAPSystemDescription
	Private Sub FindSAPSystemDescription
	
		Dim n_ChildNodes,n_ChildNode,uuid,i,j
		oHTTP.open "GET",strGlobalURL,False
		oHTTP.send
		
		If oHTTP.status <> 200 Then
			strSSD = Null
			Exit Sub 
		End If 
		
		oXML.load(oHTTP.responseXML)
	
		Set n_ChildNodes = oXML.getElementsByTagName("Messageserver")
	
		For Each n_ChildNode In n_ChildNodes
			If LCase(n_ChildNode.attributes.getNamedItem("name").text) = LCase(strSSN) Then
				uuid = n_ChildNode.attributes.getNamedItem("uuid").text ' We found the uuid of the target system
				Exit For
			End If 
		
		Next
	
		oXML.load(strLocalLandscapePATH) ' Locally stored XML

		Set n_ChildNodes = oXML.getElementsByTagName("Landscape")
	
		For Each n_ChildNode In n_ChildNodes 
			For Each i In n_ChildNode.childNodes
				If i.baseName = "Services" Then 
					Set n_ChildNodes = i.childNodes
					For Each j In n_ChildNodes
						If j.attributes.getNamedItem("msid").text = uuid Then ' We have a match
							strSSD = j.attributes.getNamedItem("name").text
							CheckSAPLogon                                                                                                                                                                                                            
							Exit Sub ' No need to continue
						End If
					Next
				End If 
			Next
		Next
		strSSD = Null ' Not found
	End Sub 
				
				
	''============================================================
'' Program:   SUB Killpopups
'' Desc:      Kill of SAP popup screens which could appear when executing SAP transactions
'' Called by: 
'' Call:      KillPopups(connection.children(0)
'' Arguments: s = connection.children(0)
'' Changes---------------------------------------------------
'' Date		Programmer	Change
'' 2020-06-01	Tomas Chudik(tomas.chudik@volvo.com)	Written as vbscript SUB with arguments; supports kill of "System Message", "Copyright"
''============================================================

	Sub KillPopups(s)
		Do While s.Children.Count > 1
			If InStr(s.ActiveWindow.Text, "System Message") > 0 Then
				s.ActiveWindow.sendVKey 12
		
			ElseIf InStr(s.ActiveWindow.Text, "Copyright") > 0 Then
				s.ActiveWindow.sendVKey 0
				'ElseIF   'Insert next type of popup windows which you want to kill
			Else
				Exit Do
			End If
		Loop
	End Sub

	' ================= P R O P E R T I E S ====================
	Public Property Get SAPLogonRunning
		SAPLogonRunning = boolSAPRunning
	End Property 	
		
	Public Property Get SAPSessionExists
		If boolSAPRunning And Not IsNull(oSES) Then
			SAPSessionExists = True
		Else 
			SAPSessionExists = False
		End If
	End Property 
	
	Public Property Get SAPsysName
		SAPsysName = strSSN
	End Property 
	
	Public Property Get SAPcliName
		SAPcliName = strSCN
	End Property 
	
	Public Property Get LandscapeURL
		LandscapeURL = strGlobalURL
	End Property 
	
	Public Property Get LandscapePATH
		LandscapePATH = strLocalLandscapePATH
	End Property
	
	Public Property Get SAPsysDescription
		SAPsysDescription = strSSD
	End Property 
	
End Class 
