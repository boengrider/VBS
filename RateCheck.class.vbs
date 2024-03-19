Class RateCheck
	
	Private oHTTP
	Private oXML
	Private oFSO
	Private oWSH
	Private oSAPGUI
	Private oAPP
	Private oCON
	Private oSES
	Private oFile
	Private oTempFile   ' temp file to hold data from sap TCURR
	Private boolSAPRunning ' Indicates whether SPA Logon is running
	Private strHomeCurrency ' Home currency e.g. CZK
	Private numFilesVerified ' Number of successfully verified files
	Private strSSN ' Sap System Name e.g FQ2
	Private strSCN  ' Sap Client Name e.g. 105
	Private strSSD ' Sap System Description. This string is found in the local landscape xml and used to connect to the sap system
	Private strGlobalURL ' Global Landscape URL
	Private strLocalLandscapePATH ' Absolute PATH to the local Landscape xml file
	Private strTempFilePath ' Absolute path to the temp file
	Private strTempFileName ' temp file name
	
	
	' ============== Constructor & Destructor ==================
	Private Sub Class_Initialize
	
		
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")   
		Set oFSO = CreateObject("scripting.filesystemobject")
		Set oWSH = CreateObject("wscript.shell")
		numFilesVerified = 0
		oSAPGUI = Null
		oAPP = Null
		oCON = Null
		oSES = Null 
		strTempFilePath = Null
		strTempFileName = Null
		strSSN = Null
		strSCN = Null
		strGlobalURL = Null
		strLocalLandscapePATH = Null
		strSSD = Null
		strHomeCurrency = Null
		boolSAPRunning = False
		

	End Sub
	
	Private Sub Class_Terminate
	
	End Sub
	
	' ============ P U B L I C  &  P R I V A T E   M E T H O D S  &   S U B R O U T I N E S ===========
	
	Public Function Init(str_global_url,str_local_path,str_sap_systemName,str_sap_clientName,str_home_curr)
	
		strHomeCurrency = str_home_curr
		strGlobalURL = str_global_url
		strSSN = str_sap_systemName
		strSCN = str_sap_clientName
		strLocalLandscapePATH = str_local_path 
		FindSAPSystemDescription
		
	End Function 
	
	' ------------ CreateGUID
	Private Function CreateGUID
  		Dim TypeLib
  		Set TypeLib = CreateObject("Scriptlet.TypeLib")
  		CreateGUID = Mid(TypeLib.Guid, 2, 36)
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
			
				
			End Select 
			
			Set oCON = oAPP.OpenConnection(strSSD,True,False) ' Open a new connection asynchronously
			Set oSES = oCON.Children(0) ' Attach to the first session
	End Sub 



	' --------- CheckRates
	Public Function CheckRates(strFiles,strExRateType) ' strFiles is comma delimited list of files to upload
	
		Dim SAPfile,filesverified,files
		
		files = Split(strFiles,",") ' Split files and use the first one to determine where to put temp file
		
		filesverified = 0
		
		If strHomeCurrency = Null Or strHomeCurrency = "" Then
			CheckRates = -1 ' ERROR, home currency not set
		End If 
		
		strTempFileName = CreateGUID ' Create a temp file name
		
		strTempFilePath = oFSO.GetParentFolderName(files(0)) & "\" & strTempFileName & ".txt" ' Create a temp file at the location of the input files
		oFSO.CreateTextFile strTempFilePath,True ' Create a temp file
		
		 
		If Not oFSO.FileExists(strTempFilePath) Then
			CheckRates = -1 ' ERROR creating a temp file
			Exit Function
		End If 
		 
	
		For Each SAPfile In Split(strFiles,",")
			
			Check SAPfile,strExRateType
			
		Next
		
		 
		oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NEX" ' Close transaction window
		oSES.findById("wnd[0]").sendVKey 0
	
	
		
		oTempFile.Close
		oFSO.DeleteFile strTempFilePath
		CheckRates = numFilesVerified ' Returns number of successfully verified files.
		

	End Function 
	
	
	
	Private Sub Check(strFile,strType) ' Private sub to check files. Call within for loop
	
		Dim lines,filename,gdatu,line,i,sapline,fileline,column
		i = 0
		lines = 0
		
		If Not oFSO.FileExists(strFile) Then
			Exit Sub 
		End If 
		 
		 Set oTempFile = oFSO.OpenTextFile(strTempFilePath,2) ' Open the temp file for writing
		 Set oFile = oFSO.OpenTextFile(strFile,1,False) ' Open file containing uploaded rates for reading
		 
		 Do While Not oFile.AtEndOfStream
		 	oFile.ReadLine
		 Loop
		 	
		 lines = oFile.Line - 1
		 	
		 	
		filename = oFSO.GetBaseName(strFile) ' Returns 20200630 
		gdatu = 99999999 - filename
		KillPopups(oSES)
		oSES.findById("wnd[0]/tbar[0]/okcd").text = "/nse17"
		oSES.findById("wnd[0]").sendVKey 0 ' ENTER
		KillPopups(oSES)
		oSES.findById("wnd[0]/usr/ctxtDD02V-TABNAME").text = "TCURR"
		oSES.findById("wnd[0]").sendVKey 0 ' ENTER
		KillPopups(oSES)
		' FIELDS
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-FSELECT[1,1]").text = LCase(strType)
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-FSELECT[1,2]").text = ""
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-FSELECT[1,3]").text = ""
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-FSELECT[1,4]").text = gdatu
		' FLAGS
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,1]").text = ""  ' MANDT. SAP puts X there despite unchecking it in the script
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,1]").text = ""  ' KURST
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,2]").text = "X" ' FCURR
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,3]").text = "X" ' TCURR
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,4]").text = ""  ' GDATU
		oSES.findById("wnd[0]/usr/tblSAPMSTAZTABCTRL200/ctxtRSTAZ-SHOWFLAG[2,5]").text = "X" ' UKURS
		oSES.findById("wnd[0]").sendVKey 8
		KillPopups(oSES)
			
		Do While lines <> i 
		
			oTempFile.Write oSES.findById("wnd[0]/usr/lbl[10," & (9 + i) & "]").text          ' Foreign currency
			oTempFile.Write oSES.findById("wnd[0]/usr/lbl[18," & (9 + i) & "]").text          ' Home currency
	 		oTempFile.Write oSES.findById("wnd[0]/usr/lbl[26," & (9 + i) & "]").text & vbCrLf ' Rate
		 	
		 	i = i + 1
		 	
		Loop 
		
		oTempFile.Close ' Close the temp file
		oFile.Close
		
		Set oFile = oFSO.OpenTextFile(strFile,1,False)       ' Open file containing uploaded rates for reading
		Set oTempFile = oFSO.OpenTextFile(strTempFilePath,1) ' Open the temp file for reading
		' Now compare the files
		
		Do While Not oTempFile.AtEndOfStream
		
			sapline = Split(oTempFile.ReadLine,vbCrLf)
			fileline = Split(oFile.ReadLine,vbCrLf)
			column = Split(fileline(0),vbTab)
			
			
			
			If Not Replace((Trim(sapline(0)))," ","") = Replace((Trim(column(0) & column(1) & column(2))),vbTab,"") Then
			
				Exit Sub 
				
			End If 
			
			
		Loop
		
		
		numFilesVerified = numFilesVerified + 1
		oTempFile.Close ' Close the temp file
		oFile.Close     ' Close the rate file
		
	End Sub  	
	
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

	Private Sub KillPopups(s)
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
	
	Public Property Get FilesVerified
		FilesVerified = numFilesVerified
	End Property 
	
End Class 