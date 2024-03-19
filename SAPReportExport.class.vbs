Option Explicit
'DEBUG
Dim cliName__ : cliName__ = 105
Dim sysName__ : sysName__ = "FP2"
Dim companyCode__ : companyCode__ = "SK01"
'DEBUG
Dim ss,sl
Dim wsh : Set wsh = CreateObject("Wscript.Shell")
Set sl = New SAPLauncher
sl.SetClientName = cliName__
sl.SetSystemName = sysName__
sl.SetLocalXML = wsh.ExpandEnvironmentStrings("%APPDATA%") & "\SAP\Common\SAPUILandscape.xml"
sl.CheckSAPLogon
sl.FindSAPSession

If Not sl.SessionFound Then 
	WScript.Quit(1)
End If 

'Instantiate SAPReportExport class
Dim oSAPR : set oSAPR = new SAPReportExport
Set ss = sl.GetSession
'Initialize object passing a sap session object
oSAPR.Init ss
'Export report first
Dim retval : retval = oSAPR.ExportReport("/NZTC_Z_REPORT_FIAP", "SK01_LAT999", "SK01", "C:\!AUTO", "report.xlsx")
'Then call ExcelObject method to get the exported report workbook
Dim wb : Set wb = oSAPR.ExcelObject



Class SAPReportExport

	Private excelInstancesAtExport__
    Private objLocator__
	Private objService__ 
    Private sapSession__
	Private reportPath__
	Private reportName__
	
    Public Sub Class_Initialize
        sapSession__ = Null
    End Sub

    Public Function Init(ByRef sapSession)
	    Set objLocator__ = CreateObject("WbemScripting.SWbemLocator")
		Set objService__ = objLocator__.ConnectServer(".", "root\cimv2")
        Set sapSession__ = sapSession
    End Function

    Public Function ExportReport(ByVal sapTransaction, ByVal sapVariant, ByVal companyCode, ByVal reportPath, ByVal reportName)
        If Right(reportPath,1) <> "\" Then
       		reportPath__ = reportPath & "\"
       	Else
       		reportPath__ = reportPath
       	End If 
       
    	reportName__ = reportName
    	
        sapSession__.FindById("wnd[0]/tbar[0]/okcd").Text = sapTransaction
        sapSession__.FindById("wnd[0]").SendVKey 0  'Enter
        KillPopups sapSession__
        sapSession__.FindById("wnd[0]").SendVKey 17   'Shift+F5
        KillPopups sapSession__
        sapSession__.FindById("wnd[1]/usr/txtV-LOW").Text = ""
        sapSession__.FindById("wnd[1]/usr/txtV-LOW").Text = sapVariant
        sapSession__.FindById("wnd[1]/usr/ctxtENVIR-LOW").Text = ""
        sapSession__.FindById("wnd[1]/usr/txtENAME-LOW").Text = ""
        sapSession__.FindById("wnd[1]/usr/txtAENAME-LOW").Text = ""
        sapSession__.FindById("wnd[1]/usr/txtMLANGU-LOW").Text = ""
        sapSession__.FindById("wnd[1]").SendVKey 8
        KillPopups sapSession__
        
        On Error Resume Next
        Dim txtField : Set txtField = sapSession__.ActiveWindow.FindById("usr/txtMESSTXT1")
        
       
            If IsObject(txtField) Then 'text field found
                If InStr(txtField.Text, "Variant  does not exist") > 0 Then 'Variant does not exist
                    sapSession__.FindById("wnd[0]").SendVKey 0
                    sapSession__.FindById("wnd[0]/tbar[0]/okcd").Text = "/nex"
                    sapSession__.FindById("wnd[0]").SendVKey 0
                    ExportReport = -2
                    Exit Function
                End If
        End If
        
        'Variant exists. Continue
        sapSession__.FindById("wnd[0]/usr/radOPEN_CLR").Select    'option Currently Open Items
        sapSession__.FindById("wnd[0]/usr/ctxtBUKRS-LOW").Text = companyCode
        sapSession__.FindById("wnd[0]/usr/txtGJAHR-LOW").Text = Year(DateAdd("m", -1, Date))   'year of reporting period as string YYYY => year(today - 1 month)
        sapSession__.FindById("wnd[0]/usr/txtMONAT-LOW").Text = Right("00" & Month(DateAdd("m", -1, Date)), 2) 'month of reporting period as string MM => month(today - 1 month)
        sapSession__.FindById("wnd[0]/usr/btn%_ZUONR_%_APP_%-VALU_PUSH").Press    'click on the Multiple selection icon on the Allocation Number field
        KillPopups sapSession__
        sapSession__.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").Select    'click on tab Select Single Values
        sapSession__.FindById("wnd[1]").SendVKey 24    'click on Upload from Clipboard
        KillPopups sapSession__
        sapSession__.FindById("wnd[1]").SendVKey 8    'click on Confirm - F8
        KillPopups sapSession__
        sapSession__.FindById("wnd[0]").SendVKey 8  'report execution
        KillPopups sapSession__
        'Following label exists if report contains no data
        
        On Error Resume Next
        Dim label : Set label = sapSession__.ActiveWindow.FindById("usr/lbl[2,8]")
        
        If IsObject(label) Then
            If InStr(label.Text, "no data") < 0 Then
                sapSession__.FindById("wnd[0]/tbar[0]/okcd").Text = "/nex"
                sapSession__.FindById("wnd[0]").SendVKey 0
                ExportReport = -1
                Exit Function
            End If
        End If
        
		Set excelInstancesAtExport__ = objService__.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name LIKE 'excel' ")

        'List contains data. Continue
        sapSession__.FindById("wnd[0]/tbar[0]/okcd").Text = "&SELECT_SPREADSHEET"
        sapSession__.FindById("wnd[0]").SendVKey 0
        KillPopups sapSession__
        sapSession__.FindById("wnd[1]/usr/radRB_OTHERS").Select
        sapSession__.FindById("wnd[1]/usr/cmbG_LISTBOX").Key = "10"
        sapSession__.FindById("wnd[1]").SendVKey 0
        sapSession__.FindById("wnd[1]/usr/ctxtDY_PATH").Text = reportPath__
        sapSession__.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = reportName__
        sapSession__.FindById("wnd[1]").SendVKey 11
        sapSession__.FindById("wnd[0]").SendVKey 0
        sapSession__.FindById("wnd[0]/tbar[0]/okcd").Text = "/nex"
        sapSession__.FindById("wnd[0]").SendVKey 0
        
        ExportReport = 1 'All ok
    
    End Function

    Private Function KillPopups(ByRef objSession)
        Do While objSession.Children.Count > 1
            If InStr(objSession.ActiveWindow.Text, "System Message") > 0 Then
                objSession.ActiveWindow.SendVKey 12
            ElseIf InStr(objSession.ActiveWindow.Text, "Information") > 0 And InStr(objSession.ActiveWindow.PopupDialogText, "Exchange rate adjusted to system settings") > 0 Then
                objSession.ActiveWindow.SendVKey 0
            ElseIf InStr(objSession.ActiveWindow.Text, "Copyright") > 0 Then
                objSession.ActiveWindow.SendVKey 0
            ElseIf InStr(objSession.ActiveWindow.Text, "License Information for Multiple Logon") > 0 Then
                objSession.ActiveWindow.FindById("usr/radMULTI_LOGON_OPT2").Select
                objSession.ActiveWindow.SendVKey 0
            'ElseIF   'Insert next type of popup windows which you want to kill
            Else
                Exit Do
            End If
        Loop
    End Function
    
    Public Function ExcelObject
          
	    On Error Resume Next
        Dim excel : set Excel = GetObject(,"Excel.Application")
        
        Do While err.Number <> 0
          err.Clear
          set excel = GetObject(,"Excel.Application")
        Loop

        Dim wb
		Dim wbToReturn
        For each wb in excel.Workbooks
	      If wb.Name = reportName__ Then
		    wscript.echo "found"
	        Set wbToReturn = wb
	        Exit For  
	      End if 
        Next
        
        Set ExcelObject = wbToReturn
		  
    End Function

End Class


Class SAPLauncher
	
	Private oNET
	Private oXML
	Private oWSH
	Private oFSO
	Private oSAPGUI
	Private oAPP
	Private oCON
	Private oSES
	Private strGlobalURL
	Private strLocalLandscapePATH
	Private boolSAPRunning  		' Indicates whether SPA Logon is runniied files
	Private boolSessionFound		' Set to true if session was found or created
	Private strSSN 					' Sap System Name e.g FQ2
	Private strSCN  	    		' Sap Client Name e.g. 105
	Private strSSD          		' Sap System Description. This string is found in the local landscape xml and used to connect to the sap system


	
	
	
	' ============== Constructor & Destructor ==================
	Private Sub Class_Initialize
	
		Set oNET = CreateObject("Wscript.NEtwork")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set oFSO = CreateObject("scripting.filesystemobject")
		Set oWSH = CreateObject("wscript.shell")
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
		boolSessionFound = False
		

	End Sub
	
	Private Sub Class_Terminate
		
	End Sub
	
	' ============ P U B L I C  &  P R I V A T E   M E T H O D S  &   S U B R O U T I N E S ===========
	
	

	' ---------- CheckSAPLogon
	Public Sub CheckSAPLogon
		Dim waitPeriod,waitTurns,currentTurn
		waitPeriod = 500 ' miliseconds
		waitTurns = 5 
		currentTurn = 1
		
		On Error Resume Next
		
		Set oSAPGUI = GetObject("SAPGUI") ' This fails is saplogon is not running. We're connecting to the COM object not creating our own instance
		If err.number <> 0 And Not IsObject(oSAPGUI) Then
			debug.WriteLine "SAP logon is not running"
			oWSH.Run "saplogon.exe",2,False
			Set oSAPGUI = GetObject("SAPGUI")
			Do While Not IsObject(oSAPGUI) And currentTurn <= waitTurns
				debug.WriteLine "Waiting for sap logon"
				WScript.Sleep waitPeriod * currentTurn '1st time wait 500 ms, 2nd time wait 1000 ms etc...
				Set oSAPGUI = GetObject("SAPGUI")
				currentTurn = currentTurn + 1
			Loop
		End If
		
		On Error GoTo 0
		
		If Not IsObject(oSAPGUI) Then
			boolSAPRunning = False
			Exit Sub
		End If 	
		
		boolSAPRunning = True
	
	End Sub 
	
	
	
	' ---------- FindSAPSession
	Public Sub FindSAPSession
		Dim waitPeriod,waitTurns,currentTurn
		waitPeriod = 5000 ' miliseconds
		waitTurns = 5 ' 5 x 5000 = 20000 ms / 20 s
		currentTurn = 1
		
		If Not boolSAPRunning Then
			oSES = Null
			Exit Sub 
		End If  
		
		FindSAPSystemDescription
		
		If IsNull(strSSD) Then
			oSES = Null
			Exit Sub
		End If
		 
		Set oAPP = oSAPGUI.GetScriptingEngine
	
			
		For currentTurn = 1 To waitTurns
			Set oCON = oAPP.OpenConnection(strSSD,True,False) ' Open a new connection synchronously
			On Error Resume Next
			Set oSES = oCON.Children(0) ' Attach to the first session
			KillPopups(oSES)
			On Error GoTo 0
			
			If Not oSES.ActiveWindow.FindByName("sbar", "GuiStatusbar") Is Nothing Then
				 If InStr(oSES.ActiveWindow.FindByName("sbar", "GuiStatusbar").text, "Enter a valid SAP user or choose one from the list") > 0 Then
					oSES.ActiveWindow.findById("usr/txtRSYST-MANDT").text = strSCN
					oSES.ActiveWindow.findById("usr/txtRSYST-BNAME").text = oNET.UserName
					oSES.ActiveWindow.findById("usr/txtRSYST-LANGU").text = "EN"
					oSES.ActiveWindow.sendvkey.0
					KillPopups(oSES) ' In case of multiple connections
				End If 
			End If 
		
			KillPopups(oSES)
				
			If Not IsObject(oSES) Or IsNull(oSES) Then
				Debug.WriteLine "No session found, waiting " & currentTurn & " out of " & waitTurns & " turns"
				WScript.Sleep waitPeriod
			ElseIf IsObject(oSES) And IsNull(oSES) Then
				Debug.WriteLine "No session found, waiting " & currentTurn & " out of " & waitTurns & " turns"
				WScript.Sleep waitPeriod
			Else
				Exit For
			End If 	
		Next
		
		If Not IsObject(oSES) Or IsNull(oSES) Then
				Debug.WriteLine "No session found after 5 retries"
				boolSessionFound = False
				Exit Sub
		End If 
		
		If IsObject(oSES) And Not IsNull(oSES) Then
			If InStr(oSES.findById("wnd[0]/sbar/pane[0]").text,"No user exists") > 0 Then
				oCON.CloseConnection
				boolSessionFound = False 
				debug.WriteLine "Session found: " & boolSessionFound
				Exit Sub 
			End If
			
			boolSessionFound = True
			If InStr(oSES.ActiveWindow.FindByName("sbar", "GuiStatusbar").text, "Enter a valid SAP user or choose one from the list") > 0 Then
				oSES.ActiveWindow.findById("usr/txtRSYST-MANDT").text = strSCN
				oSES.ActiveWindow.findById("usr/txtRSYST-BNAME").text = oNET.UserName
				oSES.ActiveWindow.findById("usr/txtRSYST-LANGU").text = "EN"
				oSES.ActiveWindow.sendvkey.0
				KillPopups(oSES) ' In case of multiple connections
			End If 
			debug.WriteLine "Session found: " & boolSessionFound
			Exit Sub 
		End If
		
		oCON.CloseConnection
		boolSessionFound = False
		debug.WriteLine "Session found: " & boolSessionFound
			
	End Sub 

	
	
	' --------- FindSAPSystemDescription
	Private Sub FindSAPSystemDescription
	

		oXML.load(strLocalLandscapePATH) ' Locally stored XML
		oXML.setProperty "SelectionLanguage", "XPath"
		
		On Error Resume next
		strSSD =  oXML.selectSingleNode("//Landscape/Services/Service[starts-with(@name, '" & strSSN & "')]").attributes.getNamedItem("name").text
		On Error GoTo 0
		
		If Not IsNull(strSSD) then
			CheckSAPLogon
			Exit Sub  
		End If 
				
		strSSD = Null ' Not found
		
	End Sub 
	
	Public Function GetSession
	
		If IsNull(oSES) Or Not IsObject(oSES) Then
			GetSession = Null
		else
			Set GetSession = oSES
		End If 

	End Function
	
	Public Function KillPopups(ByRef objSession)
		Do While objSession.Children.Count > 1
			If InStr(objSession.ActiveWindow.Text, "System Message") > 0 Then
				objSession.ActiveWindow.sendVKey 12
			ElseIf InStr(objSession.ActiveWindow.Text, "Information") > 0 And InStr(objSession.ActiveWindow.PopupDialogText, "Exchange rate adjusted to system settings") > 0 Then
				objSession.ActiveWindow.sendVKey 0
			ElseIf InStr(objSession.ActiveWindow.Text, "Copyright") > 0 Then
				objSession.ActiveWindow.sendVKey 0
			ElseIf InStr(objSession.ActiveWindow.Text, "License Information for Multiple Logon") > 0 Then
				objSession.ActiveWindow.findById("usr/radMULTI_LOGON_OPT2").select
				objSession.ActiveWindow.sendVKey 0
			'ElseIF   'Insert next type of popup windows which you want to kill
			Else
				Exit Do
			End If
		Loop
	End Function 

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
	
	
	Public Property Get SAPsysDescription
		SAPsysDescription = strSSD
	End Property 
	
	Public Property Get GetGlobalURL
	
		GetGlobalURL = strGlobalURL
		
	End Property 
	
	Public Property Let SetGlobalURL(url)
	
		strGlobalURL = url
		
	End Property 
	
	Public Property Let SetLocalXML(xml)
	
		strLocalLandscapePATH = xml
		
	End Property 
	
	Public Property Get GetLocalXML
	
		GetLocalXML = strLocalLandscapePATH
		
	End Property 
	
	Public Property Let SetSystemName(sys)
	
		strSSN = UCase(sys)
		
	End Property 
	
	Public Property Let SetClientName(cli)
	
		strSCN = cli
	
	End Property 
	
	Public Property Get SessionFound
	
		SessionFound = boolSessionFound
		
	End Property 
	
		
	
End Class 