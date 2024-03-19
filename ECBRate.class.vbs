Option Explicit 

Dim r,oTCD,Shell,oDF
Set oDF = New DateFormatter
Set Shell = CreateObject("Wscript.Shell")
Set oTCD = New TCDCalendar
Set r = New ECBRate

' ---------- I N I T I A L I Z A T I O N -------------
oTCD.Init
oTCD.AddTCD "01012020","New Year's Day"
oTCD.AddTCD "10042020","Good Friday"
oTCD.AddTCD "13042020","Easter Monday"
oTCD.AddTCD "01052020","Labour Day"
oTCD.AddTCD "25122020","Christmas Day"
oTCD.AddTCD "26122020","Christmas Holiday"
oTCD.AddTCD "01012021","New Year's Day"
oTCD.AddTCD "02042021","Good Friday"
oTCD.AddTCD "05042021","Easter Monday"
oTCD.AddTCD "01052021","Labour Day"
oTCD.AddTCD "25122021","Christmas Day"
oTCD.AddTCD "26122021","Christmas Holiday"
oTCD.AddTCD "01012022","New Year's Day"
oTCD.AddTCD "15042022","Good Friday"
oTCD.AddTCD "18042022","Easter Monday"
oTCD.AddTCD "01052022","Labour Day"
oTCD.AddTCD "25122022","Christmas Day"
oTCD.AddTCD "26122022","Christmas Holiday"

oTCD.FindNonTCDDate(Date() -1)

debug.WriteLine "Today is a TCD: " & CStr(oTCD.IsTodayTCD)
debug.WriteLine "First non TCD date: " & oTCD.FirstNonTCD

r.Init Shell.ExpandEnvironmentStrings("%TEMP%") & "\" & CreateGUID & ".txt",oDF.ToYearMonthDayWithDashes(oTCD.FirstNonTCD),"C:\ExRate\SK01"
debug.WriteLine "Target date: " & r.Ddate
r.Fcurrs = "USD,CZK,DKK,GBP,HUF,PLN,SEK,CHF,NOK"
r.OverrideQuantity = "CZK,HUF"
r.AddOutputDirs = "\\vcn.ds.volvo.net\cli-sd\sd1294\046629\output\01_SK01_ExRateProcessing\SK01" ' Additional output dir
r.SetXMLUrl = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist-90d.xml"
r.CreateTEMPFile
r.ParseRateFile
r.MakeOutputFile
r.DeleteTEMPFile








Class ECBRate 

	Private oWSH
	Private oXML
	Private strPathToOutputDir
	Private strDate
	Private oHTTP
	Private dictRate
	Private dictOverrideQuantity
	Private oOutFile
	Private oTempFile
	Private oFSO
	Private strTCurr
	Private strAdditionalOutputDirs ' Copy output file(s) here
	Private strOutputFileName
	Private strPathToTempFile
	Private strRateURL
	Private strFCurrList
	Private boolMakeOutputFile
	Private errno
	Private dictAdjustRate
	Private listOutputFiles ' Pass this to the uploading script
	
	
	' ---------- Class Constructor ------------
	Private Sub Class_Initialize
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set oWSH = CreateObject("Wscript.Shell")
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set dictRate = CreateObject("Scripting.Dictionary")
		Set dictAdjustRate = CreateObject("Scripting.Dictionary")
		Set dictOverrideQuantity = CreateObject("Scripting.Dictionary")
		strRateURL = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist-90d.xml" ' 90 days. Also contains the last rate. No need to user daily XML
		strTCurr = "EUR"
		boolMakeOutputFile = False 
		strAdditionalOutputDirs = Null 
		strPathToOutputDir = Null
		strFCurrList = Null
		strDate = Null
	End Sub 
	
	' ---------- Class Destructor -------------
	Private Sub Class_Terminate
		
	End Sub
	
	' -------------------------------------------
	' --------- P u b l i c   M e t h o d s -----
	' -------------------------------------------
	
	' CreateTEMPFile
	' Method creates a unique temp file. 
	' If successfull it returns 0
	' On error it returns -1 and errno is set appropriately
	Public Function CreateTEMPFile
		If IsEmpty(strPathToTempFile) Or strPathToTempFile = "" Then
			boolMakeOutputFile = False 
			errno = 1 ' Can't create a temp file. No path to temp file was provided
			CreateTEMPFile = -1 ' No path provided. File cannot be created
			Exit Function
		End If
		
		Set oTempFile = oFSO.OpenTextFile(strPathToTempFile,2,True)
		Debug.WriteLine "Created a temp file"
		ClearIECache ' Clear the IE cache first
		Debug.WriteLine "Cleared the IE cache"
		oHTTP.open "GET",strRateURL & "?date=" & strDate,False ' Open a http connection to the cnb server
		oHTTP.send  ' send request
		
		
		' In case URL cannot be accessed, close the temp file and delete it. Return HTTP error code
		If oHTTP.status <> 200 Then
			oTempFile.Close
			oFSO.DeleteFile strPathToTempFile
			boolMakeOutputFile = False
			errno = 2 ' Rate file can't be downloaded from the CNB web
			CreateTEMPFile = -1 ' ERROR downloading rate file
			Exit Function
		End If 
		' Write content to the temp file
		oTempFile.Write oHTTP.responseText
		oTempFile.Close
		boolMakeOutputFile = True
		CreateTEMPFile = 0 ' File created successfully
	End Function 
	
	Public Function DeleteTEMPFile
		If Not IsNull(oTempFile) And oFSO.FileExists(strPathToTempFile) Then
			oTempFile.Close
			oFSO.DeleteFile strPathToTempFile
			debug.WriteLine "Deleted the temp file"
		End If 
	End Function
	
	' Init()
	' strPath -> Absolute path to the temp file
	' strUrl  -> Rate URL or Null
	' strFCurrs -> comma delimited list of the wanted currencies or Null
	' strTargetDate -> target date or Null
	' strPathToOD -> Absolute path to the directory where to put otput files locally e.g C:\ExRate\CZ02
	Public Function Init(strPath,strTargetDate,strPathToOD) 		
		strPathToTempFile = strPath
		If Not IsNull(strPathToOD) Then
			strPathToOutputDir = strPathToOD
		Else ' Try to recover and use C:\ExRate path
				strPathToOutputDir = oWSH.ExpandEnvironmentStrings("%SYSTEMDRIVE%") & "\ExRate"
		End If
	
		Call MakeOutputDir ' Create utput directory
		
		
		If Not IsNull(strTargetDate) Then
			strDate = strTargetDate
		End If
			
		
	End Function 
	
	' ClearIECache()
	Private Sub ClearIECache
		shell.Run "RunDLL32.exe InetCpl.cpl,ClearMyTracksByProcess 8",0,True
	End Sub 

	
	
		' ----------- ParseRateFile() ------------------------
	Public Function ParseRateFile() ' Call this method to add an entry into the dictionary. This function/method will do the parsing. It needs txt file with rates
		Dim key,column,line
	
		
		If IsNull(strDate) Then
			boolMakeOutputFile = False
			errno = 1 ' No Date was provided to the ParseRateFile() function
			debug.WriteLine "Error ocuder. Errno: " & errno
			ParseRateFile = -1 
			Exit Function
		End If 
   		
   		Set oTempFile = oFSO.OpenTextFile(strPathToTempFile,1,False)
   		line = oTempFile.ReadLine ' 1st line should contain date in DD.MM.YYYY  Compare it to our strDate
   		If Not strDate = Mid(line,1,10) Then
   			debug.WriteLine Mid(line,1,10)
   			oTempFile.Close
   			If oFSO.FileExists(strPathToTempFile) Then
   				oFSO.DeleteFile strPathToTempFile
   			End If 
   			boolMakeOutputFile = False
   			errno = 2 ' Bad date. Target date and downloaded file date dont match
   			debug.WriteLine "Error ocuder. Errno: " & errno
   			ParseRateFile = -1
   			Exit Function 
   		End If 
   		' ------ All OK continue parsing
   		Do While Not oTempFile.AtEndOfStream
   			
   			line = oTempFile.ReadLine ' 1st line should contain date in DD.MM.YYYY  Compare it to our strDate
   			column = Split(line,"|")
   			
   			For Each key In Split(strFCurrList,",")
   				
   				If key = column(3) Then 
  					
   					dictRate.Add column(3),column(4) ' RATE VALUE e.g. AUD 16,421
   						
   				End If 
   					
   			Next
   			
   		Loop
   		boolMakeOutputFile = True
   		debug.WriteLine "Output file created"
   	End Function
   	
   	
   	Function MakeOutputFile
   		Dim key
   		' Check if the OD exists
   		If Not boolMakeOutputFile Then
   			errno = 1 ' Cant continue making output file
   			MakeOutputFile = -1
   			Exit Function
   		End If 
   		
   		If Not IsNull(strPathToOutputDir) Then
   			
   			strOutputFileName = Right("0000" & Year(Date() + 1),4) & Right("00" & Month(Date() + 1),2) & Right("00" & Day(Date() + 1),2) & ".txt"
   			Set oOutFile = oFSO.OpenTextFile(strPathToOutputDir & "\" & strOutputFileName,2,True)
   			For Each key In Split(strFCurrList,",")
 	
   				' First write all rates FOREIGN to HOME currency 
   				If dictAdjustRate.Exists(key) Then  ' If key exists we divide rate by 100
   					
   					If dictOverrideQuantity.Exists(key) Then ' If key exists we override ratio in the output file too

   						
   						oOutFile.WriteLine key & vbTab & strTCurr & vbTab & FormatNumber((dictRate.Item(key) / dictAdjustRate.Item(key)),5) & vbTab & "1" & vbTab & "100"
   							
   					Else
   						
   						oOutFile.WriteLine key & vbTab & strTCurr & vbTab & FormatNumber((dictRate.Item(key) / dictAdjustRate.Item(key)),5) & vbTab & "1" & vbTab & "1"
   							
   					End If
   					
   				Else   
   					
   					If dictOverrideQuantity.Exists(key) Then
   						
   						oOutFile.WriteLine key & vbTab & strTCurr & vbTab & FormatNumber(dictRate.Item(key),5) & vbTab & "1" & vbTab & "100"
   							
   					Else 
   							
   						oOutFile.WriteLine key & vbTab & strTCurr & vbTab & FormatNumber(dictRate.Item(key),5) & vbTab & "1" & vbTab & "1"
   							
   					End If 
   						
   				End If  
   				
   			Next 
   			
   			For Each key In Split(strFCurrList,",")	
   							
   				' Now write HOME to FOREIGN currency
   				' FormatNumber(Round((1/ORIGINAL RATE FROM FILE),5) * 100,5)
   				If dictAdjustRate.Exists(key) Then 
   				
   					If dictOverrideQuantity.Exists(key) Then
   																		
   						oOutFile.WriteLine strTCurr & vbTab & key & vbTab & FormatNumber((Round(1 / dictRate.Item(key),5) * dictAdjustRate.Item(key)),5) & vbTab & "100" & vbTab & "1"
   							
   					Else 
   						
   						oOutFile.WriteLine strTCurr & vbTab & key & vbTab & FormatNumber((Round(1 / dictRate.Item(key),5) * dictAdjustRate.Item(key)),5) & vbTab & "1" & vbTab & "1"
   							
   					End If 
   						
   						
   				Else  
   					
   					If dictOverrideQuantity.Exists(key) Then
   	
   						oOutFile.WriteLine strTCurr & vbTab & key & vbTab & FormatNumber(Round((1/dictRate.Item(key)),5),5) & vbTab & "1" & vbTab & "100"
   							
   					Else
   						
   						oOutFile.WriteLine strTCurr & vbTab & key & vbTab & FormatNumber(Round((1/dictRate.Item(key)),5),5) & vbTab & "1" & vbTab & "1"
   							
   					End If
   						
   						
   				End If  
   				
   				
   			Next
   			
   			oOutFile.WriteLine "EUR" & vbTab & "SEK" & vbTab & FormatNumber(Round((dictRate("EUR") / dictRate("SEK")),5),5) & vbTab & "1" & vbTab & "1"
   			oOutFile.WriteLine "SEK" & vbTab & "EUR" & vbTab & FormatNumber(Round((1 / ((dictRate("EUR") / dictRate("SEK")))),5),5) & vbTab &  "1" & vbTab & "1"
   			oOutFile.Close
   			oTempFile.Close
   			DeleteTEMPFile 
   			listOutputFiles = strPathToOutputDir & "\" & strOutputFileName & ","
   			
   			CopyOutputFile 
   			
   			MakeOutputFile = 0
   			
   		End If 
   			
   		
   		
   	End Function 
   	
   	Private Sub CopyOutputFile
   		Dim item
   		If strAdditionalOutputDirs = "" Or IsNull(strAdditionalOutputDirs) Then
   			Exit Sub
   		End If 
   		
   		For Each item In Split(strAdditionalOutputDirs,",")
   		
   			If oFSO.FolderExists(item) Then
   	 			
   				oFSO.CopyFile strPathToOutputDir & "\" & strOutputFileName,item & "\" & strOutputFileName,True
   				
   			End If 
   			
   		Next
   		
   	End Sub 
   	
   	Private Function MakeOutputDir
	Dim comps,i,l,path 
	comps = Split(strPathToOutputDir,"\")
	l = UBound(comps) ' save len
	i = 0
	
	Do While Not i = l + 1
		
		If oFSO.GetDriveName(comps(i)) = comps(i) Then
			path = comps(i)
			i = i + 1
		End If
		
		path = path & "\" & comps(i)
		If Not oFSO.FolderExists(path) Then
			oFSO.CreateFolder path
		End If
		
		i = i + 1
		
	Loop 
	
End Function

   
   			
   	
   	' ----------------------------------------------
	' -------------- P r o p e r t i e s -----------
	' ----------------------------------------------
	Public Property Get RateURL
		RateURL = strRateURL
	End Property 
	
	Public Property Let SetXMLUrl(url)
		strRateURL = url
	End Property 
	
	Public Property Get PathToTempFile
		PathToTempFile = strPathToTempFile
	End Property
	
	Public Property Get Tcurr 
		Tcurr = strTCurr
	End Property 
	
	Public Property Let Tcurr(strCurr)
		strTCurr = strCurr
	End Property 
	
	Public Property Get Fcurrs
		Fcurrs = strFCurrList
	End Property 
	
	Public Property Let Fcurrs(strCurrs)
		strFCurrList = strCurrs
	End Property 
	
	Public Property Get Ddate 
		Ddate = strDate
	End Property 
	
	Public Property Get GetRate(strCurrency)
		GetRate = dictRate.Item(strCurrency)
	End Property 
	
	Public Property Let OverrideQuantity(strQuantity)
		Dim key
		For Each key In Split(strQuantity,",")
			dictOverrideQuantity.Add key,100
		Next
	End Property 
	
	Public Property Get ErrorCode
		ErrorCode = errno
	End Property 
	
	Public Property Let AdjustRate(strRates) ' comma delimited string
		Dim key
		For Each key In Split(strRates,",")
			dictAdjustRate.Add key,100
		Next
	End Property 
	
	Public Property Let AddOutputDirs(strDirs)
		strAdditionalOutputDirs = strDirs
	End Property 
	
	Public Property Get OutputFiles
		OutputFiles = Left(listOutputFiles,Len(listOutputFiles) - 1)
	End Property 
		
	
	
End Class 
	 
	 
	 
' --------------- F U N C T I O N S ----------------

' -------------- CreateGUID() -----------------------
	
	Function CreateGUID
  		Dim TypeLib
  		Set TypeLib = CreateObject("Scriptlet.TypeLib")
  		CreateGUID = Mid(TypeLib.Guid, 2, 36)
	End Function
	
	
	
	
	
	
' -------------- TEMPORARY STUFF .... DELETE 

' TCDCalendar Class
Class TCDCalendar
	' ------------- Private members ----------- 
	Private t_dict ' Scripting.Dictionary that holds TCD entries
	Private t_len ' t_dict.Count
	Private t_isDateTCD ' this variable is set to True if current item in the dictionary is a TCD. Usefull during iterations through the dictionary
	Private t_FirstNonTCDDay
	' ------------- Constructor ----------------------
	Public Function Init
		Set t_dict = CreateObject("Scripting.Dictionary")
		t_len = t_dict.Count
		t_FirstNonTCDDay = Null
		t_isDateTCD = False ' Initially false. GetLastNonTCDDate(D) uses this variable
	End Function
	' ------------ Instance methods ------------------
	Public Function AddTCD(ddmmyyyy,holiday_name)
		t_dict.Add ddmmyyyy,holiday_name
		t_len = t_dict.Count
	End Function 
	
	Public Function FindNonTCDDate(D) ' Argument is a Date Object
		If t_dict.Exists(Right("00" & Day(D),2) & Right("00" & Month(D),2) & Right("0000" & Year(D),4)) Or Weekday(D) = 1 Or Weekday(D) = 7 Then 
			FindNonTCDDate = FindNonTCDDate((D - 1)) ' Recursive call to the GetLastNonTCDDate function
		Else
			t_FirstNonTCDDay = D
			Exit Function
		End If
	End Function 
	
	Public Function IsTodayTCD
		Dim key
		For Each key In t_dict.Keys
				If key = (Right("00" & Day(Date()),2) & Right("00" & Month(Date()),2) & Right("0000" & Year(Date()),4)) Then 
				IsTodayTCD = True
				Exit Function
			End If 
		Next
		
		IsTodayTCD = False
	End Function
	
		
	' ----------- Getters and Setters ---------------- 
	Public Property Get Len
		Len = t_len
	End Property 
	
	Public Property Get FirstNonTCD
		FirstNonTCD = t_FirstNonTCDDay
	End Property 
	
	
End Class




Class DateFormatter
	' Convert from YYYY-MM-DD to DD.MM.YYYY
	Public Function FromYyyyMmDd_WithDashesTo_DdMmYyyy_WithDots(strDate)
		Dim temp
		temp = Right(strDate,2) & "." ' Day
		temp = temp & Mid(strDate,6,2) & "." ' Month
		temp = temp & Left(strDate,4) ' Year
		FromYyyyMmDd_WithDashesTo_DdMmYyyy_WithDots = temp
	End Function 
	
	Public Function ToYearMonthDay(D) ' Date object
		ToYearMonthDay = Right("0000" & Year(D),4) & Right("00" & Month(D),2) & Right("00" & Day(D),2)
	End Function
	
	Public Function ToYearMonthDayWithDashes(D)
		ToYearMonthDayWithDashes = Right("0000" & Year(D),4) & "-" & Right("00" & Month(D),2) & "-" & Right("00" & Day(D),2)
	End Function
	
	Public Function ToDayMonthYearWithDots(D)
		ToDayMonthYearWithDots = Right("00" & Day(D),2) & "." & Right("00" & Month(D),2) & "." & Right("0000" & Year(D),4)
	End Function
	
	
End Class 