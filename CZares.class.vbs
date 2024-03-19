Class CZARES

	Private objHttp
	Private objXml
	Private strUrl
	Private aRetVals(9)
	
	Private Sub Class_Initialize()
	
		' strUrl1 + ICO + strUrl2
		strUrl = "http://wwwinfo.mfcr.cz/cgi-bin/ares/darv_bas.cgi?ico="
		Set objHttp = CreateObject("msxml2.xmlhttp")
		Set objXml = CreateObject("msxml2.domdocument")
		
	End Sub 
	
	Public Function GetData(strICO)
	
		With objHttp
			.Open "GET", strUrl & strICO & "&xml=1&aktivni=false", False
			'.SetRequestHeader "Accept","application/xml;Charset=ISO-8859-1"
			.Send
		End With 
		
		If objHttp.status = "400" Then
			debug.WriteLine "Ares response code 400 -> Bad request"
			debug.WriteLine "Detailed response -> " & objHttp.responseText
			debug.WriteLine "Skipping further processing of this entry"
			GetData = 400
			Exit Function
		End If 
		If objHttp.status = "404" Then
			debug.WriteLine "Ares response code 404 -> Not found"
			debug.WriteLine "Detailed response -> " & objHttp.responseText
			debug.WriteLine "Skipping further processing of this entry"
			GetData = 404
			Exit Function
		End If 
'		debug.WriteLine "Extract info from Ares response"
		debug.WriteLine objHttp.responseText
		On Error Resume Next 
		With objXml
			.loadXML objHttp.responseText
			.setProperty "SelectionNamespaces","xmlns:D=""http://wwwinfo.mfcr.cz/ares/xml_doc/schemas/ares/ares_datatypes/v_1.0.3"""
		End With 
		
		aRetVals(0) = objXml.getElementsByTagName("D:DIC").item(0).text ' dic
		aRetVals(1) = objXml.getElementsByTagName("D:ICO").item(0).text ' ico
		aRetVals(2) = objXml.getElementsByTagName("D:OF").item(0).text ' name
		aRetVals(3) = objXml.getElementsByTagName("D:NU").item(0).text ' street
		aRetVals(4) = objXml.getElementsByTagName("D:CD").item(0).text ' cislo 
		aRetVals(5) = objXml.getElementsByTagName("D:N").item(0).text ' mesto
		aRetVals(6) = objXml.getElementsByTagName("D:PSC").item(0).text ' psc
		aRetVals(7) = objXml.getElementsByTagName("D:NS").item(0).text ' country
		aRetVals(8) = objXml.getElementsByTagName("D:PSU").item(0).text ' additional subject info / priznaky subjektu
		aRetVals(9) = objXml.getElementsByTagName("D:TCD").item(0).text ' domovne cislo
		
		GetData = 1
	End Function 
	
	Public Property Get vat
		vat = aRetVals(0)
	End Property 
	
	Public Property Get ico
		ico = aRetVals(1)
	End Property 
	
	Public Property Get name
		name = aRetVals(2)
	End Property 
	
	Public Property Get address
		address = aRetVals(3) & " " & aRetVals(9) & "/" & aRetVals(4)
	End Property 
	
	Public Property Get psc
		psc = aRetVals(6) 
	End Property
	
	Public Property Get city
		city = aRetVals(5)
	End Property
	
	Public Property Get country
		country = aRetVals(7)
	End Property 
	
	Public Property Get SubjectValidInPublicReg
		Dim val : val = LCase(Mid(aRetVals(8),2,1))
		
		Select Case val
		
			Case "a" ' Active registration
				SubjectValidInPublicReg = 1
				Exit Property 
				
			Case "z","h" ' Expired registration, expired registration more than 4 years
				SubjectValidInPublicReg = 0
				Exit Property
				
			Case "n" ' Not found in the database
				SubjectValidInPublicReg = 2
				Exit Property 
				
			Case Else ' Something else with no particular meaning
				SubjectValidInPublicReg = -1
				Exit Property 
				
		End Select 
	End Property 
		
	
			
End Class
