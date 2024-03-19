Class SlovenskoDigital

	Private aRetVals
	Private numFields
	Private strToken 
	Private strUrl
	Private oHTTP 
	Private oScriptControl
	Private oRxDeleteWhiteSpace
	Private oRxIsValid
	
	Private Sub Class_Initialize()
		Set oRxDeleteWhiteSpace = New RegExp
		oRxDeleteWhiteSpace.Pattern = "\s+"
		oRxDeleteWhiteSpace.Global = True
		Set oRxIsValid = New RegExp
		oRxIsValid.Pattern = "^([0-9]{8}|[0-9]{6})$"
		oRxIsValid.Global = True
		numFields = 9
		ReDim aRetVals(numFields)
	    strToken = "Token eab9a6995881ba03778a8803fd05888b0acf85c7879cef9c28a43f558a43a6c655d6d2b6d09cd89d"
		strUrl = "https://datahub.ekosystem.slovensko.digital/api/datahub/corporate_bodies/search?q=cin:"
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oScriptControl = CreateObject("ScriptControl")
		oScriptControl.Language = "JScript"
	End Sub 
	
	Private Sub Class_Terminate()
		Set oHTTP = Nothing
		Set oScriptControl = Nothing
	End Sub 
	
	Public Function DeleteWhiteSpace(strIco)
		DeleteWhiteSpace = oRxDeleteWhiteSpace.Replace(strIco,"")
	End Function 
	
	Public Function IsFormatValid(strIco)
		IsFormatValid = oRxIsValid.Test(strIco)
	End Function 
	
	Public Function GetData(strIco)
		debug.WriteLine "Sending SKdigital request for ICO " & strIco
		Set oScriptControl = CreateObject("ScriptControl")
		With oHTTP
			.Open "GET", strUrl & strIco, False
			.SetRequestHeader "Authorization", strToken
			.Send
		End With 
		
		If oHTTP.Status = "400" Then
			debug.WriteLine "SKdigital response code 400 -> Bad request"
			debug.WriteLine "Detailed response -> " & oHTTP.ResponseText
			debug.WriteLine "Skipping further processing of this entry"
			GetData = 400 ' Bad request
			Exit Function
		End If 
		If oHTTP.Status = "404" Then
			debug.WriteLine "SKdigital response code 404 -> Not found"
			debug.WriteLine "Detailed response -> " & oHTTP.ResponseText
			debug.WriteLine "Skipping further processing of this entry"
			GetData = 404 ' Not found
			Exit Function 
		Else 
			debug.WriteLine "SKdigital response code 1 -> OK"
			debug.WriteLine "Detailed response -> " & oHTTP.ResponseText
			GetData = 1 ' Found
		End If 
		debug.WriteLine "Extract info from SKdigital response"
		On Error Resume next
		With oScriptControl
			.Language = "JScript"
			With .Eval("(" + oHTTP.ResponseText + ")")
				aRetVals(0) = .cin
				aRetVals(1) = .vatin
				aRetVals(2) = .name
				aRetVals(3) = .street
				aRetVals(4) = .reg_number
				aRetVals(5) = .building_number
				aRetVals(6) = .postal_code
				aRetVals(7) = .municipality
				aRetVals(8) = .country
			End With
		On Error GoTo 0
		End With 
		
	End Function 
	
	Public Property Get ico
		ico = aRetVals(0)
	End Property
	
	Public Property Get vat
		vat = aRetVals(1)
	End Property 
	
	Public Property Get name
		name = aRetVals(2)
	End Property
	
	Public Property Get address
		address = aRetVals(3) & " " & aRetVals(4) & "/" & aRetVals(5)
	End Property 
	
	Public Property Get psc
		psc = aRetVals(6)
	End Property 
	
	Public Property Get city
		city = aRetVals(7)
	End Property 
	
	Public Property Get country
		country = aRetVals(8)
	End Property 
	
End Class 