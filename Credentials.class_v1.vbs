Option Explicit


Dim MyCredentials : Set MyCredentials = New Credentials
If MyCredentials.GetCredentials("sapbi") = 0 Then
	debug.WriteLine "Credentials not found"
	WScript.Quit
End If 

debug.WriteLine "User name: " & MyCredentials.Username
debug.WriteLine "Password: " & MyCredentials.Password
debug.WriteLine "Host: " & MyCredentials.Host
debug.WriteLine "Domain: " & MyCredentials.Domain


Class Credentials

	Private connectionString__
	Private connection__
	Private recordset__
	Private username__ 
	Private password__ 
	Private host__
	Private domain__
	
	Private Sub  Class_Initialize
	
		username__ = ""
		password__ = ""
		host__ = ""
		domain__ = ""
		
		connectionString__ = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=1;RetrieveIds=Yes;" & _
							 "DATABASE=https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it/CREDENTIALS;" & _
						 	 "LIST=CREDENTIALS;"
						   
		Set connection__ = CreateObject("Adodb.Connection")
		Set recordset__ = CreateObject("Adodb.Recordset")
		
		connection__.ConnectionString = connectionString__
		connection__.Open
		
	End Sub
	
	Public Function GetCredentials(resourceName)
	
		Dim orx__ : Set orx__ = New RegExp
		orx__.Global = True 
		Dim records__
		orx__.Pattern = "^(?:https?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/\n?]+)"
		
		recordset__.Open "SELECT Host,Username,Password FROM [CREDENTIALS] WHERE Title='" & resourceName & "';", connection__, 3, 3
		
		If recordset__.EOF Or recordset__.BOF Then
			GetCredentials = 0
			Exit Function 
		End If 
		
		recordset__.MoveFirst
		domain__ = orx__.Execute(recordset__.Fields("Host").Value)(0)
		orx__.Pattern = "(http:\/\/|https:\/\/)"
		
		domain__ = orx__.Replace(domain__,"")
		username__ = recordset__.Fields("Username").Value
		password__ = recordset__.Fields("Password").Value
		host__ = recordset__.Fields("Host").Value
		
		recordset__.Close
		
		GetCredentials = 1
	
	End Function 
	
	Public Property Get Password
	
		Password = password__
		
	End Property
	
	Public Property Get Username
	
		Username = username__
		
	End Property
	
	Public Property Get Host
	
		Host = host__
		
	End Property
	
	Public Property Get Domain
	
		Domain = domain__
		
	End Property 
	 

End Class
