Option Explicit

Dim oTCD
Set oTCD = New TCDCalendar
oTCD.AddTCD "23072020","BOGUS"
oTCD.AddTCD "22072020","BOGUS"
oTCD.AddTCD "21072020","BOGUS"
oTCD.AddTCD "20072020","BOGUS"

oTCD.FindNonTCDDate(Date())
debug.WriteLine "Today a TCD: " & oTCD.IsTodayTCD
debug.WriteLine "First non TCD date: " & oTCD.FirstNonTCD
' TCDCalendar Class
Class TCDCalendar
	' ------------- Private members ----------- 
	Private t_dict ' Scripting.Dictionary that holds TCD entries
	Private t_len ' t_dict.Count
	Private t_isDateTCD ' this variable is set to True if current item in the dictionary is a TCD. Usefull during iterations through the dictionary
	Private t_FirstNonTCDDay
	
	' ------------- Constructor ----------------------
	
	Private Sub Class_Initialize
		Set t_dict = CreateObject("Scripting.Dictionary")
		t_len = t_dict.Count
		t_FirstNonTCDDay = Null
		t_isDateTCD = False ' Initially false. GetLastNonTCDDate(D) uses this variable
	End Sub 
		
	Private Sub Class_Terminate
	
	End Sub 
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
		If Weekday(Date()) = 1 Or Weekday(Date()) = 7 Then 
			IsTodayTCD = True
			Exit Function
		End If 
		
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