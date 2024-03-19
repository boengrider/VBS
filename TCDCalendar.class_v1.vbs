Option Explicit

Dim oTCD
Set oTCD = New TCDCalendar
oTCD.AddTCD "2307","BOGUS"
oTCD.AddTCD "2207","BOGUS"
oTCD.AddTCD "2107","BOGUS"
oTCD.AddTCD "2007","BOGUS"
oTCD.AddTCD "1609","BOGUS"
oTCD.FindNonTCDDate(Date())


debug.WriteLine "Easter (" & Year(Date) & "): " & oTCD.IsTodayEaster("1904")
'debug.WriteLine "Today a TCD: " & oTCD.IsTodayTCD
'debug.WriteLine "First non TCD date: " & oTCD.FirstNonTCD
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
		If t_dict.Exists(Right("00" & Day(D),2) & Right("00" & Month(D),2)) Or Weekday(D) = 1 Or Weekday(D) = 7 Then 
			FindNonTCDDate = FindNonTCDDate((D - 1)) ' Recursive call to the GetLastNonTCDDate function
		Else
			t_FirstNonTCDDay = D
			Exit Function
		End If
	End Function 
	
	Public Function IsTodayTCD	
		If t_dict.Exists(Right("00" & Day(Date),2) & Right("00" & Month(Date),2)) Or Weekday(Date) = 1 Or Weekday(Date) = 7 Then 
			IsTodayTCD = True
			Exit Function
		End If 
		
		IsTodayTCD = False
	End Function
	
	Public Function IsTodayEaster(date__) 'D should be in format DDMM w/o year component
		Dim A, B, C, P, Q, M, N, D, E, days__,i__,easter__
		
		A = Year(Date) mod 19 'First, calculate the location of the year Y in the Metonic cycle
		B = Year(Date) mod 4  'Now, find the number of leap days according to Julian’s calendar.
		C = Year(Date) mod 7  'Then, let’s take into account that the non-leap year is one day longer than 52 weeks.
		
		P = Round(CDbl(Year(Date)) / 100.0)
		Q = Round(CDbl(13 + 8 * P) / 25.0)
		M = CInt(15 - Q + P - Round(P / 4)) mod 30
		N = CInt(4 + P - Round(P / 4)) mod 17
		
		D = CInt(19 * A + M) mod 30
		E = CInt(2 * B + 4 * C + 6 * D + N) mod 7
		
		days__ = CInt(22 + D + E)
		
		If D = 29 And E = 6  Then
			'1904YYYY
			easter__ = "1904"
		ElseIf D = 28 And E = 6  Then
			'1904YYYY
			easter__ = "1804"
		ElseIf days__ > 31 Then
			i__ = days__ - 31
			easter__ = Right("00" & (days__ - 31),2) & "04"
		Else
			easter__ Right("00" & days__,2) & "03"
		End If 
			 
		Debug.WriteLine
		Debug.WriteLine "Easter is: " & easter__ & Year(Date)
		
	End Function 
	
		
	' ----------- Getters and Setters ---------------- 
	Public Property Get Len
		Len = t_len
	End Property 
	
	Public Property Get FirstNonTCD
		FirstNonTCD = t_FirstNonTCDDay
	End Property 
	
	
End Class