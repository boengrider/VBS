Class StopWatch

	Private start
	Private finish
	
	Private Sub Class_Initialize
	
		start = 0
		finish = 0
		
	End Sub
	
	Private Sub Class_Terminate
	
	End Sub 
	

	Public Function Activate
		start = Hour(Now()) * 3600 + Minute(Now()) * 60 + Second(Now())
	End Function
	
	Public Function Deactivate
		finish = Hour(Now()) * 3600 + Minute(Now()) * 60 + Second(Now())
	End Function 
	
	
	Public Property Get Duration ' Seconds
		Duration = finish - start
	End Property 
	
	
End Class