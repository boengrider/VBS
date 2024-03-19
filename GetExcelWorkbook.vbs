'Function to catch specific excel workbook opened by SAP command &SELECT_SPREDSHEET or option Export to Spreadsheet.
'Function retries 5 times and waits 500 msec each retry for the specific workbook name to open
'Function checks only one instance and if the report is opened in other instance it will not catch. Suggest the user to close all excel application before running the script again.

'--------------- MAIN ---------------

Dim oWB
Dim report_name : report_name = "report.xlsx"
If GetExcelWorkbook(report_name,oWB) Then 
   Wscript.echo "Workbook found"
   Wscript.echo oWB.Name & " " & oWB.fullname
   'at this point the oWB is Excel object with the specific workbook
   'intellisense will not work in VBSEDIT since references are not active - either close and open workbook object again or use dummy object while dev
Else 
   wscript.echo "Workbook not found"
   'report was not found as open so the script should not continue
   WScript.Quit
   
End If 

'--------------- MAIN END ---------------


Function GetExcelWorkbook(ByVal reportName, ByRef outWorkbook)
          
	  On Error Resume Next
      Dim waitTime : waitTime = 500
      Dim waitTurns : waitTurns = 5
      Dim turn : turn = 0

      err.Clear
      Dim excel : set Excel = GetObject(,"Excel.Application")
    
       Do While err.Number <> 0
         If turn > waitTurns Then
            Exit Do
         End if  

         wscript.sleep waitTime
         turn = turn + 1

         err.Clear
         set excel = GetObject(,"Excel.Application")
       Loop

      If Not isobject(excel) Then
          GetExcelWorkbook = 0
          Exit Function
      End If 

      'Excel instance exists
      Dim workbook__
     
      For each workbook__ in excel.Workbooks
        if workbook__.Name = reportName Then
          set outWorkbook = workbook__
          GetExcelWorkbook = 1
          Exit Function
        end if 
      Next

      GetExcelWorkbook = 0
      
End Function
