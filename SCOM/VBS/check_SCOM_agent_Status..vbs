HEALTHCARE\svc_RegSCOM_UAT''*********************************************************************
' Script by Shiva Kumar *
' To Find out if SCOM Agent is installed as a service on a remote server *
' and if it is Started/Stopped *
'*********************************************************************
 
Set fSO = CreateObject("Scripting.FileSystemObject")
 
'**********************************************************************
' servers.txt will contain the list of servers where the check has to run against.
'**********************************************************************
Set strInp = fso.OpenTextFile("servers.txt",1)
 
'**********************************************************************
' scom_log.txt will contain the result of the check.
'**********************************************************************
Set strLog = fso.CreateTextFile("SCOM_Log.txt")
 
set WshShell = CreateObject("wscript.Shell")
 
Do Until strInp.AtEndOfStream
 
strComputer = Trim(strInp.ReadLine)
 
On Error Resume Next
 
Set objPing = WshShell.Exec("Ping -n 1 " & strComputer)
 
strObjPingOut = objPing.StdOut.ReadAll
 
If InStr(strObjPingOut,"Lost = 0") Then
 
	Wscript.Echo " Now Processing :" & strComputer
 
	strService = "System Center Management"
 
	Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 
	Set colListOfServices = objWMIService.ExecQuery _
	("Select * from Win32_Service where DisplayName = '" & strService & "'")
 
	nItems = colListOfServices.Count
 
	If nItems > 0 Then
 
		For Each objService in colListOfServices
 
			If objService.State = "Stopped" Then
 
				strLog.writeline objService.DisplayName & " Installed/Stopped on " & strComputer
 
			ElseIf objService.State = "Running" Then
 
				strLog.writeline objService.DisplayName & " Installed/Running on " & strComputer
 
			End If
		Next
 
	Else
 
		strLog.writeline "Service Not Installed on " & strComputer
 
	End If
 
Else
	Wscript.Echo " Could Not Ping :" & vbtab & strComputer
	strLog.writeline " Could Not Ping :" & vbtab & strComputer
End If
 
Loop
 
strLog.Close
