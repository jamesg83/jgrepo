Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const HKEY_LOCAL_MACHINE = &H80000002
Const Quote = """"

strComputers = WScript.Arguments(0)
strComputersFile = Replace(strComputers,".txt","")

Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objTextFile = objFSO.OpenTextFile(strComputers, ForReading) 
strFile = objTextFile.ReadAll
arrComputers = Split(strFile,VbCrLf)
Set objTextFile = Nothing
Set objFSO = Nothing

strLogFile = "C$\Windows\CCM\Logs\AppEnforce.log"
strErrorString = "INSTALLATION: Cmd = NETSH AdvFirewall Firewall ADD Rule Name=" & Quote & "SQL Server 2012 Express TCP" & Quote & "DIR=IN Action=Allow Protocol=TCP Localport=1433"


strResultFile = "results.csv"


dtStart = Now

'Loop through all computers in the array
For Each strComputer in arrComputers
	strComputer = UCase(strComputer)
	strResult = strComputer
	
	'output some progress if running via cscript
	If UCase(Right(WScript.Fullname,11)) = "CSCRIPT.EXE" Then
		wscript.echo intComputerCount & "/" & intSizeOfArray & "  (" & strComputer & ")"
	End If
	
	'If computer responds to a ping run checks
	If PingStatus(strComputer) = "True" Then
		If UNCFolderExists(strComputer,"C$\Windows") Then
			'strLogFile exists
			If UNCFileExists(strComputer,strLogFile) = "True" Then
				strResult = strResult & "," & strLogFile & " exists"
			
				'BridgeRemote.log Performance counter disabled error found
				If UNCFileContent(strComputer,strLogFile,strErrorString) = "True" Then
					strResult = strResult & "," & strErrorString & " Found"
				Else
					strResult = strResult & ",Not found"
				End If
			
			Else
				strResult = strResult & "," & strLogFile & " doesn't exist"
			End If
		Else
			strResult = strResult & ",can't access $ share"
		End If
	Else
		'Can't ping
		strResult = strResult & ",Cannot ping computer"
	End If
	
	'Log result line if unable to ping
	WriteLog strResultFile,ForAppending,strResult

Next


'Functions

Function WriteLog(logfile,logmode,logentry)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objLogFile = objFSO.OpenTextFile(logfile, logmode, True)
	objLogFile.WriteLine(logentry) 
	objLogFile.Close
	Set objLogFile = Nothing
	Set objFSO = Nothing
End Function

Function PingStatus(computer)
	PingStatus = "False"
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colItems = objWMIService.ExecQuery ("Select * from Win32_PingStatus Where Address = '" & computer & "' and Timeout = 1000")
	For Each objItem in colItems
		If objItem.StatusCode = 0 Then 
			PingStatus = "True"
		End If
	Next
	Set colItems = Nothing
	Set objWMIService = Nothing
End Function

Function UNCFolderExists(computer,folder)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists("\\" & computer & "\" & folder & "") Then
		UNCFolderExists = "True"
	Else
		UNCFolderExists = "False"
	End If
	Set objFSO = Nothing
End Function

Function UNCFileExists(computer,file)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists("\\" & computer & "\" & file & "") Then
		UNCFileExists = "True"
	Else
		UNCFileExists = "False"
	End If
	Set objFSO = Nothing
End Function

Function UNCFileContent(computer,file,content)
	UNCFileContent = "False"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile("\\" & computer & "\" & file & "", ForReading, False)
	Do until objFile.AtEndOfStream
		strLine = objFile.ReadLine
		If InStr(strLine,content) Then
			UNCFileContent = "True"
		End If
	Loop
	Set objFile = Nothing
	Set objFSO = Nothing

End Function