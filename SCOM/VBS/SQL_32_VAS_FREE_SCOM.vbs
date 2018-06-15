Dim oAPI, oBag, objConnection, objRecordSet
Set oAPI = CreateObject("MOM.ScriptAPI")
Set oBag = oAPI.CreatePropertyBag()

Const adOpenStatic = 3
Const adLockOptimistic = 3

strSCOMSource = "SCOM_TECH_MONITOR1"
strSCOMUnHealthyID = Right(strSCOMSource,1) & "2"
strConnectionString = "Provider=SQLOLEDB.1;Password=password;Persist Security Info=True;User ID=username;Initial Catalog=master;Data Source=servername"
strSQL = "SELECT MAX(region_size_in_bytes)/1024 [MaxContigBlockSizeKB] FROM sys.dm_os_virtual_address_dump where region_state = 0x00010000"
intThreshold = "5120"
strState = "UNKNOWN"
intResult = 0
strStateHistoryFile = strSCOMSource & "_last_state.txt"
strStateCountFile = strSCOMSource & "_last_count.txt"
intMaxFailCount = 3

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check for previous state
If Not objFSO.FileExists(strStateHistoryFile) Then
	Set objFile = objFSO.CreateTextFile(strStateHistoryFile,True)
	strLastState = "UNKNOWN"
	objFile.Write strLastState
	objFile.Close
	Set objFile = Nothing
Else
	Set objFile = objFSO.OpenTextFile(strStateHistoryFile, 1)
	strLastState = Trim(objFile.ReadAll)
	objFile.Close
	Set objFile = Nothing
End If

' Check for previous count
If Not objFSO.FileExists(strStateCountFile) Then
	Set objFile = objFSO.CreateTextFile(strStateCountFile,True)
	objFile.Write 0
	objFile.Close
	Set objFile = Nothing
Else
	Set objFile = objFSO.OpenTextFile(strStateCountFile, 1)
	intFailCount = Trim(objFile.ReadAll)
	objFile.Close
	Set objFile = Nothing
End If

On Error Resume Next
objConnection.Open strConnectionString

objRecordSet.Open strSQL, objConnection, adOpenStatic, adLockOptimistic

For Each x In objRecordSet.fields
    intResult = Trim(x.value)
Next

If Abs(intResult) > Abs(intThreshold) Then
	strState = "GOOD"
	Set objFile = objFSO.CreateTextFile(strStateCountFile,True)
	objFile.Write 0
	objFile.Close
	Set objFile = Nothing
	Set objFile = objFSO.CreateTextFile(strStateHistoryFile,True)
	objFile.Write strState
	objFile.Close
	Set objFile = Nothing
	If strState <> strLastState Then
		LogSCOMHealthyEvent
	End If
Else
	intFailCount = intFailCount + 1
	Set objFile = objFSO.CreateTextFile(strStateCountFile,True)
	objFile.Write intFailCount
	objFile.Close
	Set objFile = Nothing
	If Abs(intFailCount) = Abs(intMaxFailCount) Then
		strState = "BAD"
		Set objFile = objFSO.CreateTextFile(strStateHistoryFile,True)
		objFile.Write strState
		objFile.Close
		Set objFile = Nothing
		If strState <> strLastState Then
			LogSCOMUnHealthyEvent
		End If
	Else
		strState = "NOTGOOD"
		Set objFile = objFSO.CreateTextFile(strStateHistoryFile,True)
		objFile.Write strState
		objFile.Close
		Set objFile = Nothing
	End If
End If

Call oBag.AddValue("State", strState)
If strState = "UNKNOWN" Then
	Call oBag.AddValue("Result", "Failed to run VAS query.")
Else
	Call oBag.AddValue("Result", intResult)
End If

objRecordset.Close
objConnection.Close

' SCOM Event creation procedures
Sub LogSCOMHealthyEvent
	Set WshShell = CreateObject("WScript.Shell")
	strCommand = "eventcreate /ID " & SCOMHealthyID & " /L APPLICATION /T INFORMATION /SO " & strSCOMSource & " /D " & Chr(34) & SCOMHealthyDescription & Chr(34)
	WshShell.Run strcommand
	Set WshShell = Nothing
End Sub

Sub LogSCOMUnHealthyEvent
	Set WshShell = CreateObject("WScript.Shell")
	strCommand = "eventcreate /ID " & SCOMUnHealthyID & " /L APPLICATION /T ERROR /SO " & strSCOMSource & " /D " & Chr(34) & SCOMUnHealthyDescription & Chr(34)
	WshShell.Run strcommand
	Set WshShell = Nothing
End Sub

Call oAPI.Return(oBag)

Set objFSO = Nothing
Set objRecordSet = Nothing
Set objConnection = Nothing
Set oBag = Nothing
Set oAPI = Nothing