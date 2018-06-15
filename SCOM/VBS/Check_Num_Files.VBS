' Check the number of files in a folder.
' If there are more than 500 then send a SCOM alert.
' This script records it's state (GOOD/BAD) and will only alert when the state changes.

' Get computer name
Set WshShell = CreateObject("WScript.Shell")
strComputerName = wshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
Set WshShell = Nothing

'---------------------------------------------------------------------
' Script config
'---------------------------------------------------------------------
strFolderPath = "D:\Staff\Error"
intCountThreshold = 0
strSCOMSource = "SCOM_TECH_MONITOR1"
strSCOMHealthyDescription = "RESOLVED: The folder " & strFolderPath & " on server " & strComputerName & " has less than " & intCountThreshold & " files in it."
strSCOMUnHealthyDescription = "WARNING: The folder " & strFolderPath & " on server " & strComputerName & " has more than " & intCountThreshold & " files in it."

' You shouldn't need to modify these 3
strSCOMHealthyID = Right(strSCOMSource,1) & "1"
strSCOMUnHealthyID = Right(strSCOMSource,1) & "2"
strStateHistoryFile = strSCOMSource & "_last_state.txt"

'---------------------------------------------------------------------
' Script guts
'---------------------------------------------------------------------

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check previous state
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

' Get number of files in the folder
Set objFolder = objFSO.GetFolder(strFolderPath)
Set objFiles = objFolder.Files
intFileCount = CInt(objFiles.Count)
Set objFiles = Nothing
Set objFolder = Nothing

' If more than the threshold set current state to bad otherwise good
If intFileCount >= intCountThreshold Then
	strCurrentState = "BAD"
Else
	strCurrentState = "GOOD"
End If

' If the current state is different to the previous state
If Not strCurrentState = strLastState Then
	' Write out the new state
	Set objFile = objFSO.CreateTextFile(strStateHistoryFile,True)
	objFile.Write strCurrentState
	objFile.Close
	' Create SCOM event log entries
	If strCurrentState = "BAD" Then
		LogSCOMUnHealthyEvent
	Else
		If Not strLastState = "UNKNOWN" Then
			LogSCOMHealthyEvent
		End If
	End If
End If

' Clean up
Set objFile = Nothing
Set objFSO = Nothing

'---------------------------------------------------------------------
' SCOM Event creation procedures

Dim strCommand

Sub LogSCOMHealthyEvent
	Set WshShell = CreateObject("WScript.Shell")
	strCommand = "eventcreate /ID " & strSCOMHealthyID & " /L APPLICATION /T INFORMATION /SO " & strSCOMSource & " /D " & Chr(34) & strSCOMHealthyDescription & Chr(34)
	WshShell.Run strcommand
	Set WshShell = Nothing
End Sub

Sub LogSCOMUnHealthyEvent
	Set WshShell = CreateObject("WScript.Shell")
	strCommand = "eventcreate /ID " & strSCOMUnHealthyID & " /L APPLICATION /T ERROR /SO " & strSCOMSource & " /D " & Chr(34) & strSCOMUnHealthyDescription & Chr(34)
	WshShell.Run strcommand
	Set WshShell = Nothing
End Sub
