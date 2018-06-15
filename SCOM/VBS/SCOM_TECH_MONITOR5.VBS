' Check the last modified date/time of the "D:\PDRWebService\In\Pass" folder.
' If it's over the configured threshold then we haven't received any new files from the PDR web service in that time so send a SCOM alert.
' This script records it's state (GOOD/BAD) and will only alert when the state changes.

' Get computer name
Set WshShell = CreateObject("WScript.Shell")
strComputerName = wshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
Set WshShell = Nothing

'---------------------------------------------------------------------
' Script config
'---------------------------------------------------------------------
strFolderPath = "D:\Orion Health\SopranoMedicalTemplates\exportSMT$"
intAgeThreshold = 4 'Hours
SCOMSource = "SCOM_TECH_MONITOR5"
SCOMHealthyDescription = "RESOLVED: The folder " & strFolderPath & " on server " & strComputerName & " has been modified."
SCOMUnHealthyDescription = "WARNING: The folder " & strFolderPath & " on server " & strComputerName & " hasn't been modified in the last " & intAgeThreshold & " hours."

' You shouldn't need to modify these 3
SCOMHealthyID = Right(SCOMSource,1) & "1"
SCOMUnHealthyID = Right(SCOMSource,1) & "2"
strStateHistoryFile = SCOMSource & "_last_state.txt"

'---------------------------------------------------------------------
' Script guts
'---------------------------------------------------------------------

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

' Get folder last modified date/time
Set objFolder = objFSO.GetFolder(strFolderPath)
dtLastModified = objFolder.DateLastModified
Set objFolder = Nothing

' See how long ago folder was modified
dtFolderAge = DateDiff("h",dtLastModified,Now)

' If older then the set threshold set the current state to bad otherwise good
If dtFolderAge >= intAgeThreshold Then
	strCurrentState = "BAD"
Else
	strCurrentState = "GOOD"
End If

' If the current state is different to the previous state then create SCOM event log entry
If Not strCurrentState = strLastState Then
	If strCurrentState = "BAD" Then
		LogSCOMUnHealthyEvent
	Else
		If Not strLastState = "UNKNOWN" Then
			LogSCOMHealthyEvent
		End If
	End If
End If

' Clean up
Set objFile = objFSO.CreateTextFile(strStateHistoryFile,True)
objFile.Write strCurrentState
objFile.Close
Set objFile = Nothing
Set objFSO = Nothing

'---------------------------------------------------------------------
' SCOM Event creation procedures
Sub LogSCOMHealthyEvent
	Set WshShell = CreateObject("WScript.Shell")
	strCommand = "eventcreate /ID " & SCOMHealthyID & " /L APPLICATION /T INFORMATION /SO " & SCOMSource & " /D " & Chr(34) & SCOMHealthyDescription & Chr(34)
	WshShell.Run strcommand
	Set WshShell = Nothing
End Sub

Sub LogSCOMUnHealthyEvent
	Set WshShell = CreateObject("WScript.Shell")
	strCommand = "eventcreate /ID " & SCOMUnHealthyID & " /L APPLICATION /T ERROR /SO " & SCOMSource & " /D " & Chr(34) & SCOMUnHealthyDescription & Chr(34)
	WshShell.Run strcommand
	Set WshShell = Nothing
End Sub
