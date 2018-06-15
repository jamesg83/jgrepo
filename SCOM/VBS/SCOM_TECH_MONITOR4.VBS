' Find the age of the oldest file in a folder.
' If it's over the configues age then send a SCOM alert.
' This script records it's state (GOOD/BAD) and will only alert when the state changes.

' Get computer name
Set WshShell = CreateObject("WScript.Shell")
strComputerName = wshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
Set WshShell = Nothing

'---------------------------------------------------------------------
' Script config
'---------------------------------------------------------------------
strFolderPath = "D:\hlink\FF_in\LAB"
' For strAgeInterval select one of the following: yyyy (Year), q (Quarter), m (Month), y (Day of year), d (Day), w (Weekday), ww (Week of year), h (Hour), n (Minute), s (Second)
strAgeInterval = "h"
intAgeThreshold = 4
strSCOMSource = "SCOM_TECH_MONITOR1"
strSCOMHealthyDescription = "RESOLVED: There are no files in the folder " & strFolderPath & " on server " & strComputerName & " older than " & intAgeThreshold & " hours."
strSCOMUnHealthyDescription = "WARNING: There are files in the folder " & strFolderPath & " on server " & strComputerName & " that are older than " & intAgeThreshold & " hours."

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

' Iterate through any files in the folder and find the one with oldest modified date/time
strOldestFile = "NotFound"
dtmOldestFileDate = Now
Set objFolder = objFSO.GetFolder(strFolderPath)
Set colFiles = objFolder.Files
For Each objFile in colFiles
    If objFile.DateLastModified < dtmOldestFileDate Then
        strOldestFile = objFile.Path
		dtmOldestFileDate = objFile.DateLastModified
    End If
Next
Set colFiles = Nothing
Set objFolder = Nothing

' If there aren't any files in the folder then just quit
If strOldestFile = "NotFound" Then
	Wscript.Quit
End If

' See how long ago the oldest file found was modified
dtFileAge = DateDiff(strAgeInterval,dtmOldestFileDate,Now)

' If older then the set threshold set the current state to bad otherwise good
If dtFileAge >= intAgeThreshold Then
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
