' Check the last modified date/time of the "\\mmhapp02\GPMessageDeliveryService_Prod\Undeliverable" folder.
' If it's over the configured threshold then we haven't received any new files from the PDR web service in that time so send a SCOM alert.
' This script records it's state (GOOD/BAD) and will only alert when the state changes.

' Get computer name
Set WshShell = CreateObject("WScript.Shell")
strComputerName = wshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
Set WshShell = Nothing

'---------------------------------------------------------------------
' Script config
'---------------------------------------------------------------------
SCOMSource = "SCOM_FUNC_RULE1"

dtNow = Now
dtYear = DatePart("yyyy",dtNow)
dtMonth = DatePart("m",dtNow)
dtDay = DatePart("d",dtNow)

If Len(dtMonth) = 1 Then
	dtMonth = "0" & dtMonth
End If
If Len(dtDay) = 1 Then
	dtDay = "0" & dtDay
End If

strFolderPath = "E:\Data\eGate\GPMessageDeliveryService_Prod\Undeliverable\CMDHB_BadgerNet\" & dtYear & "-" & dtMonth & "-" & dtDay & ""
strFolderUNC = "\\mmhapp02\GPMessageDeliveryService_Prod\Undeliverable\CMDHB_BadgerNet\" & dtYear & "-" & dtMonth & "-" & dtDay & ""
SCOMUnHealthyDescription = "WARNING: New undelivered GP messages found in " & strFolderUNC & "."
SCOMUnHealthyID = Right(SCOMSource,1) & "2"
strStateHistoryFile = SCOMSource & "_last_state.txt"

'---------------------------------------------------------------------
' Script guts
'---------------------------------------------------------------------

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check for previous state
If Not objFSO.FileExists(strStateHistoryFile) Then
	Set objFile = objFSO.CreateTextFile(strStateHistoryFile,True)
	dtLastState = "UNKNOWN"
	objFile.Write dtLastState
	objFile.Close
	Set objFile = Nothing
Else
	Set objFile = objFSO.OpenTextFile(strStateHistoryFile, 1)
	dtLastState = Trim(objFile.ReadAll)
	objFile.Close
	Set objFile = Nothing
End If

' Get folder last modified date/time
If objFSO.FolderExists(strFolderPath) Then
	Set objFolder = objFSO.GetFolder(strFolderPath)
	dtLastModified = Trim(objFolder.DateLastModified)
	Set objFolder = Nothing

	' If modified date/time has changed then log scom event
	If dtLastModified <> dtLastState Then
		LogSCOMUnHealthyEvent
		Set objFile = objFSO.CreateTextFile(strStateHistoryFile,True)
		objFile.Write dtLastModified
		objFile.Close
	End If
	Set objFile = Nothing
	Set objFolder = Nothing
End If

Set objFSO = Nothing

'---------------------------------------------------------------------
' SCOM Event creation procedures
Sub LogSCOMUnHealthyEvent
	Set WshShell = CreateObject("WScript.Shell")
	strCommand = "eventcreate /ID " & SCOMUnHealthyID & " /L APPLICATION /T ERROR /SO " & SCOMSource & " /D " & Chr(34) & SCOMUnHealthyDescription & Chr(34)
	WshShell.Run strcommand
	Set WshShell = Nothing
End Sub
