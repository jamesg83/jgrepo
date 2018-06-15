' Check the last modified date/time of a folder
' If it's over the configured threshold, then check the last modified file, if the file name is same as given name, then check the existance of a predefind string, if exist then generate a SCOM alert.
' This script records it's state (GOOD/BAD) and will only alert when the state changes.

' Get computer name
Set WshShell = CreateObject("WScript.Shell")
strComputerName = wshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
Set WshShell = Nothing

'---------------------------------------------------------------------
' Script config
'---------------------------------------------------------------------
strFolderPath = "D:\TCare_Production\Audits"
intAgeThreshold = 15 'minutes
strErrorString = "Blocking function in progress"
strCurrentState = "unknown"

SCOMSource = "SCOM_TECH_MONITOR1"
SCOMHealthyDescription = "RESOLVED: The folder " & strFolderPath & " on server " & strComputerName & " has been modified."
SCOMUnHealthyDescription = "WARNING: The folder " & strFolderPath & " on server " & strComputerName & " hasn't been modified in the last " & intAgeThreshold & " minutes."

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

' See how long ago folder was modified, h = hour, s = second, n = minute, m = month, d = day, w = week
dtFolderAge = DateDiff("n",dtLastModified,Now)

' If newer then looking for the lastest file, the set threshold set the current state to bad otherwise good
If dtFolderAge < intAgeThreshold Then
	Set fileLastModified = GetLastModifiedFile(strFolderPath)
	if UCase(Left(fileLastModified.name, 13)) = "HL7 PROCESSOR" then
		If UNCFileContent(fileLastModified,strErrorString) = "True" then
			strCurrentState = "BAD"
		Else
			strCurrentState = "GOOD"
		End If
	End If
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


'-----------------------------------------------------------------------
'Get last modified file
Function GetLastModifiedFile(ByVal sFolderPath)
  Dim FSO, objFolder, objFile
  Dim objFileResult, longDateTime
  Dim boolRC
 
  Set FSO = CreateObject("Scripting.FileSystemObject")
  boolRC = FSO.FolderExists(sFolderPath)
  If Not boolRC Then
    Set FSO = Nothing
    Set GetLastModifiedFile = Nothing
    Exit Function
  End If

  Set objFolder = FSO.GetFolder(sFolderPath)
  If objFolder.Files.Count = 0 Then
    Set FSO = Nothing
    Set objFolder = Nothing
    Set GetLastModifiedFile = Nothing
    Exit Function
  End If
 
  Set objFileResult = Nothing
  longDateTime = CDate(0)
 
  For Each objFile in objFolder.Files
 
    If objFile.DateLastModified > longDateTime Then
      Set objFileResult = objFile
      longDateTime = objFile.DateLastModified
    End If
   
  Next
 
  Set FSO = Nothing
  Set objFolder = Nothing
  Set GetLastModifiedFile = objFileResult

End Function


Function UNCFileContent(file,content)
	UNCFileContent = "False"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(file, 1, False) 'ForReading = 1, ForWriting = 2, ForAppending = 8
	Do until objFile.AtEndOfStream
		strLine = objFile.ReadLine
		If InStr(strLine,content) Then
			UNCFileContent = "True"
		End If
	Loop
	Set objFile = Nothing
	Set objFSO = Nothing

End Function