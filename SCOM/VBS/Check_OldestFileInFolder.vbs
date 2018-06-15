' Find the age of the oldest file in a folder.
' If it's over the configues age then send a SCOM alert.
' This script records it's state (GOOD/BAD) and will only alert when the state changes.

' Written by James Geng

OPTION EXPLICIT
call Main

sub Main()


'---------------------------------------------------------------------
' Script config
'---------------------------------------------------------------------
Dim fso, folder, AgeThreshold, strAgeInterval, intAgeThreshold, strOldestFile, dtmOldestFileDate, objFolder, colFiles, objFile, dtFileAge
Dim oArgs, oAPI, oBag

Set oArgs = Wscript.Arguments
' For strAgeInterval select one of the following: yyyy (Year), q (Quarter), m (Month), y (Day of year), d (Day), w (Weekday), ww (Week of year), h (Hour), n (Minute), s (Second)
' Retrieve parameters
folder = CStr(oArgs.Item(0))
intAgeThreshold = CInt(oArgs.Item(1))
strAgeInterval = CStr(oArgs.Item(2))
WScript.Echo folder 


'---------------------------------------------------------------------
' Script guts
'---------------------------------------------------------------------

Set fso = CreateObject("Scripting.FileSystemObject")

' Instantiate MOM API
Set oAPI = CreateObject("MOM.ScriptAPI")
Set oBag = oAPI.CreatePropertyBag()

' Iterate through any files in the folder and find the one with oldest modified date/time
strOldestFile = "NotFound"
dtmOldestFileDate = Now

Set objFolder = fso.GetFolder(folder)
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
	Exit Sub
End If

' See how long ago the oldest file found was modified
dtFileAge = DateDiff(strAgeInterval,dtmOldestFileDate,Now)

' If older then the set threshold set the current state to bad otherwise good
If dtFileAge >= intAgeThreshold Then
	'strCurrentState = "BAD"
	Call oBag.AddValue("File age exceed threshold","Yes")
	oAPI.AddItem(oBag)
    Call oAPI.ReturnItems
Else
	Call oBag.AddValue("File age exceed threshold","No")
	oAPI.AddItem(oBag)
    Call oAPI.ReturnItems
  Exit Sub
End If

' If the current state is different to the previous state

' Clean up
Set objFile = Nothing
Set fso = Nothing

End Sub