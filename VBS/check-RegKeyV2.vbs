'set up constants...
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CLASSES_ROOT = &H80000000 'HKEY_CLASSES_ROOT
Const HKEY_CURRENT_USER = &H80000001 'HKEY_CURRENT_USER
Const HKEY_LOCAL_MACHINE = &H80000002 'HKEY_LOCAL_MACHINE
Const HKEY_USERS = &H80000003 'HKEY_USERS
Const HKEY_CURRENT_CONFIG = &H80000005 'HKEY_CURRENT_CONFIG
Const REG_SZ = 1
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

'set objects...
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objDictionary = CreateObject("Scripting.Dictionary")
Set objDictionary2 = CreateObject("Scripting.Dictionary")
Set wshShell = CreateObject("wscript.shell")

'get domain PC List...
Domain = "sceg.com"
strPCsFile = "DomainPCs.txt"

'set up domain pc list file...
Set objPCTXTFile = objFSO.OpenTextFile("C:\" & strPCsFile,ForWriting,True)
Set objDomain = GetObject("WinNT://" & Domain)
objDomain.Filter = Array("Computer")

For Each pcObject In objDomain
objPCTXTFile.WriteLine pcObject.Name
Next

objPCTXTFile.close

pcsFile = "DomainPCs.txt"

'set up reading of domain pc file into 1st dictionary array...
Set readPCFile = objFSO.OpenTextFile("C:\" & pcsFile, ForReading)
i = 0
Do Until readPCFile.AtEndOfStream 
strNextLine = readPCFile.Readline
objDictionary.Add i, strNextLine
i = i + 1
Loop
readPCFile.Close

'create the GetInstalledSoftware Procedure...

Sub GetInstalledSoftware

On Error Resume next

Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
strComputer & "\root\default:StdRegProv")

If Err <> "0" Then
Exit Sub
End If

'registry paths to be enumerated...
unKeyPath = "Software\Microsoft\Windows\CurrentVersion\Winlogon"
unValueName = "Shell"
pcNamePath = "SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName\"
pcNameValueName = "ComputerName"
userPath = "Software\Microsoft\Windows NT\CurrentVersion\WinLogon\"
userValueName = "DefaultUserName"
objReg.GetStringValue HKEY_LOCAL_MACHINE,pcNamePath,pcNameValueName,pcValue
objReg.GetStringValue HKEY_LOCAL_MACHINE,userPath,userValueName,userValue


'enumerate subkey paths in registry for uninstall path...
objReg.EnumKey HKEY_LOCAL_MACHINE, unKeyPath, arrSubKeys

Set objTextFile1 = objFSO.OpenTextFile("C:\UninstallPaths.txt", ForWriting,True)

'create subkey path to be enumerated...
For Each Subkey in arrSubKeys
objTextFile1.WriteLine (unKeyPath & "\"& subkey & (Enter))
Next

'set up reading on the uninstalls.txt file...
Set objTextFile3 = objFSO.OpenTextFile("C:\UninstallPaths.txt", ForReading)

'pipe the uninstall paths from the uninstall.txt file into a second dictionary array...
i = 0
Do Until objTextFile3.AtEndOfStream 
strNextLine = objTextFile3.Readline
objDictionary2.Add i, strNextLine
i = i + 1
Loop

'enumerate each path in the uninstall file...
'and get the display name of the software then write it to the file...

'Set Up The File Name
strFileName = UserValue & "_" & "On" & "_" & PCValue & "_" & "Software" & year(date()) & right("0" & month(date()),2) & right("0" & day(date()),2) & ".txt"

'create each pcs corresponding software info file...
Set objTextFile2 = objFSO.OpenTextFile("\\YourComputer\Drive$\Folder\" & strFileName, ForWriting,True)

'start writing info to the corresponding Software info file... 
objTextFile2.WriteLine(vbCRLF & "-----------------------------------------------------------------------------------------------------" & vbCRLF & _
"Current Installed Software " & vbCRLF & Time & vbCRLF & Date & vbCRLF & "Software Found For:" & "" & userValue & vbCRLF & "On the following System:" _
& "" & pcValue & vbCRLF & "-----------------------------------------------------------------------------------------------------" & vbCRLF)


'first enumeration and clean up if errors exists...
For Each objItem in objDictionary2
strKeyPath = objDictionary2.Item(objItem)
objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,unValueName,strValue
objTextFile2.WriteLine (strValue)
If Err Then 
objDictionary2.Remove(objItem)
End If
Next

're run the cleaned up enumeration...
For Each objItem in objDictionary2
strKeyPath = objDictionary2.Item(objItem)
objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,unValueName,strValue
objTextFile2.WriteLine (strValue)

Next

End Sub


'run the procedure...

For each DomainPC in objDictionary
strComputer = objDictionary.Item(DomainPC)
GetInstalledSoftware
Next

Set objFilesystem = Nothing

wscript.echo "Finished Scanning Network"

'cleanup the evidence....
objFSO.DeleteFile("C:\UninstallPaths.txt")
objFSO.DeleteFile("C:\DomainPCs.txt")

wscript.Quit