'function checkRegKeyValue (stringPath, stringRegkey, stringValue)
    'Const HKEY_CLASSES_ROOT = &H80000000 'HKEY_CLASSES_ROOT
    'Const HKEY_CURRENT_USER = &H80000001 'HKEY_CURRENT_USER
Const HKEY_LOCAL_MACHINE=&H80000002
Const ForReading = 1
Const ForAppending = 8

on error resume next
    'Const HKEY_USERS = &H80000003 'HKEY_USERS
    'Const HKEY_CURRENT_CONFIG = &H80000005 'HKEY_CURRENT_CONFIG

    strInputFile = WScript.Arguments.Item(0)
    strOutputFile = WScript.Arguments.Item(1)

    Dim strRegkey, strComputer, strKeyPath, strKeyValue

    Set objFSO = CreateObject("Scripting.FileSystemObject") 

    Set objTextFileIn = objFSO.OpenTextFile(strInputFile, ForReading)
    Set objTextFileOut = objFSO.OpenTextFile(strOutputFile, ForAppending, True)
'
    Do Until objTextFileIn.AtEndOfStream
        strComputer = Trim(objTextFileIn.Readline)
        Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
'
        strRegKey= "Shell"
        strKeyValue = "<Win32/Nuqel worm copy>"
        strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Winlogon"

        oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strRegKey,strKeyValue
'
        'oReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, strRegkey, strKeyValue, arrSubKeys
'
        'For Each subkey In arrSubKeys
            'bFound = (lcase(subkey) = lcase(regkey))
            'if bFound then exit for
        'Next
        if IsNull(strKeyValue) then
            Wscript.Echo "The registry key does not exist on server:" & strComputer
        Else
            wsh.echo "Found:"
            objTextFileOut.WriteLine(strComputer)
        end if
    LOOP
'end function
