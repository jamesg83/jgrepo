'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2009
'
' NAME: DoesFileExist
'
' AUTHOR: Pete Zerger, MVP (Cloud and Datacenter Admin)
' DATE  : 3/12/2012
'
'  COMMENT: Verifies a target file (including path) exists. 
'           Intended for use with OpsMgr two state script monitor.
'
'==========================================================================

OPTION EXPLICIT
Call Main
Sub Main()

'Declare Variables 
'File-related variables 
Dim fso, folder, file, FilePath
'OpsMgr related variables 
Dim oArgs, oAPI, oBag

Set oArgs = Wscript.Arguments

' Retrieve parameters
folder = CStr(oArgs.Item(0))
file = CStr(oArgs.Item(1))
FilePath = folder & "\" & file
WScript.Echo folder 
WScript.echo file
WScript.Echo FilePath 

' Instantiate File System Object
Set fso = CreateObject("Scripting.FileSystemObject")

' Instantiate MOM API
Set oAPI = CreateObject("MOM.ScriptAPI")
Set oBag = oAPI.CreatePropertyBag()

' Verify the path to the file exists xists
If (fso.FolderExists(folder)) Then

  'Folder exists, submit property bag and continue
  Call oBag.AddValue("FolderExists","Yes")
  WScript.Echo "Folder exists"

 Else

  'Folder does not exist, submit property bag and exit
  Call oBag.AddValue("FolderExists","No")
  Call oBag.AddValue("FileExists","No")
  WScript.Echo "Folder doesn't exist"
  oAPI.AddItem(oBag)
  Call oAPI.ReturnItems
  Exit Sub

End If

' Verify the file exists 
If (fso.FileExists(FilePath)) Then
  'File exists, submit property bag and exit
  Call oBag.AddValue("FileExists","Yes")
  oAPI.AddItem(oBag)
  Call oAPI.ReturnItems

Else

  'File does not exist, submit property bag and exit
  Call oBag.AddValue("FileExists","No")
  WScript.Echo "File doesn't exist"
  oAPI.AddItem(oBag)
  Call oAPI.ReturnItems
  Exit Sub

End If

End Sub