######################################################################################
######################################################################################
#Add-SCOMAgentManagementGroup.ps1
#Author - Ross Worth
#http://www.bytepro.co.uk
#
#Uses COM to add a management group to a SCOM Agent.
#Free to redistribute and use.
#Inspired by Bob Cornelissen's VB Script.
######################################################################################
######################################################################################
#Edit Variables
$inputFile = "D:\SCOM\Script\ServerList.txt"
$servers = Get-Content $inputFile
$NewManagementGroupName = "hA_Reg_PROD" #SCOM Management Group Name
$NewManagementServerOrGateway = "vhal1omm001.healthcare.huarahi.health.govt.nz" #SCOM Management server or Gateway FQDN
#CreateNew Object.
$a = New-Object -ComObject AgentConfigManager.MgmtSvcCfg
#Edn of Variables

foreach ($server in $servers)
{

#Add Management Group
$a.AddManagementGroup("$NewManagementGroupName", "$NewManagementServerOrGateway",5723)
#Restart SCOM Agent
Restart-Service HEALTHSERVICE
}