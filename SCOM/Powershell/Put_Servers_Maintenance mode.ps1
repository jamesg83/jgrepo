######################################################################## 
#  
# SCOMMaintenanceModeFromFile 
# Autor: Christopher Keyaert 
# Email : Christopher.Keyaert@inovativ.be 
#  
# Date: 21 FEB 2014 
# Version: 1.0 
# 
# PS C:\SCOMMaintenanceModeFromFile> .\SCOMMaintenanceModeFromFile.Ps1 -FileName srvlist.txt -Duration 10 
# 
######################################################################## 
# Functions 
######################################################################## 
##################################### 
# Module 
##################################### 
param( 
  [string]$FileName, 
  [string]$Duration 
) 
 
Import-Module OperationsManager 
new-SCOMManagementGroupConnection -ComputerName Localhost 
 
##################################### 
# Script 
##################################### 
$path = "C:\temp" 
$domain = "healthcare.huarahi.health.govt.nz" 
 
##################################### 
# Params 
##################################### 
 
#Get Server list 
 
$MyFile = Get-content "$path\$Filename" 
$MyFile 
foreach($srv in $MyFile) 
    { 
    Write-host "ServerName : $srv" 
     
    $startTime = [DateTime]::Now 
    $endTime = $startTime.AddMinutes($Duration) 
     
    $srv += ".$domain" 
     
    $Class = get-SCOMclass | where-object {$_.Name -eq "Microsoft.Windows.Computer"}; 
    $Instance = Get-SCOMClassInstance -Class $Class | Where-Object {$_.Displayname -eq "$srv"}; 
    Start-SCOMMaintenanceMode -Instance $Instance -Reason "PlannedOther" -EndTime $endTime -Comment "Scheduled SCOM Maintenance Window" 
     
     
    } 