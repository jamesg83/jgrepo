$Time = ((Get-Date).AddMinutes(30))
$comment = "WannaCry patching"
$reason = "Patching"
<# -Reason Valid values are: PlannedOther,
UnplannedOther, PlannedHardwareMaintenance, UnplannedHardwareMaintenance, PlannedHardwareInstallation,
UnplannedHardwareInstallation, PlannedOperatingSystemReconfiguration, UnplannedOperatingSystemReconfiguration,
PlannedApplicationMaintenance, ApplicationInstallation, ApplicationUnresponsive, ApplicationUnstable,
SecurityIssue, LossOfNetworkConnectivity.
#> 

$filepath = "C:\temp\servers.csv"
$servers = Import-Csv $filepath

foreach ($server in $servers) 
{
$Instance = Get-SCOMClassInstance -Name $server.ToUpper()
Write-Host $Instance
Start-SCOMMaintenanceMode -Instance $Instance -EndTime $Time -Reason $reason -Comment $comment 
}