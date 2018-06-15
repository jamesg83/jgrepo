#SCOM 2012x Health Check HTML Report
#Version 2.3
#18-09-2015
#Modified by Marnix Wolf:
##2.3 modifications:
###Added GW Server Primary & Failover MS server information
###Added warning when no Failover MS server for GW server is detected
###Added Generic Alerts Overview section
###Fixed issue where Unhealthy MMAs table didn't show in report
###Fixed issue where MMAs in MM table didn't show in report
###Fixed issue where MMAs in MM table missed Maintenance Mode Comments
###Added report section with SCOM license information
###Added detection and warning when only one SCOM MS server is present
###Modified report formatting for improved reading experience
##2.1 modifications:
###Added Generic Information section
###Added Top 10 Rule Based Closed Alerts section
###Added Top 10 Monitor Based Closed Alerts section
##2.0 modifications:
###Added Stopwatch for calculating total report creation time, displayed in last section 'Report Creation Time'
###Modified HTML Report file name, containing creation date and name of the MG
##1.9 modifications:
###Modified some report formatting and section titles
###Disabled the section which e-mails the HTML file
#Resources: 
## Jason Rydstrand for 2007 R2: http://blogs.technet.com/b/jasonrydstrand/archive/2013/03/27/daily-scom-health-check-with-powershell.aspx
## Scott Moss, updated to SCOM 2012x: https://gallery.technet.microsoft.com/scriptcenter/SCOM-2012-Daily-Health-5950d801
## Bob Cornelissen: SCOM Connection
## Tao Yang: Sanity Check
## PowerShell.com for Stopwatch class: http://powershell.com/cs/blogs/tips/archive/2014/04/09/logging-script-runtime.aspx
#Additional Remarks:
## Update $UserName and $Password for your email server on Gmail
## Also update '$mailmessage.from' and '$mailmessage.To'. Add with who its coming from and going to

#Define Stopwatch and start it
$StopWatch = [system.diagnostics.stopwatch]::startNew()
$StopWatch.Start()

#Importing the SCOM PowerShell module
Import-module OperationsManager
#Connect to localhost when running on the management server
$connect = New-SCOMManagementGroupConnection -ComputerName localhost

#Or enable the two lines below and of course enter the FQDN of the management server in the first line.
#$MS = "enter.fqdn.name.here"
#$connect = New-SCOMManagementGroupConnection “ComputerName $MS

#Define Report Date, used in filename of HTML Report
$ReportDate = Get-Date -UFormat "%Y%m%d_"
$ReportTime = Get-Date -Format "H-mm-ss"

#Get Date and Time in special formatting for HTML report
$Date = Get-Date -UFormat "%A %d %B %Y"
$Time = Get-Date -UFormat "%R"

# Create header for HTML Report
$Head = "<style>"
$Head +="BODY{background-color:#CCCCCC;font-family:Calibri; font-size: 12pt;}"
$Head +="TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse; width: 100%;}"
$Head +="TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:#0066ff;color:white;padding: 5px; font-weight: bold;text-align:left;}"
$Head +="TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:#F0F0F0; padding: 2px;}"
$Head +="</style>"

#Define the customer name here. Modify the value for parameter $CustomerName
$CompanyName = "ENTER COMPANY NAME HERE"

#Collect Generic information
Write-Host "Collecting Generic Information" -ForegroundColor Gray
##Fill a bunch of variables
$MGName = Get-SCOMManagementGroup | Select-Object -Expand Name
$AmountGWs = (Get-SCOMManagementServer | where {$_.IsGateway -eq $true}).count
$AmountMSs = (Get-SCOMManagementServer | where {$_.IsGateway -eq $false}).count
$AmountMMAs = (Get-SCOMAgent).count
#$SCOMLicenseProduct = Get-SCOMManagementGroup | ForEach-Object {$_.SkuForProduct}
#$SCOMLicenseVersion = Get-SCOMManagementGroup | ForEach-Object {$_.Version}
#$SCOMLicenseSKU = Get-SCOMManagementGroup | ForEach-Object {$_.SkuForLicense}
#$SCOMLicenseExp = Get-SCOMManagementGroup | ForEach-Object {$_.TimeOfExpiration}
$SCOMMG = Get-SCOMManagementGroup
$SCOMLicenseProduct = $SCOMMG.SkuForProduct
$SCOMLicenseVersion = $SCOMMG.Version
$SCOMLicenseSKU = $SCOMMG.SkuForLicense
$SCOMLicenseExp = $SCOMMG.TimeOfExpiration
$SCOMLicensesLogicalCPU = Get-SCOMAccessLicense | measure-object -property LogicalProcessorCount -sum | ForEach-Object {$_.Sum}
$SCOMLicensesPhysicalCPU = Get-SCOMAccessLicense | measure-object -property PhysicalProcessorCount -sum | ForEach-Object {$_.Sum}
##Pipe these variables into the report
#When enabling the e-mail section of this PS script enable the next line (62) and add a '+' to line 63 so it becomes '$ReportOutput =+' instead of '$ReportOutput ='
#$ReportOutput = "To enable HTML view, click on `"This message was converted to plain text.`" and select `"Display as HTML`"" 
$ReportOutput = "<span style='color:#190707'>"
$ReportOutput += "<p><H1>SCOM Health Check Report for Management Group: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$MGName</H1></p>"
$ReportOutput += "<p><H2>Company: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$CompanyName</H2></p>"
$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<p><H2>Generic Information</H2></p>"
$ReportOutput += "<span style='color:#190707'>"
$ReportOutput += "<p><H3><u>Report Information</H3></p></u>"
$ReportOutput += "<ul>"
$ReportOutput += "<li>Report File Name: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$ReportDate$ReportTime _SCOM Health Check Report MG $MGName.html.</li>"
$ReportOutput += "<li>Report File Folder: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "C:\Server Management.</li>"
$ReportOutput += "<li>Creation Date & Time: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$Date, $Time hrs.</li>"
$ReportOutput += "</ul>"
$ReportOutput += "<p><H3><u>SCOM Infrastructure Overview</H3></p></u>"
$ReportOutput += "<ul>"
$ReportOutput += "<li>Management Group Name: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$MGName</li>"
$ReportOutput += "<li>Total amount of SCOM Management Servers: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$AmountMSs</li>"
$ReportOutput += "<li>Total amount of SCOM Gateway Servers: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$AmountGWs</li>"
$ReportOutput += "<li>Total amount of Microsoft Monitoring Agents: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$AmountMMAs</li>"
$ReportOutput += "</ul>"
#Enumerate SCOM Management Servers
$ManagementServers = Get-SCOMManagementServer | where {$_.IsGateway -eq $false} | Sort-Object DisplayName
$ReportOutput += "<H3><u>SCOM Management Servers Overview</H3></u>"
$ReportOutput += "<ol>"
$ReportOutput += "<span style='color:#0000FF'>"
foreach ($ManagementServer in $ManagementServers)
{
$ManagementServerName = $ManagementServer.DisplayName 
$ReportOutput += "<li>$ManagementServerName</li>"
}
$ReportOutput += "</ol>"
#Throw warning when only ONE SCOM MS server is detected
If ($AmountMSs -eq 1)
    {
    $ReportOutput += "<mark><span style='color:#FF0000'>!!!WARNING!!! Only ONE SCOM Management Server detected. This kind of setup is limited to labs ONLY, NOT for POC(s) and OTAP streets where TWO SCOM MS servers are the minimum requirement.</mark>"
    $ReportOutput += "<span style='color:#190707'>"
    }
Else {}
#Enumerate SCOM Gateway Servers
if($AmountGWs -gt 0) { 
 $GatewayServers = Get-SCOMManagementServer | where {$_.IsGateway -eq $true} | Sort-Object DisplayName
 $ReportOutput += "<H3><u>SCOM Gateway Servers Overview</H3></u>"
 $ReportOutput += "<ol>"
 $ReportOutput += "<span style='color:#0000FF'>"
 foreach ($GatewayServer in $GatewayServers)
 {
 $GatewayServerName = $GatewayServer.DisplayName 
 $ReportOutput += "<li>$GatewayServerName"
 $PrimaryMS = ($GatewayServer.GetPrimaryManagementServer()).ComputerName
 $FailoverMS = ($GatewayServer.GetFailoverManagementServers()).ComputerName
 If ($FailoverMS -eq $null) {
 $FailoverMS = "<mark><span style='color:#FF0000'>None present. Please add a Failover MS Server for this Gateway Server!!!<span style='color:#190707'></mark>"
 }
 Else{}
 $ReportOutput +="<span style='color:#190707'>. Primary MS: <span style='color:#0000FF'> $PrimaryMS. <span style='color:#190707'> Failover MS: <span style='color:#0000FF'> $FailoverMS</li>"
 #$ReportOutput += "<span style='color:#190707'>"
 }
 $ReportOutput += "</ol>"
} 
else { 
 $ReportOutput += "<H3><u>Gateway Servers</H3></u>"
 $ReportOutput += "<p>No SCOM Gateway Servers found.</p>"
}
#SCOM License Information
$ReportOutput += "<p><H3><u>System Center Licenses Overview</H3></p></u>"
$ReportOutput += "<ul>"
$ReportOutput += "<li>Licensed product: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$SCOMLicenseProduct</li>"
$ReportOutput += "<li>License SKU: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$SCOMLicenseSKU</li>"
$ReportOutput += "<li>License version: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$SCOMLicenseVersion</li>"
$ReportOutput += "<li>License valid til: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$SCOMLicenseExp.</li>"
$ReportOutput += "<li>SCOM Logical CPU licenses in use (sum): "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$SCOMLicensesLogicalCPU</li>"
$ReportOutput += "<li>SCOM Physical CPU licenses in use (sum): "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$SCOMLicensesPhysicalCPU</li>"
$ReportOutput += "</ul>"

# Generic Alert Information
write-host "Collecting Generic Alert Information" -ForegroundColor Gray
$AlertsUnclosedAll = Get-SCOMAlert -Criteria "ResolutionState <> '255'" | Measure
$AlertsUnclosedMonitorBased = Get-SCOMAlert -Criteria "ResolutionState <> '255' AND IsMonitorAlert = '$true'" | Measure
$AlertsUnclosedRuleBased = Get-SCOMAlert -Criteria "ResolutionState <> '255' AND IsMonitorAlert = '$false'" | Measure
$CountUnclosedAlerts = $AlertsUnclosedAll.count
$CountMonitorBasedUnclosedAlerts = $AlertsUnclosedMonitorBased.count
$CountRuleBasedUnclosedAlerts = $AlertsUnclosedRuleBased.count
$ReportOutput += "<p><H3><u>Alerts Overview</H3></p></u>"
$ReportOutput += "<ul>"
$ReportOutput += "<li>Total # Unclosed Alert(s): "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$CountUnclosedAlerts</li>"
$ReportOutput += "<li>Total # Monitor based unclosed Alert(s): "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$CountMonitorBasedUnclosedAlerts</li>"
$ReportOutput += "<li>Total # Rule  based unclosed Alert(s): "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$CountRuleBasedUnclosedAlerts</li>"
$ReportOutput += "</ul>"

# Get status of Management Server Health and input them into report
write-host "Collecting Management Health Server States" -ForegroundColor Gray 
$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<p><H2>Unhealthy SCOM Management Servers</H2></p>"
$ReportOutput += "<span style='color:#190707'>"

$Count = Get-SCOMManagementServer | where {$_.HealthState -ne "Success"} | Measure-Object
if($Count.Count -gt 0) { 
 $ReportOutput += "<span style='color:#0000FF'>"
 $ReportOutput += Get-SCOMManagementServer | where {$_.HealthState -ne "Success"} | select Name,HealthState,IsRootManagementServer,IsGateway | ConvertTo-HTML -fragment
} else { 
 $ReportOutput += "<span style='color:#0000FF'>"
 $ReportOutput += "<p> All SCOM Management Servers are healthy.</p>"
} 

#Set font color back to black
$ReportOutput += "<span style='color:#190707'>"

#Get MMA Health Status and pipe unhealthy MMAs (or MMAs in MM) into the report
write-host "Collecting MMA in MM" -ForegroundColor Gray
$MG = get-scommanagementgroup 
$criteria = new-object Microsoft.EnterpriseManagement.Monitoring.MonitoringObjectGenericCriteria("InMaintenanceMode=1")
$objectsInMM = $MG.GetPartialMonitoringObjects($criteria.Criteria)
$ObjectsFound = $objectsInMM | select-object DisplayName, @{name="Object Type";expression={foreach-object {$_.GetLeastDerivedNonAbstractMonitoringClass().DisplayName}}},@{name="StartTime";expression={foreach-object {$_.GetMaintenanceWindow().StartTime.ToLocalTime()}}},@{name="EndTime";expression={foreach-object {$_.GetMaintenanceWindow().ScheduledEndTime.ToLocalTime()}}},@{name="Path";expression={foreach-object {$_.Path}}},@{name="User";expression={foreach-object {$_.GetMaintenanceWindow().User}}},@{name="Reason";expression={foreach-object {$_.GetMaintenanceWindow().Reason}}},@{name="Comment";expression={foreach-object {$_.GetMaintenanceWindow().Comments}}}
#$ReportOutput += "<li>$ObjectsFound</li>"

$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<h2>Unhealthy Microsoft Monitoring Agents (MMAs)</h2>"
$ReportOutput += "<span style='color:#190707'>"
$ReportOutput += Get-SCOMAgent | where {$_.HealthState -ne "Success"} | Sort-Object HealthState -descending | select Name,HealthState | ConvertTo-HTML -fragment

$SickMMAs = (Get-SCOMAgent | where {$_.HealthState -ne "Success"}).count
If ($SickMMAs -eq 0) 
    {
    $ReportOutput += "<span style='color:#0000FF'>" 
    $ReportOutput += "All Microsoft Monitoring Agents are healthy."
    }
Else
    { 
    $ReportOutput += "<span style='color:#FF0000'>"
    $ReportOutput += "Please check the table 'Microsoft Monitoring Agents (MMAs) in Maintenance Mode(MM)' since some 'unhealthy' MMAs might be only in MM and as such 'Uninitialized'."
    $ReportOutput += "<span style='color:#190707'>"
    }


$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<h2>Microsoft Monitoring Agents (MMAs) in Maintenance Mode(MM)</h2>"
$ReportOutput += "<span style='color:#190707'>"
$Agents = Get-scomAgent | where {$_.HealthState -ne "Success"} | Sort-Object HealthState -descending | select Name,HealthState

If ($Agents -eq $null)
    {
    $ReportOutput += "<span style='color:#0000FF'>" 
    $ReportOutput += "All Microsoft Monitoring Agents are active."
    $ReportOutput += "<span style='color:#FF0000'>"
    $ReportOutput += "<p>Remark: When an Object is only for some minutes out of MM, it may be listed here with the 'MM' status 'No'. This is due to the design of the SCOM database and can't be altered.</p>"
    $ReportOutput += "<span style='color:#190707'>"
    }
Else
{
$AgentTable = New-Object System.Data.DataTable "$AvailableTable"
$AgentTable.Columns.Add((New-Object System.Data.DataColumn Name,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn HealthState,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MM,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMUser,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMReason,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMComment,([string])))
$AgentTable.Columns.Add((New-Object System.Data.DataColumn MMEndTime,([string])))

foreach ($Agent in $Agents)
    {
        $FoundObject = $null
	$MaintenanceModeUser = $null
	$MaintenanceModeComment = $null
	$MaintenanceModeReason = $null
	$MaintenanceModeEndTime = $null
        $FoundObject = 0
        $FoundObject = $objectsFound | ? {$_.DisplayName -match $Agent.Name -or $_.Path -match $Agent.Name}
        if ($FoundObject -eq $null)
            {
                $MaintenanceMode = "No"
                $MaintenanceObjectCount = 0
            }
        else
            {
                $MaintenanceMode = "Yes"
                $MaintenanceObjectCount = $FoundObject.Count
		$MaintenanceModeUser = (($FoundObject | Select User)[0]).User
		$MaintenanceModeReason = (($FoundObject | Select Reason)[0]).Reason
		$MaintenanceModeComment = (($FoundObject | Select Comment)[0]).Comment
		$MaintenanceModeEndTime = ((($FoundObject | Select EndTime)[0]).EndTime).ToString()
            }
        $NewRow = $AgentTable.NewRow()
        $NewRow.Name = ($Agent.Name).ToString()
        $NewRow.HealthState = ($Agent.HealthState).ToString()
        $NewRow.MM = $MaintenanceMode
	    $NewRow.MMUser = $MaintenanceModeUser
        $NewRow.MMReason = $MaintenanceModeReason
        $NewRow.MMComment = $MaintenanceModeComment
        $NewRow.MMEndTime = $MaintenanceModeEndTime
        $AgentTable.Rows.Add($NewRow)
    }
}
    
#$ReportOutput += $AgentTable | Sort-Object MM | Select Name, HealthState, MM, MMUser, MMReason, MMComment, MMEndTime | ConvertTo-HTML -fragment
$MMAResults = $AgentTable | Sort-Object MM | Select Name, HealthState, MM, MMUser, MMReason, MMComment, MMEndTime | ConvertTo-HTML -fragment
$ReportOutput += $MMAResults


# Get Alerts specific to Management Servers and put them in the report
write-host "Collecting Management Server Alerts" -ForegroundColor Gray 
$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<h2>SCOM Management Server Alerts</h2>"
$ReportOutput += "<span style='color:#190707'>"
$ManagementServers = Get-SCOMManagementServer
$SCOMMSAlerts = get-SCOMalert -Criteria ("NetbiosComputerName = '" + $ManagementServer.ComputerName + "'") | where {$_.ResolutionState -ne '255' -and $_.MonitoringObjectFullName -Match 'Microsoft.SystemCenter'}
If ($SCOMMSAlerts -eq $null)
    {
    $ReportOutput += "<span style='color:#0000FF'>" 
    $ReportOutput += "No SCOM Management Server Alerts found."
    }
Else
{
foreach ($ManagementServer in $ManagementServers){ 
 $ReportOutput += "<h3><u>Alerts on " + $ManagementServer.ComputerName + "</h3></u>"
 $ReportOutput += get-SCOMalert -Criteria ("NetbiosComputerName = '" + $ManagementServer.ComputerName + "'") | where {$_.ResolutionState -ne '255' -and $_.MonitoringObjectFullName -Match 'Microsoft.SystemCenter'} | select TimeRaised,Name,Description,Severity | ConvertTo-HTML -fragment
}
}

# Get all alerts
write-host "Collecting All Alerts" -ForegroundColor Gray
$Alerts = Get-SCOMAlert -Criteria 'ResolutionState < "255"'

# Get alerts for last 24 hours
write-host "Collecting Alerts - Last 24 hrs" -ForegroundColor Gray
$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<h2>Top 10 Alerts With Same Name - 24 hours</h2>"
$ReportOutput += "<span style='color:#190707'>"
$ReportOutput += $Alerts | where {$_.LastModified -le (Get-Date).addhours(-24)} | Group-Object Name | Sort-object Count -desc | select-Object -first 10 Count, Name | ConvertTo-HTML -fragment

$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<h2>Top 10 Repeating Alerts - 24 hours</h2>"
$ReportOutput += "<span style='color:#190707'>"
$ReportOutput += $Alerts | where {$_.LastModified -le (Get-Date).addhours(-24)} | Sort-Object -desc RepeatCount | select-Object -first 10 RepeatCount, Name, MonitoringObjectPath, Description | ConvertTo-HTML -fragment

# Get the Top 10 Unresolved alerts still in console and put them into report
write-host "Collecting Top 10 Unresolved Alerts With Same Name - All Time" -ForegroundColor Gray 
$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<h2>Top 10 Unresolved Alerts</h2>"
$ReportOutput += "<span style='color:#190707'>"
$ReportOutput += $Alerts  | Group-Object Name | Sort-object Count -desc | select-Object -first 10 Count, Name | ConvertTo-HTML -fragment

# Get the Top 10 Repeating Alerts and put them into report
write-host "Collecting Top 10 Repeating Alerts - All Time" -ForegroundColor Gray 
$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<h2>Top 10 Repeating Alerts</h2>"
$ReportOutput += "<span style='color:#190707'>"
$ReportOutput += $Alerts | Sort -desc RepeatCount | select-object -first 10 Name, RepeatCount, MonitoringObjectPath, Description | ConvertTo-HTML -fragment

# Get the Top 10 Rule Based Closed Alerts and put them into report
Write-Host "Collecting Top 10 Rule Based Closed Alerts" -ForegroundColor Gray 
$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<h2>Top 10 Rule Based Closed Alerts</h2>"
$ReportOutput += "<span style='color:#190707'>"
$RuleBasedClosedAlerts = Get-SCOMAlert -Criteria "ResolutionState = '255' AND IsMonitorAlert = 'False'" 
$ReportOutput += $RuleBasedClosedAlerts | Sort Name | Group Name | Sort Count -Descending | Select Count, Name -First 10 | ConvertTo-HTML -fragment

# Get the Top 10 Monitor Based Closed Alerts and put them into report
Write-Host "Collecting Top 10 Monitor Based Closed Alerts" -ForegroundColor Gray 
$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<h2>Top 10 Monitor Based Closed Alerts</h2>"
$ReportOutput += "<span style='color:#190707'>"
$MonitorBasedClosedAlerts = Get-SCOMAlert -Criteria "ResolutionState = '255' AND IsMonitorAlert = 'True'" 
$ReportOutput += $MonitorBasedClosedAlerts | Sort Name | Group Name | Sort Count -Descending | Select Count, Name -First 10 | ConvertTo-HTML -fragment

# Get list of agents still in Pending State and put them into report
write-host "Collecting MMAs In Pending State" -ForegroundColor Gray 
$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<h2>Microsoft Monitoring Agents in Pending State</h2>"
$ReportOutput += "<span style='color:#190707'>"
$ReportOutput += Get-SCOMPendingManagement | sort AgentPendingActionType | select AgentName,ManagementServerName,AgentPendingActionType | ConvertTo-HTML -fragment

# List Management Packs updated in last 24 hours
write-host "Collecting Management Packs - Updated Last 24 hrs" -ForegroundColor Gray
$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<h2>Recently Updated Management Packs</h2>"
$ReportOutput += "<span style='color:#190707'>"
$MPDates = (Get-Date).adddays(-1)
$ReportOutput += Get-SCManagementPack | Where {$_.LastModified -gt $MPDates} | Select-Object DisplayName, LastModified | Sort LastModified | ConvertTo-Html -fragment

#End Stopwatch and display total run time 
$StopWatch.Stop()
$ElapsedMinutes = $StopWatch.Elapsed.Minutes
$ElapsedSeconds = $StopWatch.Elapsed.Seconds
$ReportOutput += "<span style='color:#DF7401'>"
$ReportOutput += "<h2>Report Creation Time: "
$ReportOutput += "<span style='color:#0000FF'>"
$ReportOutput += "$ElapsedMinutes minutes and $ElapsedSeconds seconds</h2>"

#Test for presence of folder 'C:\Server Management'. When not found create it
$Folder = Test-Path -Path "C:\Server Management"
If ($Folder -eq 'True')
##Folder found no further actions required
{}
Else 
##Folder not present. It will be created now
    {
    md "C:\Server Management" | Out-Null
    }

# Take all $ReportOutput and combine it with $Body to create completed HTML output
$Body = ConvertTo-HTML -head $Head -body "$ReportOutput"
$Body | Out-File "C:\Server Management\$ReportDate$ReportTime _SCOM Health Check Report MG $MGName.html"
Invoke-Item "C:\Server Management\$ReportDate$ReportTime _SCOM Health Check Report MG $MGName.html"

#Setup and send output as email message to GMAIL
##REMOVE ALL TAGS IN ORDER TO ENABLE THIS SECTION!!!
##Update $UserName and $Password for your email server on Gmail
##Also update '$mailmessage.from' and '$mailmessage.To'. Add with who its coming from and going to
#$SMTPServer = "smtp.gmail.com"
#$SMTPPort = "587"
#$Username = "GMAIL User Name HERE"
#$Password = "GMAIL Password HERE"
#$Body = ConvertTo-HTML -head $Head -body "$ReportOutput"
#$SmtpClient = New-Object system.net.mail.smtpClient($SMTPServer, $SMTPPort);
#$MailMessage = New-Object system.net.mail.mailmessage
#$mailmessage.from = "sender@company.com"
#$mailmessage.To.add("receiver01@company.com")
#$mailmessage.To.add("receiver02@company.com")
#$mailmessage.Subject = "SCOM Daily Healthcheck Report MG $MGName"
#$MailMessage.IsBodyHtml = $true
#$smtpClient.EnableSSL = $true
#$smtpClient.Credentials = New-Object System.Net.NetworkCredential($Username, $Password);
#$mailmessage.Body = $Body
#$smtpclient.Send($mailmessage)

#End of script