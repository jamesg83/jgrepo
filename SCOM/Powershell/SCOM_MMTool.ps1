<#

#########================================================================================================#########

#

#

# ** Name: SCOMGroupMaintMode.ps1

# ** Author: System Center MVP – Steve Buchanan

# ** Date: 7/30/2015

# ** Version: 1.4

# ** Website: www.buchatech.com

#

# ** Description:

# ** This script can be used to a group of objects within SCOM into Maintenance Mode.

# ** This script needs to be run in an administrative PowerShell console.

# ** This script needs to be run with some mandatory parameters. Parameters shown in examples.

#

############################################

Example when running from PowerShell window:

############################################

-ExecutionPolicy Bypass .\SCOMGroupMaintMode.ps1 -OMServer SCOMSERVERNAME.DOMAINNAME.COM -GroupName ‘NAMEOFYOURGROUP’ -Duration ENTERMINUTESFORDURATION -Reason PlannedOther -Comment ‘ADDCOMMENTIFYOUWANT’

#################################################################################

Example of syntax when placing in Windows Task Scheduler (Actions>>Ad arguments):

#################################################################################

-ExecutionPolicy Bypass C:\LOCATIONOFSCRIPT\SCOMGroupMaintMode.ps1 -OMServer SCOMSERVERNAME.DOMAINNAME.COM -GroupName ‘All Windows Computers’ -Duration 10 -Reason PlannedOther -Comment ‘ADDCOMMENTIFYOUWANT’

###############

Parameters are:

###############

-OMServer

Mandatory parameter containing mgmt server name (Be sure to use FQDN).

-GroupName

Mandatory parameter containing display name of the target group (Be sure to enclose with ”).

-Duration

Mandatory parameter containing integer of desired duration in minutes.

-Reason

Mandatory parameter containing reason. The acceptable values for this parameter are:

— PlannedOther

— UnplannedOther

— PlannedHardwareMaintenance

— UnplannedHardwareMaintenance

— PlannedHardwareInstallation

— UnplannedHardwareInstallation

— PlannedOperatingSystemReconfiguration

— UnplannedOperatingSystemReconfiguration

— PlannedApplicationMaintenance

— ApplicationInstallation

— ApplicationUnresponsive

— ApplicationUnstable

— SecurityIssue

— LossOfNetworkConnectivity

-Comment

Optional parameter description of maintenance action (Free text. Be sure to enclose with ”).

#########================================================================================================#########

#>

###########################

## General Script Settings

###########################

# Error Handling

$ErrorActionPreference = “SilentlyContinue”

# Set Parameters

Param (

[Parameter(Mandatory=$true)][string]$OMServer,

[Parameter(Mandatory=$true)][string]$GroupName,

[Parameter(Mandatory=$true)][Int32]$Duration,

[Parameter(Mandatory=$true)][string]$Reason,

[Parameter(Mandatory=$false)][string]$Comment

)

####################################

# Load needed PowerShell modules

####################################

# Load Operations Manager module

Import-Module -Name OperationsManager

##################

## Start Script

##################

# Make Connection to SCOM Server

New-SCOMManagementGroupConnection -ComputerName $OMServer

$RetriveGroup = (Get-SCOMGroup | where {$_.DisplayName -like “*$GroupName*”})

# Use this for minutes. If you want hours instead comment this line out and use line below.

$Time = ((Get-Date).AddMinutes($Duration))

# Use this for hours. If you want minutes instead comment this line out and use line above.

#$Time = ((Get-Date).AddHours($Duration))

Start-SCOMMaintenanceMode -Instance $RetriveGroup -EndTime $Time -Comment “$Comment” -Reason “$Reason”

# Variable cleanup

Remove-Variable -Name *

##################

## End Script

##################