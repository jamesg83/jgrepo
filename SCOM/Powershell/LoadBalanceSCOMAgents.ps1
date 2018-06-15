# author - Ravi Yadav [MVP] - www.SCOMandOtherGeekyStuff.com | Cloud & Datacenter MVP 
# Dec 12, 2015
#--------------------------------------- 
# Path to Operations Manager 2012 R2 Module 
 
$module_path = "C:\Program Files\Microsoft System Center 2012 R2\Operations Manager\Powershell\OperationsManager\OperationsManager.psm1" 
 
# SCOM Server to connect to 
$scom_server = “AKSPMOM01.SCEG.com” 
 
# Name of Resource Pool to retrieve Management Servers from 
$resource_pool_name = “All Management Servers Resource Pool” 
 
# Clear out runtime variables 
$mgt_servers = @() 
 
# Attempt to import SCOM PS Module from $module_path 
Import-Module $module_path -ErrorAction Stop 
 
# Connect to Management Group 
New-SCOMManagementGroupConnection -ComputerName $scom_server 
 
# Get the All Management Servers resource pool 
$resource_pool = Get-SCOMResourcePool -Name $resource_pool_name 
 
# Get management server objects 
foreach ($member in $resource_pool.Members) 
{ 
# Retrieve Management Server Objects in specified resource pool. 
$mgt_servers += Get-SCOMManagementServer -Name $member.DisplayName 
} 
 
# Get all of the SCOM agents 
$agents = Get-SCOMAgent 
 
#Iterate through all agents 
foreach ($agent in $agents) 
{ 
# Split out mgt server array to store first server, and then rest in $other_mgts 
$primary_mgt, $other_mgts = $mgt_servers 
 
# Assign secondary management server to variable 
$secondary_mgt = $mgt_servers[1] 
 
# Clear out mgt_servers 
$mgt_servers = @() 
 
# Clear out failover management server 
Set-SCOMParentManagementServer -Agent $agent -FailoverServer $null 
 
#Set primary management server 
Set-SCOMParentManagementServer -Agent $agent -PrimaryServer $primary_mgt 
 
# Set failover management server 
Set-SCOMParentManagementServer -Agent $agent -FailoverServer $secondary_mgt 
 
# Push primary onto the end of the array for the next loop 
$mgt_servers += $other_mgts 
$mgt_servers += $primary_mgt 
}