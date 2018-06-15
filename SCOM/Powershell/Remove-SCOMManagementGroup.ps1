param(
		[parameter(mandatory=$true)]
		$ManagementGroup, 
		$ComputerName = "Localhost")

Function Remove-SCOMManagementGroup {
	param(
		[parameter(mandatory=$true)]
		$ManagementGroup,
		$ComputerName = "Localhost"
	)
	$sb = {
		param($ManagementGroup, 
		  $ComputerName = "Localhost")
		Try {
			$OMCfg = New-Object -ComObject AgentConfigManager.MgmtSvcCfg
		} catch {
			throw "$ComputerName doesn't have the SCOM 2012 agent installed"
		}
		$mgs = $OMCfg.GetManagementGroups() | %{$_.managementGroupName}
		if ($mgs -contains $ManagementGroup) {
			$OMCfg.RemoveManagementGroup($ManagementGroup)
			return "$ManagementGroup removed from $ComputerName"
		} else {
			return "$ComputerName does not report to $ManagementGroup"
		}
	}
	Invoke-Command -ScriptBlock $sb -ComputerName $ComputerName -ArgumentList @($ManagementGroup,$ComputerName)
}
Remove-SCOMManagementGroup -ManagementGroup $ManagementGroup -ComputerName $ComputerName