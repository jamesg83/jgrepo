Option Explicit
Dim objMSConfig
Set objMSConfig = CreateObject("AgentConfigManager.MgmtSvcCfg")
Call objMSConfig.AddManagementGroup ("ManagementGroupNameToAdd", "server.domain.com",5723)

