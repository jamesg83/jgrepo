Option Explicit
Dim objMSConfig
Set objMSConfig = CreateObject("AgentConfigManager.MgmtSvcCfg")
Call objMSConfig.RemoveManagementGroup ("ADHB Prod")

