 param(
                  [string] $sourceId,
                  [string] $managedEntityId,
                  [string] $filePath  )

                  #Initialize SCOM API

                  $api = new-object -comObject 'MOM.ScriptAPI'
                  $discoveryData = $api.CreateDiscoveryData(0, $SourceId, $ManagedEntityId)
                  write-eventlog -logname "Operations Manager" -Source "Health Service Script" -EventID 999 -Message "TCP Port Monitoring: looking for CSV file" -EntryType Information
                  # $filePath variable contains UNC path of CSV Config file
                  if (test-path $filePath) {
                  write-eventlog -logname "Operations Manager" -Source "Health Service Script" -EventID 999 -Message "TCP Port Monitoring: Accessing CSV file" -EntryType Information
                  $contents = Import-Csv $filePath
                  try{
                  $Path = (Get-ItemProperty "HKLM:SOFTWARE\Microsoft\System Center Operations Manager\12\Setup\Powershell\V2").InstallDirectory
                  $Path1 = $Path + "OperationsManager\OperationsManager.psm1"
                  if (Test-Path $Path1)
                  {
                  Import-Module $Path1
                  }
                  else
                  {
                  Import-Module OperationsManager
                  }
                  New-SCOMManagementGroupConnection
                  #Retrieve all windows computers which can be used as watcher nodes
                  $allServers = Get-SCClass | where { $_.Name -eq ("Microsoft.Windows.Computer")} | get-scommonitoringobject
                  }
                  catch{
                  write-eventlog -logname "Operations Manager" -Source "Health Service Script" -EventID 999 -Message "TCP Port Monitoring: $_" -EntryType Information
                  }
                  #Read line by line from configuration file and create instance of TCP Port Monitoring Class
                  $contents | ForEach-Object{
                  $ServerName = $_.ServerName
                  $PortNumber = $_.PortNumber
                  $WatcherNode = $_.WatcherNode
                  $NoOfRetries = $_.NoOfRetries
                  $TimeWindowInSeconds = $_.TimeWindowInSeconds
                  $Config = "$ServerName"+":"+"$PortNumber" # Will be used as display name
                  write-eventlog -logname "Operations Manager" -Source "Health Service Script" -EventID 555 -Message "Checking servers" -EntryType Information
                  $allServers | ForEach-Object{
                  #Create instance only if the watcher node is managed by SCOM as the instance will hosted by the watcher node.
                  #The hosting object is windows computer whose display name is equal to watcher node value from CSV
                  #If there is no matching windows computer managed by SCOM, then the instance cannot be hosted. Hence the instance is not discovered.
                  if((($_.DisplayName).toLower()).contains($WatcherNode.toLower())){
                  write-eventlog -logname "Operations Manager" -Source "Health Service Script" -EventID 555 -Message "Creating Instance for $Config" -EntryType Information
                  $instance = $discoveryData.CreateClassInstance("$MPElement[Name='RealDolmen.TCPPortMonitoring.ManagementPack.Class']$")
                  $instance.AddProperty("$MPElement[Name='RealDolmen.TCPPortMonitoring.ManagementPack.Class']/ServerName$", $ServerName)
                  $instance.AddProperty("$MPElement[Name='RealDolmen.TCPPortMonitoring.ManagementPack.Class']/Port$", $PortNumber)
                  $instance.AddProperty("$MPElement[Name='RealDolmen.TCPPortMonitoring.ManagementPack.Class']/NoOfRetries$", $NoOfRetries)
                  $instance.AddProperty("$MPElement[Name='RealDolmen.TCPPortMonitoring.ManagementPack.Class']/TimeWindowInSeconds$", $TimeWindowInSeconds)
                  #The hosting object is windows computer whose display name is equal to watcher node value from CSV
                  $instance.AddProperty("$MPElement[Name='Windows!Microsoft.Windows.Computer']/PrincipalName$", $_.DisplayName)
                  $instance.AddProperty("$MPElement[Name='System!System.Entity']/DisplayName$", $Config)
                  $discoveryData.AddInstance($instance)
                  return
                  }
                  }
                  }
                  }
                  $discoveryData
                  Remove-variable api
                  Remove-variable discoveryData