import-module operationsmanager

Get-SCOMNotificationSubscription | foreach {
 $ns = $_.DisplayName
 $_.ToRecipients | foreach { 
  If ($_.Name -match "NAME") {
   Write-Host $ns
  }
 }
}