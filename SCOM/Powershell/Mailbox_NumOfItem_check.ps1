Set-executionpolicy unrestricted -force
$exchangeServer = "vhal1exg002","vhal1exg003" | Where {
    Test-Connection -ComputerName $_ -Count 1 -Quiet
} | Get-Random
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://{0}.healthcare.huarahi.health.govt.nz/PowerShell" -f $exchangeServer)

Import-PSSession $ExchangeSession -AllowClobber

Get-MailboxFolderStatistics healthcare\EBS_Prod |
where {$_.name -match "Inbox"} |
Select @{Name="User";Expression={
(Split-Path $_.Identity -Parent)}},
@{Name="Folder";Expression={Split-Path $_.Identity -Leaf}},
@{Name="Items";Expression={$_.ItemsinFolderandSubFolders}}

$mailboxStats = Get-MailboxFolderStatistics healthcare\EBS_Prod -FolderScope Inbox

if ($mailboxStats.ItemsInFolderAndSubFolders -ge 25) {
    #write-Output "BAD"
    Write-EventLog -LogName Application -Source "SCOM_FUNC_MONITOR1" -EntryType Error -EventId 12 -Message "There are $($mailboxStats.ItemsInFolderAndSubFolders) messages in the ebs_prob mailbox, please restart imap service on exchange1"
}

Remove-PSSession $ExchangeSession