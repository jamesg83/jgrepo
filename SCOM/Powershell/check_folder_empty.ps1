$directory = "D:\temp\test"

$directoryInfo = Get-ChildItem $directory | Measure-Object
$directoryInfo.count


if($directoryInfo.Count -eq 0) 
{ 

    Write-EventLog -LogName Application -Source "SCOM_TECH_MONITOR2" -EntryType Information -EventId 21 -Message "Error file does not exist"

} 

Else
{ 
    Write-EventLog -LogName Application -Source "SCOM_TECH_MONITOR2" -EntryType ERROR -EventId 22 -Message "Error file does exist"
} 