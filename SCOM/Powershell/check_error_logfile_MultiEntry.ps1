[CmdletBinding()]
Param (
    [Parameter(Mandatory=$false, Position=0)]
    [string] $LogFilePath = "C:\Program Files\PlatoAtlas\Bin\errorlog.txt"
)

# Version 1.06
# Last Changed: 2017-06-27 3:07pm
# Author: Alex Hague - 021 515 755
# 1.06 - Added ability to run under windows server 2003 - James Geng
# 1.05 - Added ability to revert to healthy after x minutes
#      - Added ability to control the event log type logged (error / information / etc)
#
# 1.04 - Improved event log message.
#      - Fixed bug where unhealthy text was not being found.
#
# 1.03 - Added Volume Shadow Copy usage
#      - Added support for arrays of healthy / unhealthy strings

# General Config
$StateStoreFileName = "SCOM-PlatoAltasErrorData.xml"
$StateStorePath = [System.IO.Path]::Combine($env:TEMP, $StateStoreFileName)
#$LogEntryDateTimeFormat = "\d\d\d\d-\d\d-\d\d \d\d:\d\d:\d\d"                         # This is a regex format whose matches must be processable by DateTime.Parse, this may need to be modified with different time format in your log file
$LogEntryDateTimeFormat = ".*\d\D\d\d\D\d\d\d\d *\d:\d\d:\d\d [ap].m."
# Event Log Config
$EventLogName = "Application"
$EventLogSourceName = "SCOM_TECH_MONITOR1"

# Event IDs
$EventId_NoStateChange = 10        # The script ran but no new evidence of a change of state has been found
$EventId_UnhealthyState = 11       # Positive evidence of being unhealthy has been found more recently than the last healthy entry
$EventId_HealthyState = 12         # Positive evidence of being healthy has been found more recently than the last unhealthy entry
$EventId_ErrorRunningScript = 13   # An error occurred while running this script

$EventEntryType_NoStateChange = "Information"
$EventEntryType_HealthyState = "Information"
$EventEntryType_UnhealthState = "Error"
$EventEntryType_ErrorRunningScript = "Error"

# Healthy / Unhealthy Indicators
$HealthyText = $null                               # Set to $null if no healthy text is known, otherwise an array of text e.g. @("A") or @("A", "B", "C")
$UnhealthyText = @("SCRIPT:WorksheetDocProps METHOD:autoinc LINE:14 ERROR NO:1104 MESSAGE:Error reading file")   # Set to $null if no unhealthy text is known, otherwise an array of text e.g. @("A") or @("A", "B", "C")

$RevertToHealthyEventIdAfterMinutes = 14           # After 14 minutes of no unhealthy states, then revert the $EventId to $EventId_HealthyState

function Get-LastDateTimeOfStringOccurrenceInLog
{
    [OutputType([System.DateTime])]
    Param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $LogContents,

        [Parameter(Mandatory=$true, Position=1)]
        [string[]] $TextToMatch,

        [Parameter(Mandatory=$true, Position=2)]
        [string] $LogEntryDateTimeFormat
    )

    $MostRecentLastDateTime = [System.DateTime]::MinValue

    $TextToMatch | % {
        $RegexForMatching = New-Object System.Text.RegularExpressions.Regex ("^($LogEntryDateTimeFormat)(?:.*)ERROR NO:1104 .*$", ([System.Text.RegularExpressions.RegexOptions]::Multiline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase))
        $Matches = $RegexForMatching.Matches($LogContents)
        $LastDateTime = [System.DateTime]::MinValue

        # Check: did we find a match?
        if ($Matches.Count -gt 0) {

            # Yes - we found a match
            #
            # Convert it to an actual date time object
            $LastDateTime = [System.DateTime]::Parse($Matches[$Matches.Count - 1].Groups[1].Value)

            # Check: is this match more recent than the currently most recent match?
            if ($MostRecentLastDateTime -lt $LastDateTime) {
                
                # Yes - this is a more recent match
                #
                # Update the most recent match with this one
                $MostRecentLastDateTime = $LastDateTime
            }
        }
    }

    return $MostRecentLastDateTime
}

function Read-FileContentViaVolumeShadowCopy {
    Param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $FilePath
    )

    # Extract the drive letter
    $DrivePath = [IO.Path]::GetPathRoot($FilePath)

    $Content = $null

    Write-Host "Creating Volume Shadow Copy of $DrivePath ... " -NoNewline

    # Create a volume shadow copy
    $VolumeShadowCopyResult = ([WMICLASS]"root\cimv2:Win32_ShadowCopy").Create($DrivePath, "ClientAccessible")
    if ($?) {
        Write-Host "Completed" -ForegroundColor Green

        # Access the object that represents the volume shadow copy we created
        $VolumeShadowCopy = Get-WmiObject Win32_ShadowCopy | ? { $_.ID -eq $VolumeShadowCopyResult.ShadowID }

        # Construct a temp folder that we can use as a junction point to access the shadow copy
        $TempFolder = $DrivePath + "VSC-TMP-" + ([IO.Path]::GetRandomFileName().Remove(8, 4)) + "\"
        
        # Create a junction point that we can use to access the shadow copy
        Write-Host "Creating NTFS Junction Point ($TempFolder <==> $($VolumeShadowCopy.DeviceObject))... " -NoNewline
        #Invoke-Command { cmd.exe /c mklink /d "$TempFolder" "$($VolumeShadowCopy.DeviceObject)\" } | Out-Null #Uncommon this line for windows server 2008 and above, JG
        Invoke-Command { cmd.exe /c fsutil hardlink create "$TempFolder" "$($VolumeShadowCopy.DeviceObject)\" } | Out-Null #This is for windows server 2003. JG
       
        Write-Host "Completed" -ForegroundColor Green
    
        # Construct the path to the file in the shadow copy
        $FileShadowCopyPath = [IO.Path]::Combine($TempFolder, $FilePath.Remove(0, 3))

        # Read the content of the file at the file path
        $Content = [IO.File]::ReadAllText($FileShadowCopyPath) #Powershell V4
        #$Content = Get-Content $FileShadowCopyPath | Out-String #Powwershell V2

        # Remove the shadow copy
        Write-Host "Removing Volume Shadow Copy... " -NoNewline
        $VolumeShadowCopy.Delete()
        Write-Host "Completed" -ForegroundColor Green

        # Remove the junction point for accessing the shadow copy
        Write-Host "Removing NTFS Junction Point... " -NoNewline
        Invoke-Command { cmd.exe /c rmdir "$TempFolder" }
        Write-Host "Completed" -ForegroundColor Green
    } else {
        Write-Host "Failed" -ForegroundColor Red
    }

    return $Content
}

function Read-FileContent
{
    Param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $FilePath
    )
    $file = [System.io.File]::Open($FilePath, 'Open', 'Read', 'ReadWrite')
    $reader = New-Object System.IO.StreamReader($file)
    $Content = $reader.ReadToEnd()
    $reader.Close()
    $file.Close()
    
    Return $Content

}

function Read-PersistedLastStateEntries
{
    [OutputType([System.Management.Automation.PSCustomObject])]
    Param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $DataPath
    )

    $Data = $null

    # Check: Does the data file need to be created?
    if (-not (Test-Path $DataPath)) {
        
        # Yes - the file does not exist, it needs to be created
        #
        # Create a default blank file 
        $Data = Write-EmptyPersistedLastStateEntries $DataPath

    } else {

        # No - the file exists, we can read from it
        #
        # Load the data from the file
        $Data = Import-Clixml -Path $DataPath

    }

    return $Data
}

function Test-EventLogSource
{
    [OutputType([System.Boolean])]
    Param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $LogSourceName
    )

    $EventLogSourceFound = $false

    try {
        $EventLogSourceFound = [System.Diagnostics.EventLog]::SourceExists($LogSourceName)
    } catch {
        # Nothing to do here, just ignore the error generated by inaccessible logs
    }

    return $EventLogSourceFound
}

function Write-EmptyPersistedLastStateEntries
{
    [OutputType([System.Management.Automation.PSCustomObject])]
    Param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $DataPath
    )

    $Data = New-Object PSObject -Property @{
        LastHealthyState = [System.DateTime]::MinValue
        LastUnhealthyState = [System.DateTime]::MinValue
    }

    $Data | Export-Clixml -Path $DataPath

    return $Data
}

function Write-PersistedLastStateEntries
{
    Param (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $DataPath,

        [Parameter(Mandatory=$true, Position=1)]
        [PSObject] $Data
    )

    $Data | Export-Clixml -Path $DataPath
}

####### Start of Script #######

# Check: is this script registered as an event log source?
if (-not (Test-EventLogSource $EventLogSourceName)) {

    # No - we need to register as an event log source
    #
    # Register as an event log source 
    New-EventLog -LogName $EventLogName -Source $EventLogSourceName

}

# Check: does the log file path actually exist?
if (Test-Path $LogFilePath) {

    # Yes - the log file path exists
    #
    # Use volume shadow copy to access the log file
    
    #$LogContents = Read-FileContentViaVolumeShadowCopy $LogFilePath #For server 2008 and above
    
    $LogContents = Read-FileContent $LogFilePath #For server 2003

    # Initialize variables with MinValue as the date
    $MinValue = [System.DateTime]::MinValue
    $CurrentScanLastHealthyStateEntry = $MinValue
    $CurrentScanLastUnhealthyStateEntry = $MinValue

    # Read the data from the persistsed state file
    $PreviousScanStates = Read-PersistedLastStateEntries $StateStorePath

    # Assign the data from the persisted state file to variables
    $PreviousScanLastHealthyStateEntry = $PreviousScanStates.LastHealthyState
    $PreviousScanLastUnhealthyStateEntry = $PreviousScanStates.LastUnhealthyState

    $NewHealthyState = $MinValue
    $NewUnhealthyState = $MinValue

    $Message = ""

    # Check: do we know how to determine an unhealthy state?
    if ($HealthyText -ne $null) {

        # Yes - we know what indicates an unhealthy state
        #
        # Retrieve the date time of the last unhealthy state
        $CurrentScanLastUnhealthyStateEntry = Get-LastDateTimeOfStringOccurrenceInLog $LogContents $HealthyText $LogEntryDateTimeFormat

        # Check: do we have a new healthy state?
        if ($CurrentScanLastHealthyStateEntry -ne $MinValue -and $CurrentScanLastHealthyStateEntry -gt $PreviousScanLastHealthyStateEntry) {
            
            # Yes - we have a new healthy state
            #
            # Record the datetime of our new healthy state
            $NewHealthyState = $CurrentScanLastHealthyStateEntry

            $PreviousScanStates.LastHealthyState = $CurrentScanLastHealthyStateEntry
        }
    }

    # Check: do we know how to determine a healthy state?
    if ($UnhealthyText -ne $null) {
        
        # Yes - we know what indicates a healthy state
        #
        # Retrieve the date time of the last healthy state
        $CurrentScanLastUnhealthyStateEntry = Get-LastDateTimeOfStringOccurrenceInLog $LogContents $UnhealthyText $LogEntryDateTimeFormat

        # Check: do we have a new unhealthy state?
        if ($CurrentScanLastUnhealthyStateEntry -ne $MinValue -and $CurrentScanLastUnhealthyStateEntry -gt $PreviousScanLastUnhealthyStateEntry) {
            
            # Yes - we have a new unhealthy state
            #
            # Record the datetime of our new unhealthy state
            $NewUnhealthyState = $CurrentScanLastUnhealthyStateEntry

            $PreviousScanStates.LastUnhealthyState = $CurrentScanLastUnhealthyStateEntry
        }

    }

    # This event id should be changed in the logic that determines what kind of state change has occurred (if any has)
    $EventId = $EventId_ErrorRunningScript
    $EventEntryType = $EventEntryType_ErrorRunningScript

    # Check: did we find any state information in our scan of the file
    if ($NewHealthyState -ne $MinValue -or $NewUnhealthyState -ne $MinValue) {

        # Yes - the log file does contain an indicator of state
        #
        # Check: Is the latest indicator of a healthy state?
        if ($NewHealthyState -gt $NewUnhealthyState) {
            
            # Yes - the most recent indicator was of a healthy state
            $EventId = $EventId_HealthyState
            $EventEntryType = $EventEntryType_HealthyState
            $Message = "Newest indicator was of a healthy state."
            
        } else {

            # No - the most recent indicator was of an unhealthy state
            $EventId = $EventId_UnhealthyState
            $EventEntryType = $EventEntryType_UnhealthState
            $Message = "Newest indicator was of an unhealthy state."
        
        }

        # The previous scan states have been updated when detecting a new scan state
        # these need to now be saved to disk
        Write-PersistedLastStateEntries $StateStorePath $PreviousScanStates

    } else {

        # No - we did not find any new state indicators when we looked
        #
        # No state change has occurred
        $EventId = $EventId_NoStateChange
        $EventEntryType = $EventEntryType_NoStateChange
        $Message = "No new healthy / unhealthy indicator."
    }

    # Check: do we have a positive assertion of a state? (i.e. not no state change)
    if ($EventId -eq $EventId_NoStateChange) {

        # No - we are not asserting healthy / unhealthy. We are asserting
        #      that no new state change has occurred since we last looked
        #
        # Check: has sufficient time elapsed to automatically revent to healthy
        if ($RevertToHealthyEventIdAfterMinutes -ne $null -and ([DateTime]::Now.Subtract($PreviousScanLastUnhealthyStateEntry)).TotalMinutes -ge $RevertToHealthyEventIdAfterMinutes) {
            
            # Yes - enough time has elapsed that we can flag this as healthy
            #
            # Set the event id to the healthy state
            $EventId = $EventId_HealthyState
            $EventEntryType = $EventEntryType_HealthyState
            $Message = "$Message Updated event id to healthy as no unhealthy message was found in the last $RevertToHealthyEventIdAfterMinutes minutes."
        }
    }

    $FullMessage = @"
$Message

File To Check: $LogFilePath
Script Data File: $StateStorePath

Latest Healthy State: $($PreviousScanStates.LastHealthyState.ToString("yyyy-MM-dd hh:mm:ss"))
Latest Unhealthy State: $($PreviousScanStates.LastUnhealthyState.ToString("yyyy-MM-dd hh:mm:ss"))

Healthy Text: $HealthyText
Unhealthy Text: $UnhealthyText
"@

   Write-EventLog -LogName $EventLogName -Source $EventLogSourceName -EntryType $EventEntryType -EventId $EventId -Message $FullMessage

} else {
    
    Write-EventLog -LogName $EventLogName -Source $EventLogSourceName -EntryType $EventEntryType -EventId $EventId_ErrorRunningScript -Message "The log file path $LogFilePath does not exist!"

} 

