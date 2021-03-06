<#
.SYNOPSIS
Gets files older than a specified age.
.DESCRIPTION
The Get-FilesOlderThan function returns files from within a directory and optionally subdirectories that 
are older than a specified period.  The path to begin the search from can be passed as an argument or via
a pipeline.
.INPUTS
System.String
You can pipe a string or an array of strings containing paths to Get-FilesOlderThan        
.PARAMETER Path
Specifies the path the program should begin searching for files from.  Multiple paths can be searched
by passing in an array of paths to search.  If no path is specified the current working directory is used
.PARAMETER Filter
Specifies the file types to be included in the search, by default all file types are included however this
can be filtered by supplying an array of file extensions to include.
.PARAMETER PeriodName
Specifies the name of the period to find files older than.  Accepted values are Seconds, Minutes, Hours,
Days, Months, Years

The PeriodName parameter is mandatory.
.PARAMETER PeriodValue
Specifies the number of 'X' to find files older than, where X is the PeriodName.  For example if PeriodName
is set to Hours and PeriodValue is set to 8 the function will return all files older than 8 hours.

The PeriodValue parameter is mandatory.
.PARAMETER Recurse
Specifies that all subfolders of the search path specified should also be searched when searching for files
older than the period specified.
.LINK
http://www.mywinkb.com/Get-files-older-than-specified-time-PowerShell
.EXAMPLE
C:\PS>Get-FilesOlderThan -Path D:\scratch -PeriodName minutes -PeriodValue 5
This command returns all files within the directory D:\scratch which are older that 5 minutes.
.EXAMPLE
C:\PS>Get-FilesOlderThan -Path D:\scratch -PeriodName hours -PeriodValue 10 -Recurse
This command returns all files older than 10 hours from the directory D:\scratch and all subfolders.
.EXAMPLE
C:\PS>Get-FilesOlderThan -Path D:\scratch -Filter *.txt,*.jpg -PeriodName hours -PeriodValue 10 -Recurse
This command returns all .txt and .jpg files older than 10 hours from within the directory D:\scratch
and all subfolders.
.EXAMPLE
C:\PS>Get-Content D:\scratch\directories.txt | Get-FilesOlderThan -PeriodName days -PeriodValue 1
This command reads a list of directories from a .txt file and then lists all files older than 1 day
in each of the folders listed within the file
#>

Function Get-FilesOlderThan {
    [CmdletBinding()]
    [OutputType([Object])]   
    param (
        [parameter(ValueFromPipeline=$true)]
        [string[]] $Path = (Get-Location),
        [parameter()]
        [string[]] $Filter,
        [parameter(Mandatory=$true)]
        [ValidateSet('Seconds','Minutes','Hours','Days','Months','Years')]
        [string] $PeriodName,
        [parameter(Mandatory=$true)]
        [int] $PeriodValue,
        [parameter()]
        [switch] $Recurse = $false
    )
    
    process {
        
        #If one of more of the paths specified does not exist generate an error  
        if ($(test-path $path) -eq $false) {
            write-error "Cannot find the path: $path because it does not exist"
        }
        
        Else {
        
            <#  
            If the recurse switch is not passed get all files in the specified directories older than the period specified, if no directory is specified then
            the current working directory will be used.
            #>
            If ($recurse -eq $false) {
        
                Get-ChildItem -Path $(Join-Path -Path $Path -ChildPath \*) -Include $Filter | Where-Object { $_.LastWriteTime -lt $(get-date).('Add' + $PeriodName).Invoke(-$periodvalue) `
                -and $_.psiscontainer -eq $false } | `
                #Loop through the results and create a hashtable containing the properties to be added to a custom object
                ForEach-Object {
                    $properties = @{ 
                        Path = $_.Directory 
                        Name = $_.Name 
                        DateModified = $_.LastWriteTime }
                    #Create and output the custom object     
                    New-Object PSObject -Property $properties | select Path,Name,DateModified 
                }                
                  
            } #Close if clause on Recurse conditional
        
            <#  
            If the recurse switch is passed get all files in the specified directories and all subfolders that are older than the period specified, if no directory
            is specified then the current working directory will be used.
            #>   
            Else {
            
                Get-ChildItem  -Path $(Join-Path -Path $Path -ChildPath \*) -Include $Filter -recurse | Where-Object { $_.LastWriteTime -lt $(get-date).('Add' + $PeriodName).Invoke(-$periodvalue) `
                -and $_.psiscontainer -eq $false } | `
                #Loop through the results and create a hashtable containing the properties to be added to a custom object
                ForEach-Object {
                    $properties = @{ 
                        Path = $_.Directory 
                        Name = $_.Name 
                        DateModified = $_.LastWriteTime }
                    #Create and output the custom object     
                    New-Object PSObject -Property $properties | select Path,Name,DateModified 
                }

            } #Close Else clause on recurse conditional       
        } #Close Else clause on Test-Path conditional
    
    } #End Process block
} #End Fuction