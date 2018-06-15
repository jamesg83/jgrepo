# SCOMDWHealthCheck.ps1
# Author: Tao Yang
# Version: 1.1
# SCOM DW Health Check Script

[CmdletBinding()]
PARAM (
	[Parameter(Mandatory=$true,HelpMessage='Please enter the Management Server name')][Alias('DAS','Server','s')][String]$SDK,
	[Parameter(Mandatory=$false,HelpMessage='Please enter the user name to connect to the OpsMgr management group')][Alias('u')][String]$Username = $null,
	[Parameter(Mandatory=$false,HelpMessage='Please enter the password to connect to the OpsMgr management group')][Alias('p')][System.Security.SecureString]$Password = $null,
	[Parameter(Mandatory=$false,HelpMessage='Please enter the report output directory. If not specified, the report files will be written to the script root directory')][Alias('o')][String]$OutputDir = $null,
	[Parameter(Mandatory=$false,HelpMessage='Use -OpenReport switch to automatically open the report HTML page when the script execution is completed.')][Switch]$OpenReport
)
$thisscript = $MyInvocation.MyCommand.path
$ScriptRoot = Split-Path(Resolve-Path $thisscript)
$arrOutputFiles = New-object System.Collections.ArrayList
#region functions
Function Import-OpsMgrSDK
{
<# 
	.Synopsis
	Load OpsMgr 2012 SDK DLLs

	.Description
	Load OpsMgr 2012 SDK DLLs from either the Global Assembly Cache or from the DLLs located in OpsMgrSDK PS module directory. It will use GAC if the DLLs are already loaded in GAC. If all the DLLs have been loaded, a boolean value $true will be returned, otherwise, a boolean value of $false is returned if any there are any errors occurred.

	.Example
	# Load the OpsMgr SDK DLLs
	Import-SDK

#>
	#OpsMgr 2012 R2 SDK DLLs
	$DLLPath = $ScriptRoot
	$arrDLLs = @()
	$arrDLLs += 'Microsoft.EnterpriseManagement.Core.dll'
	$arrDLLs += 'Microsoft.EnterpriseManagement.OperationsManager.dll'
	$arrDLLs += 'Microsoft.EnterpriseManagement.Runtime.dll'
	$DLLVersion = '7.0.5000.0'
	$PublicKeyToken='31bf3856ad364e35'

	#Load SDKs
	$bSDKLoaded = $true
	Foreach ($DLL in $arrDLLs)
	{
		$AssemblyName = $DLL.TrimEnd('.dll')
		#try load from GAC first
		Try {
			Write-Verbose "Trying to load $AssemblyName from GAC..."
			[Void][System.Reflection.Assembly]::Load("$AssemblyName, Version=$DLLVersion, Culture=neutral, PublicKeyToken=$PublicKeyToken")
		} Catch {
			Write-Verbose "Unable to load $AssemblyName from GAC. Trying PowerShell script root folder '$ScriptRoot'..."
			#Can't load from GAC, now try PS module folder
			Try {
				$DLLFilePath = Join-Path $DLLPath $DLL
				Write-Verbose "Loading '$DLLFilePath'"
				[Void][System.Reflection.Assembly]::LoadFrom($DLLFilePath)
			} Catch {
				$_.Exception.InnerException
				Write-Verbose "Unable to load $DLL from either GAC or the Powershell script folder. Please verify if the SDK DLLs exist in at least one location!"
				$bSDKLoaded = $false
			}
		}
	}
	$bSDKLoaded
}

Function Connect-OMManagementGroup
{
<# 
	.Synopsis
	Connect to OpsMgr Management Group using SDK

	.Description
	Connect to OpsMgr Management Group Data Access Service using SDK

	.Parameter -SDK
	Management Server name.

	.Parameter -UserName
	Alternative user name to connect to the management group (optional).

	.Parameter -Password
	Alternative password to connect to the management group (optional).

	.Parameter -DLLPath
	Optionally, specify an alternative path to the OpsMgr SDK DLLs if they have not been installed in GAC.

	.Example
	# Connect to OpsMgr management group via management server "OpsMgrMS01"
	Connect-OMManagementGroup -SDK "OpsMgrMS01"

	.Example
	# Connect to OpsMgr management group via management server "OpsMgrMS01" using different credential
	$Password = ConvertTo-SecureString -AsPlainText "password1234" -force
	$MG = Connect-OMManagementGroup -SDK "OpsMgrMS01" -Username "domain\SCOM.Admin" -Password $Password

	.Example
	# Connect to OpsMgr management group via management server "OpsMgrMS01" using current user's credential
	$MG = Connect-OMManagementGroup -SDK "OpsMgrMS01"
	OR
	$MG = Connect-OMManagementGroup -Server "OPSMGRMS01"
#>
	[CmdletBinding()]
	PARAM (
		[Parameter(Mandatory=$true,HelpMessage='Please enter the Management Server name')][Alias('DAS','Server','s')][String]$SDK,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the user name to connect to the OpsMgr management group')][Alias('u')][String]$Username = $null,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the password to connect to the OpsMgr management group')][Alias('p')][System.Security.SecureString]$Password = $null
	)
	#Check User name and password parameter
	If ($Username)
	{
		If (!$Password)
		{
			Write-Error "Password for user name $Username must be specified!"
		}
	}

	#Try Loadings SDK DLLs in case they haven't been loaded already
	$bSDKLoaded = Import-OpsMgrSDK

	#Connect to the management group
	if ($bSDKLoaded)
	{
		$MGConnSetting = New-Object Microsoft.EnterpriseManagement.ManagementGroupConnectionSettings($SDK)
		If ($Username -and $Password)
		{
			$MGConnSetting.UserName = $Username
			$MGConnSetting.Password = $Password
		}
		$MG = New-Object Microsoft.EnterpriseManagement.ManagementGroup($MGConnSetting)
	}
	$MG
}

Function Get-OMManagementGroupDefaultSettings
{
<# 
	.Synopsis
	Get OpsMgr management group default settings

	.Description
	Get OpsMgr management group default settings via OpsMgr SDK. A System.Collections.ArrayList is returned containing all management group default settings. Each setting in the arraylist is presented in a hashtable format.

	.Parameter -SDK
	Management Server name

	.Parameter -UserName
	Alternative user name to connect to the management group (optional).

	.Parameter -Password
	Alternative password to connect to the management group (optional).

	.Example
	# Connect to OpsMgr management group via management server "OpsMgrMS01" and retrieve all the settings:

	Get-OMManagementGroupDefaultSettings -SDK "OpsMgrMS01"

	.Example
	# Connect to OpsMgr management group via management server "OpsMgrMS01" using alternative credentials and retrieve all the settings:

	$Password = ConvertTo-SecureString -AsPlainText "password1234" -force
	Get-OMManagementGroupDefaultSettings -SDK "OpsMgrMS01" -Username "domain\SCOM.Admin" -Password $Password
#>
	[CmdletBinding()]
	PARAM (
		[Parameter(Mandatory=$true,HelpMessage='Please enter the Management Server name')][Alias('DAS','Server','s')][String]$SDK,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the user name to connect to the OpsMgr management group')][Alias('u')][String]$Username = $null,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the password to connect to the OpsMgr management group')][Alias('p')][System.Security.SecureString]$Password = $null
	)

	#Connect to MG
	Write-Verbose "Connecting to Management Group via SDK $SDK`..."
	If ($Username -and $Password)
	{
		$MG = Connect-OMManagementGroup -SDK $SDK -UserName $Username -Password $Password
	} else {
		$MG = Connect-OMManagementGroup -SDK $SDK
	}

	$Admin = $MG.Administration
	$Settings = $Admin.Settings

	#Get Setting Types
	Write-Verbose 'Get all nested setting types'
	$arrRumtimeTypes = New-Object System.Collections.ArrayList
	$Assembly = [AppDomain]::CurrentDomain.GetAssemblies() |Where-Object { $_.FullName -eq 'Microsoft.EnterpriseManagement.OperationsManager, Version=7.0.5000.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'}
	$SettingType = $assembly.definedtypes | Where-Object{$_.name -eq 'settings'}
	$TopLevelNestedTypes = $SettingType.GetNestedTypes()
	Foreach ($item in $TopLevelNestedTypes)
	{
		if ($item.GetFields().count -gt 0)
		{
			[void]$arrRumtimeTypes.Add($item)
		}
		$NestedTypes = $item.GetNestedTypes()
		foreach ($NestedType in $NestedTypes)
		{
			[void]$arrRumtimeTypes.Add($NestedType)
		}
	}

	#Get Setting Values
	Write-Verbose 'Getting setting values'
	$arrSettingValues = New-Object System.Collections.ArrayList
	Foreach ($item in $arrRumtimeTypes)
	{
		Foreach ($field in $item.GetFields())
		{
			$FieldSetting = $field.GetValue($field.Name)
			$SettingValue = $Settings.GetDefaultValue($FieldSetting)
			$hash = @{
				FieldName = $Field.Name
				Value = $SettingValue
				AllowOverride = $FieldSetting.AllowOverride
				SettingName = $item.Name
				SettingFullName = $item.FullName
			}
			$objSettingValue = New-object psobject -Property $hash
			[void]$arrSettingValues.Add($objSettingValue)
		}
	}
	Write-Verbose "Total number of Management Group default value found: $($arrSettingValues.count)."
	$arrSettingValues
}

Function Get-OMManagementGroupName
{
<# 
	.Synopsis
	Get OpsMgr management group name

	.Description
	Get OpsMgr management group name via OpsMgr SDK. 

	.Parameter -SDK
	Management Server name

	.Parameter -UserName
	Alternative user name to connect to the management group (optional).

	.Parameter -Password
	Alternative password to connect to the management group (optional).

	.Example
	# Connect to OpsMgr management group via management server "OpsMgrMS01" and get the management group name:

	Get-OMManagementGroupName -SDK "OpsMgrMS01"

	.Example
	# Connect to OpsMgr management group via management server "OpsMgrMS01" using alternative credentials and get the management group name:

	$Password = ConvertTo-SecureString -AsPlainText "password1234" -force
	Get-OMManagementGroupName -SDK "OpsMgrMS01" -Username "domain\SCOM.Admin" -Password $Password
#>
	[CmdletBinding()]
	PARAM (
		[Parameter(Mandatory=$true,HelpMessage='Please enter the Management Server name')][Alias('DAS','Server','s')][String]$SDK,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the user name to connect to the OpsMgr management group')][Alias('u')][String]$Username = $null,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the password to connect to the OpsMgr management group')][Alias('p')][System.Security.SecureString]$Password = $null
	)

	#Connect to MG
	Write-Verbose "Connecting to Management Group via SDK $SDK`..."
	If ($Username -and $Password)
	{
		$MG = Connect-OMManagementGroup -SDK $SDK -UserName $Username -Password $Password
	} else {
		$MG = Connect-OMManagementGroup -SDK $SDK
	}

	$MG.Name
}
Function Invoke-SQLQuery
{
<# 
	.Synopsis
	Invoke a SQL query

	.Description
	Invoke a SQL query and conver the returned dataset into an arraylist containing custom PSObject for each record.

	.Parameter -SQLServer
	SQL Server name

	.Parameter -SQLInstance
	SQL Instance Name (if different from the default instance name 'MSSQLSERVER').

	.Parameter -Database
	SQL Database name

	.Parameter -Query
	SQL Query that to be executed against the database.

	.Parameter -UserName
	Alternative user name to connect to the SQL database (optional).

	.Parameter -Password
	Alternative password to connect to the SQL database (optional).

	.Example
	# Connect to Database "SCOMDB01" and run a SQL query listed below to get SQL server version information: "SELECT SERVERPROPERTY('productversion') as 'Version', SERVERPROPERTY ('productlevel') as 'ServicePack', SERVERPROPERTY ('edition') as 'Edition'"

	$SQLQuery = "SELECT SERVERPROPERTY('productversion') as 'Version', SERVERPROPERTY ('productlevel') as 'ServicePack', SERVERPROPERTY ('edition') as 'Edition'"
	Invoke-SQLQuery -SQLServer SCOMDB01 -Database master -Query $SQLQuery

	.Example
	# Connect to Database "SCOMDB01" and run a SQL query listed below to get SQL server version information: "SELECT SERVERPROPERTY('productversion') as 'Version', SERVERPROPERTY ('productlevel') as 'ServicePack', SERVERPROPERTY ('edition') as 'Edition'"using alternative credentials and alternative SQL Query timeout seconds (default is 600 seconds).

	$SQLQuery = "SELECT SERVERPROPERTY('productversion') as 'Version', SERVERPROPERTY ('productlevel') as 'ServicePack', SERVERPROPERTY ('edition') as 'Edition'"
	$Password = ConvertTo-SecureString -AsPlainText "password1234" -force
	Invoke-SQLQuery -SQLServer SCOMDB01 -SQLInstance SCOMDBInstance -SQLPort 12345 -Database master -Query $SQLQuery -SQLQueryTimeout 900 -Username "domain\SCOM.Admin" -Password $Password
#>
	[CmdletBinding()]
	PARAM (
		[Parameter(Mandatory=$true,HelpMessage='Please enter the SQL Server name')][Alias('SQL','Server','s')][String]$SQLServer,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the SQL Instance name')][Alias('Instance','i')][String]$SQLInstance,
        [Parameter(Mandatory=$false,HelpMessage='Please enter the TCP Port number for the SQL Instance')][Alias('port')][int]$SQLPort,
		[Parameter(Mandatory=$true,HelpMessage='Please enter the SQL Database name')][Alias('d')][String]$Database,
		[Parameter(Mandatory=$true,HelpMessage='Please enter the SQL query')][Alias('q')][String]$Query,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the SQL Query Timeout in seconds')][Alias('timeout', 't')][int]$SQLQueryTimeout = 600,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the user name to connect to the SQL Server')][Alias('u')][String]$Username = $null,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the password to connect to the SQL Server')][Alias('p')][System.Security.SecureString]$Password = $null
	)
	
	#Connect to DB
	if ($SQLInstance)
	{
		$SQLServerConn = "$SQLServer`\$SQLInstance"
	} else {
		$SQLServerConn = $SQLServer
	}

	if ($SQLPort)
	{
		$SQLServerConn = "$SQLServerConn`,$SQLPort"
	}

	$conString = "Server`=$SQLServerConn;Integrated Security=true;Initial Catalog=$Database"
    Write-Verbose "SQL connection string: `"$conString`""
	If ($username -and $password)
	{
		#Connecting to SQL by creating a WinRM remote session because we are unable to use a Windows / AD credential in the SQL connection string (only SQL credential is accepted).
		Write-Verbose "Alternative credential specified, establishing a WinRM remote session to $SQLServer to execute the SQL query against DB $Database."
		$SQLServerCred = New-Object System.Management.Automation.PSCredential ($UserName, $Password)
		$DWRemoteSession = New-PSSession -ComputerName $SQLServer -Credential $SQLServerCred
		$arrReturnedData = Invoke-Command -Session $DWRemoteSession -ScriptBlock {
			Param ($SQLServerConn, $Database, $Query, $SQLQueryTimeout)
			$conString = "Server`=$SQLServerConn;Integrated Security=true;Initial Catalog=$Database"
			$SQLCon = New-Object System.Data.SqlClient.SqlConnection
			$SQLCon.ConnectionString = $conString
			$SQLCon.Open()

			#execute SQL query
			$sqlCmd = $SQLCon.CreateCommand()
			$sqlCmd.CommandTimeout=$SQLQueryTimeout
			$sqlCmd.CommandText = $Query
			$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
			$SqlAdapter.SelectCommand = $sqlCmd
			$DataSet = New-Object System.Data.DataSet
			[void]$SqlAdapter.Fill($DataSet)
			$SQLCon.Close()

			#Process result
			$arrReturnedData = New-Object System.Collections.ArrayList
			Foreach ($set in $DataSet.Tables[0])
			{
				$objDS = New-object psobject
				Foreach ($objProperty in (Get-member -InputObject $set -MemberType Property))
				{
					$PropertyName = $objProperty.Name
					Add-Member -InputObject $objDS -MemberType NoteProperty -Name $PropertyName -Value $set.$PropertyName
				}
				[void]$arrReturnedData.Add($objDS)
			}
			#Return the array list
			,$arrReturnedData
		} -ArgumentList $SQLServerConn, $Database, $Query, $SQLQueryTimeout
		Remove-PSSession -Session $DWRemoteSession
	} else {
		#Connecting to SQL using integrated security
		Write-Verbose "Connecting to SQL DB $Database on server $SQLServerConn using integrated security."
		$conString = "Server`=$SQLServerConn;Integrated Security=true;Initial Catalog=$Database"
		$SQLCon = New-Object System.Data.SqlClient.SqlConnection
		$SQLCon.ConnectionString = $conString
		$SQLCon.Open()

		#execute SQL query
		$sqlCmd = $SQLCon.CreateCommand()
		$sqlCmd.CommandTimeout=$SQLQueryTimeout
		$sqlCmd.CommandText = $Query
		$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
		$SqlAdapter.SelectCommand = $sqlCmd
		$DataSet = New-Object System.Data.DataSet
		[void]$SqlAdapter.Fill($DataSet)
		$SQLCon.Close()

		#Process result
		$arrReturnedData = New-Object System.Collections.ArrayList

		Foreach ($set in $DataSet.Tables[0])
		{
			$objDS = New-object psobject
			Foreach ($objProperty in (Get-member -InputObject $set -MemberType Property))
			{
				$PropertyName = $objProperty.Name
				Add-Member -InputObject $objDS -MemberType NoteProperty -Name $PropertyName -Value $set.$PropertyName
			}
			[void]$arrReturnedData.Add($objDS)
		}
	}
	#Return the array list
	Write-Verbose "Number of rows returned: $($arrReturnedData.count)"
	,$arrReturnedData
}
Function Check-UserPrivilege
{
<# 
	.Synopsis
	Check if an user account has a particular privilege

	.Description
	Check if an user account has a particular privilege on the local or a remote computer. When checking on a remote computer, WinRM must be enabled on the remote computer. This function uses the C# code from Carbon PowerShell module, Copyright 2012 Aaron Jensen (https://bitbucket.org/splatteredbits/carbon) under the Apache license version 2.0

	.Parameter -UserAccount
	User Account that needs to check the privilege against

	.Parameter -Privilege
	The privilege to check the user account against

	.Parameter -ComputerName
	Specify the computer name if checking on a remote computer. Note: WinRM (PowerShell Remoting) must be enabled in order to perform this function against the remote computer

	.Parameter -UserName
	Alternative user name to connect to the remote computer via WinRM (optional).

	.Parameter -Password
	Alternative password to connect to the remote computer via WinRM (optional).

	.Example
	# Check if "Domain\User" has "SeLockMemoryPrivilege" (Lock Page in Memory) privilege
	Check-UserPrivilege -UserAccount "Domain\User" -Privilege "SeLockMemoryPrivilege"

	.Example
	# Check if "Domain\User" has "SeLockMemoryPrivilege" (Lock Page in Memory) privilege on the remote computer "Computer01.domain.com"
	Check-UserPrivilege -UserAccount "Domain\User" -Privilege "SeLockMemoryPrivilege" -ComputerName "Computer01.domain.com"

	.Example
	# Check if "Domain\User" has "SeLockMemoryPrivilege" (Lock Page in Memory) privilege on the remote computer "Computer01.domain.com" using alternative username and password
	$Password = ConvertTo-SecureString -String "password1234" -AsPlainText -Force
	Check-UserPrivilege -UserAccount "Domain\User" -Privilege "SeLockMemoryPrivilege" -ComputerName "Computer01.domain.com" -Username "Domain\Admin" -Password $Password
#>
	[CmdletBinding()]
	PARAM (
		[Parameter(Mandatory=$true,HelpMessage='Please enter the user account to check the privilege against')][Alias('a')][String]$UserAccount,
		[Parameter(Mandatory=$true,HelpMessage='Please enter the the privilege')][Alias('r')][String]$Privilege,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the the Remote Computer name')][Alias('c')][String]$ComputerName,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the user name to connect to the remote computer')][Alias('u')][String]$Username = $null,
		[Parameter(Mandatory=$false,HelpMessage='Please enter the password to connect to the remote computer')][Alias('p')][System.Security.SecureString]$Password = $null
	)
$SourceCode = @"
/*
Original sources available at: https://bitbucket.org/splatteredbits/carbon
*/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
namespace PSCarbon
{
    public sealed class Lsa
    {
        // ReSharper disable InconsistentNaming
        [StructLayout(LayoutKind.Sequential)]
        internal struct LSA_UNICODE_STRING
        {
            internal LSA_UNICODE_STRING(string inputString)
            {
                if (inputString == null)
                {
                    Buffer = IntPtr.Zero;
                    Length = 0;
                    MaximumLength = 0;
                }
                else
                {
                    Buffer = Marshal.StringToHGlobalAuto(inputString);
                    Length = (ushort)(inputString.Length * UnicodeEncoding.CharSize);
                    MaximumLength = (ushort)((inputString.Length + 1) * UnicodeEncoding.CharSize);
                }
            }
            internal ushort Length;
            internal ushort MaximumLength;
            internal IntPtr Buffer;
        }
        [StructLayout(LayoutKind.Sequential)]
        internal struct LSA_OBJECT_ATTRIBUTES
        {
            internal uint Length;
            internal IntPtr RootDirectory;
            internal LSA_UNICODE_STRING ObjectName;
            internal uint Attributes;
            internal IntPtr SecurityDescriptor;
            internal IntPtr SecurityQualityOfService;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct LUID
        {
            public uint LowPart;
            public int HighPart;
        }
        // ReSharper disable UnusedMember.Local
        private const uint POLICY_VIEW_LOCAL_INFORMATION = 0x00000001;
        private const uint POLICY_VIEW_AUDIT_INFORMATION = 0x00000002;
        private const uint POLICY_GET_PRIVATE_INFORMATION = 0x00000004;
        private const uint POLICY_TRUST_ADMIN = 0x00000008;
        private const uint POLICY_CREATE_ACCOUNT = 0x00000010;
        private const uint POLICY_CREATE_SECRET = 0x00000014;
        private const uint POLICY_CREATE_PRIVILEGE = 0x00000040;
        private const uint POLICY_SET_DEFAULT_QUOTA_LIMITS = 0x00000080;
        private const uint POLICY_SET_AUDIT_REQUIREMENTS = 0x00000100;
        private const uint POLICY_AUDIT_LOG_ADMIN = 0x00000200;
        private const uint POLICY_SERVER_ADMIN = 0x00000400;
        private const uint POLICY_LOOKUP_NAMES = 0x00000800;
        private const uint POLICY_NOTIFICATION = 0x00001000;
        // ReSharper restore UnusedMember.Local
        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool LookupPrivilegeValue(
            [MarshalAs(UnmanagedType.LPTStr)] string lpSystemName,
            [MarshalAs(UnmanagedType.LPTStr)] string lpName,
            out LUID lpLuid);
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
        private static extern uint LsaAddAccountRights(
            IntPtr PolicyHandle,
            IntPtr AccountSid,
            LSA_UNICODE_STRING[] UserRights,
            uint CountOfRights);
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = false)]
        private static extern uint LsaClose(IntPtr ObjectHandle);
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern uint LsaEnumerateAccountRights(IntPtr PolicyHandle,
            IntPtr AccountSid,
            out IntPtr UserRights,
            out uint CountOfRights
            );
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern uint LsaFreeMemory(IntPtr pBuffer);
        [DllImport("advapi32.dll")]
        private static extern int LsaNtStatusToWinError(long status);
        [DllImport("advapi32.dll", SetLastError = true, PreserveSig = true)]
        private static extern uint LsaOpenPolicy(ref LSA_UNICODE_STRING SystemName, ref LSA_OBJECT_ATTRIBUTES ObjectAttributes, uint DesiredAccess, out IntPtr PolicyHandle );
        [DllImport("advapi32.dll", SetLastError = true, PreserveSig = true)]
        static extern uint LsaRemoveAccountRights(
            IntPtr PolicyHandle,
            IntPtr AccountSid,
            [MarshalAs(UnmanagedType.U1)]
            bool AllRights,
            LSA_UNICODE_STRING[] UserRights,
            uint CountOfRights);
        // ReSharper restore InconsistentNaming
        private static IntPtr GetIdentitySid(string identity)
        {
            var sid =
                new NTAccount(identity).Translate(typeof (SecurityIdentifier)) as SecurityIdentifier;
            if (sid == null)
            {
                throw new ArgumentException(string.Format("Account {0} not found.", identity));
            }
            var sidBytes = new byte[sid.BinaryLength];
            sid.GetBinaryForm(sidBytes, 0);
            var sidPtr = Marshal.AllocHGlobal(sidBytes.Length);
            Marshal.Copy(sidBytes, 0, sidPtr, sidBytes.Length);
            return sidPtr;
        }
        private static IntPtr GetLsaPolicyHandle()
        {
            var computerName = Environment.MachineName;
            IntPtr hPolicy;
            var objectAttributes = new LSA_OBJECT_ATTRIBUTES
            {
                Length = 0,
                RootDirectory = IntPtr.Zero,
                Attributes = 0,
                SecurityDescriptor = IntPtr.Zero,
                SecurityQualityOfService = IntPtr.Zero
            };
            const uint ACCESS_MASK = POLICY_CREATE_SECRET | POLICY_LOOKUP_NAMES | POLICY_VIEW_LOCAL_INFORMATION;
            var machineNameLsa = new LSA_UNICODE_STRING(computerName);
            var result = LsaOpenPolicy(ref machineNameLsa, ref objectAttributes, ACCESS_MASK, out hPolicy);
            HandleLsaResult(result);
            return hPolicy;
        }
        public static string[] GetPrivileges(string identity)
        {
            var sidPtr = GetIdentitySid(identity);
            var hPolicy = GetLsaPolicyHandle();
            var rightsPtr = IntPtr.Zero;
            try
            {
                var privileges = new List<string>();
                uint rightsCount;
                var result = LsaEnumerateAccountRights(hPolicy, sidPtr, out rightsPtr, out rightsCount);
                var win32ErrorCode = LsaNtStatusToWinError(result);
                // the user has no privileges
                if( win32ErrorCode == STATUS_OBJECT_NAME_NOT_FOUND )
                {
                    return new string[0];
                }
                HandleLsaResult(result);
                var myLsaus = new LSA_UNICODE_STRING();
                for (ulong i = 0; i < rightsCount; i++)
                {
                    var itemAddr = new IntPtr(rightsPtr.ToInt64() + (long) (i*(ulong) Marshal.SizeOf(myLsaus)));
                    myLsaus = (LSA_UNICODE_STRING) Marshal.PtrToStructure(itemAddr, myLsaus.GetType());
                    var cvt = new char[myLsaus.Length/UnicodeEncoding.CharSize];
                    Marshal.Copy(myLsaus.Buffer, cvt, 0, myLsaus.Length/UnicodeEncoding.CharSize);
                    var thisRight = new string(cvt);
                    privileges.Add(thisRight);
                }
                return privileges.ToArray();
            }
            finally
            {
                Marshal.FreeHGlobal(sidPtr);
                var result = LsaClose(hPolicy);
                HandleLsaResult(result);
                result = LsaFreeMemory(rightsPtr);
                HandleLsaResult(result);
            }
        }
        public static void GrantPrivileges(string identity, string[] privileges)
        {
            var sidPtr = GetIdentitySid(identity);
            var hPolicy = GetLsaPolicyHandle();
            try
            {
                var lsaPrivileges = StringsToLsaStrings(privileges);
                var result = LsaAddAccountRights(hPolicy, sidPtr, lsaPrivileges, (uint)lsaPrivileges.Length);
                HandleLsaResult(result);
            }
            finally
            {
                Marshal.FreeHGlobal(sidPtr);
                var result = LsaClose(hPolicy);
                HandleLsaResult(result);
            }
        }
        const int STATUS_SUCCESS = 0x0;
        const int STATUS_OBJECT_NAME_NOT_FOUND = 0x00000002;
        const int STATUS_ACCESS_DENIED = 0x00000005;
        const int STATUS_INVALID_HANDLE = 0x00000006;
        const int STATUS_UNSUCCESSFUL = 0x0000001F;
        const int STATUS_INVALID_PARAMETER = 0x00000057;
        const int STATUS_NO_SUCH_PRIVILEGE = 0x00000521;
        const int STATUS_INVALID_SERVER_STATE = 0x00000548;
        const int STATUS_INTERNAL_DB_ERROR = 0x00000567;
        const int STATUS_INSUFFICIENT_RESOURCES = 0x000005AA;
        private static readonly Dictionary<int, string> ErrorMessages = new Dictionary<int, string>
                                    {
                                        {STATUS_OBJECT_NAME_NOT_FOUND, "Object name not found. An object in the LSA policy database was not found. The object may have been specified either by SID or by name, depending on its type."},
                                        {STATUS_ACCESS_DENIED, "Access denied. Caller does not have the appropriate access to complete the operation."},
                                        {STATUS_INVALID_HANDLE, "Invalid handle. Indicates an object or RPC handle is not valid in the context used."},
                                        {STATUS_UNSUCCESSFUL, "Unsuccessful. Generic failure, such as RPC connection failure."},
                                        {STATUS_INVALID_PARAMETER, "Invalid parameter. One of the parameters is not valid."},
                                        {STATUS_NO_SUCH_PRIVILEGE, "No such privilege. Indicates a specified privilege does not exist."},
                                        {STATUS_INVALID_SERVER_STATE, "Invalid server state. Indicates the LSA server is currently disabled."},
                                        {STATUS_INTERNAL_DB_ERROR, "Internal database error. The LSA database contains an internal inconsistency."},
                                        {STATUS_INSUFFICIENT_RESOURCES, "Insufficient resources. There are not enough system resources (such as memory to allocate buffers) to complete the call."}
                                    };
        private static void HandleLsaResult(uint returnCode)
        {
            var win32ErrorCode = LsaNtStatusToWinError(returnCode);
            if( win32ErrorCode == STATUS_SUCCESS)
                return;
            if( ErrorMessages.ContainsKey(win32ErrorCode) )
            {
                throw new Win32Exception(win32ErrorCode, ErrorMessages[win32ErrorCode]);
            }
            throw new Win32Exception(win32ErrorCode);
        }
        public static void RevokePrivileges(string identity, string[] privileges)
        {
            var sidPtr = GetIdentitySid(identity);
            var hPolicy = GetLsaPolicyHandle();
            try
            {
                var currentPrivileges = GetPrivileges(identity);
                if (currentPrivileges.Length == 0)
                {
                    return;
                }
                var lsaPrivileges = StringsToLsaStrings(privileges);
                var result = LsaRemoveAccountRights(hPolicy, sidPtr, false, lsaPrivileges, (uint)lsaPrivileges.Length);
                HandleLsaResult(result);
            }
            finally
            {
                Marshal.FreeHGlobal(sidPtr);
                var result = LsaClose(hPolicy);
                HandleLsaResult(result);
            }
        }
        private static LSA_UNICODE_STRING[] StringsToLsaStrings(string[] privileges)
        {
            var lsaPrivileges = new LSA_UNICODE_STRING[privileges.Length];
            for (var idx = 0; idx < privileges.Length; ++idx)
            {
                lsaPrivileges[idx] = new LSA_UNICODE_STRING(privileges[idx]);
            }
            return lsaPrivileges;
        }
    }
}
"@
	If ($ComputerName)
	{
		#Checking against a remote computer
		If ($Username -and $Password)
		{
			#Create a WinRM session using alternative credential
			$RemoteComputerCred = New-Object System.Management.Automation.PSCredential ($Username, $Password)
			$RemoteSession = New-PSSession -ComputerName $ComputerName -Credential $RemoteComputerCred
		}else {
			$RemoteSession = New-PSSession -ComputerName $ComputerName
		}
		$bUserHasAccess = Invoke-Command -Session $RemoteSession -ScriptBlock {
			param ($UserAccount, $Privilege, $SourceCode)
			Add-Type -TypeDefinition $SourceCode -Language CSharp
			if ([PSCarbon.Lsa]::GetPrivileges($UserAccount).Contains($Privilege))
			{
				$bUserHasAccess = $true
			} else {
				$bUserHasAccess = $false
			}
			$bUserHasAccess
		} -ArgumentList $UserAccount, $Privilege, $SourceCode
		#Housekeeping - remove the remote session
		Remove-PSSession $RemoteSession
	} else {
		#Checking against local computer
		Add-Type -TypeDefinition $SourceCode -Language CSharp
		if ([PSCarbon.Lsa]::GetPrivileges($UserAccount).Contains($Privilege))
		{
			$bUserHasAccess = $true
		} else {
			$bUserHasAccess = $false
		}

	}
	$bUserHasAccess
}

function Test-PsRemoting 
{ 
    #The original version of this function is taken from: http://www.leeholmes.com/blog/2009/11/20/testing-for-powershell-remoting-test-psremoting/
	param( 
        [Parameter(Mandatory = $true)][String]$ComputerName,
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential
    ) 
    
    try 
    { 
        $errorActionPreference = "Stop"
		if ($Credential)
		{
			$result = Invoke-Command -ComputerName $computername { 1 } -Credential $Credential
		} else {
			$result = Invoke-Command -ComputerName $computername { 1 } 
		}
        
    } 
    catch 
    { 
        Write-Verbose $_ 
        return $false 
    } 
    
    ## I’ve never seen this happen, but if you want to be 
    ## thorough…. 
    if($result -ne 1) 
    { 
        Write-Verbose "Remoting to $computerName returned an unexpected result." 
        return $false 
    } 
    
    $true    
}
#endregion

#region initiation
$now = Get-Date -Format yyyy-MMM-dd
#Create a PSCredential object if username and password are specified
If ($Username -and $Password)
{
	$PSCredential = New-Object System.Management.Automation.PSCredential ($UserName, $Password)
}

#Pre-requisites check - Make sure all requirements are met before continuing with the script
Write-verbose "Performing Pre-requisites checks."
$bPrereqPassed = $true
#Check PowerShell version. must be at least v.3
$PoShMajorVersion = $Host.Version.Major
If ($PoShMajorVersion -lt 3)
{
	Write-Error "The PowerShell version on this computer is $($host.Verion.ToString()). The minimum version required by this script is version 3.0."
	$bPrereqPassed = $false
}
#Load SDK
Write-verbose "Pre-requisite check - loading OpsMgr 2012 SDK assemblies."
If (Import-OpsMgrSDK)
{
	#Connect to the SCOM management group and retrieve some basic information
	Write-verbose "Pre-requisite check - Connecting to the OpsMgr management group via management server $SDK."
	If ($Username -and $Password)
	{
		Try {
			$MG = Connect-OMManagementGroup -SDK $SDK -UserName $Username -Password $Password
		} Catch {
			$bPrereqPassed = $false
			Write-Error $_.Exception.InnerException.Message
		}
		
	} else {
		Try {
			$MG = Connect-OMManagementGroup -SDK $SDK
		} Catch {
			$bPrereqPassed = $false
			Write-Error $_.Exception.InnerException.Message
		}
	}
	If ($MG -eq $null)
	{
		$bPrereqPassed = $false
		Write-Error "Unable to connect to OpsMgr SDK via management server $SDK"
	}
} else {
	$bPrereqPassed = $false
	Write-Error "Unable to load OpsMgr 2012 SDK assemblies. Please make sure the 3 SDK DLLs are copied to the same folder as this script, or try to run this script on a management server, a web console server, or a computer that has OpsMgr 2012 console installed."
}
If ($bPrereqPassed -eq $true)
{
	#Get SCOM MG Default Settings
	Write-verbose "Pre-requisite check - Retrieving OpsMgr management group's default settings and testing WinRM connection to the Data Warehouse SQL Server."
	If ($Username -and $Password)
	{
		$MGSettings = Get-OMManagementGroupDefaultSettings -SDK $SDK -UserName $Username -Password $Password
	} else {
		$MGSettings = Get-OMManagementGroupDefaultSettings -SDK $SDK
	}

	#Get DW Server name and Database name
	$DWServerAddress = ($MGSettings | Where-Object {$_.SettingName -eq "DataWarehouse" -and $_.FieldName -eq "DataWarehouseServername"}).Value
	$DWDatabaseName = ($MGSettings | Where-Object {$_.SettingName -eq "DataWarehouse" -and $_.FieldName -eq "DataWarehouseDatabaseName"}).Value
	$DWSQLAddress = $DWServerAddress.split("\")[1]
	$DWSQLPort = $DWServerAddress.split(",")[1]
	$DWServerName = $DWServerAddress.split(",")[0].split("\")[0]
	Write-Verbose "DW Server Name: $DWServerName"
	$DWSQLAddress = $DWServerAddress.split("\")[1]
	$DWSQLPort = $DWServerAddress.split(",")[1]
	If ($DWSQLAddress -ne $null)
	{
		$DWSQLInstance = $DWSQLAddress.split(",")[0]
		Write-Verbose "DW SQL Instance: $DWSQLInstance"
	} else {
		Write-Verbose "DW SQL Instance: The Default Instance"
	}
	
	
	If ($DWSQLPort -ne $null)
	{
		Write-Verbose "DW SQL Server Port: $DWSQLPort"
	} else {
		Write-Verbose "DW SQL Server Port: Default or dynamic port"
	}

	Write-Verbose "DW SQL Database: $DWDatabaseName"

	#Test WinRM on DW SQL Server
	If ($Username -and $password)
	{
		$bDWRemotingEnabled = Test-PsRemoting -Computer $DWServerName -Credential $PSCredential
	} else {
		$bDWRemotingEnabled = Test-PsRemoting -Computer $DWServerName
	}
	If (!$bDWRemotingEnabled)
	{
		$bPrereqPassed = $false
		Write-Error "Unable to establish WinRM (PS Remoting) session to the OpsMgr Dataware House SQL server $DWServerName. Please make sure WinRM is enabled and properly configured."
	}
}
#Test Connectivities to the management servers
$Admin = $MG.Administration
$arrMgmtServers = @($Admin.GetAllManagementServers() | Where-Object {$_.IsGateway -eq $false})
$MgmtServersCount = $arrMgmtServers.count
Foreach ($MgmtServer in $arrMgmtServers)
{
	$MSFQDN = $MgmtServer.PrincipalName
	Write-Verbose "Testing remote WMI connectivity to management server `"$MSFQDN`""
	If ($Username -and $Password)
	{
		$TestWMI = Get-WmiObject -ComputerName $MSFQDN -Query "Select * from Win32_OperatingSystem" -Credential $PSCredential -ErrorAction SilentlyContinue -ErrorVariable TestWMIError
	} else {
		$TestWMI = Get-WmiObject -ComputerName $MSFQDN -Query "Select * from Win32_OperatingSystem" -ErrorAction SilentlyContinue -ErrorVariable TestWMIError
	}
	If ($TestWMIError.count -gt 0)
	{
		Write-Error "Unable to remotely query WMI on $MSFQDN. Error: $($TestWMIError[0].Exception)"
		$bPrereqPassed = $false
	}
	Write-Verbose "Testing Remote Windows Event Log connectivity to management server `"$MSFQDN`"" 
	If ($Username -and $Password)
	{
		Try {
			$TestWinEvent = Get-WinEvent -ListLog "Operations Manager" -ComputerName $MSFQDN -Credential $PSCredential -ErrorAction SilentlyContinue -ErrorVariable TestWinEventError
		} Catch {
			Write-Error "Unable to remotely query Windows Event log 'Operations Manager' on management server `"$MSFQDN`". Error: $($_.Exception.message)"
			$bPrereqPassed = $false
		}
		
	} else {
		Try {
			$TestWinEvent = Get-WinEvent -ListLog "Operations Manager" -ComputerName $MSFQDN -ErrorAction SilentlyContinue -ErrorVariable TestWinEventError
		} Catch {
			Write-Error "Unable to remotely query Windows Event log 'Operations Manager' on management server `"$MSFQDN`". Error: $($_.Exception.message)"
			$bPrereqPassed = $false
		}
	}
	If ($TestWMIError) {Remove-Variable TestWMIError}
}

if (!$bPrereqPassed)
{
	Write-Output "Pre-requisites check failed. This script is unable to continue. Please make sure all pre-requisites are met and then re-run the script."
	Exit
} else {
	Write-Verbose "All pre-requisites checks passed. continue running this script."
}

#Check WinRM configuration by creating a PS Remote session to the DW SQL server.
$resultTemplateXML = Join-Path $scriptRoot "DWHealthCheckResultTemplate.xml"
$OutputFileName = "DWHealthCheckResult-$now"
If ($OutputDir)
{
	If (Test-Path $OutputDir)
	{
		Write-Verbose "The Report files will be written to '$OutputDir'"
	} else {
		Write-Warning "The specified Output directory '$OutputDir' is not valid. Report files will be written to the script root folder '$ScriptRoot'"
		$OutputDir = $ScriptRoot
	}
} else {
	Write-Verbose "The Output directory is not specified. Report files will be written to the script root folder '$ScriptRoot'"
	$OutputDir = $ScriptRoot
}
$outputXML = Join-Path $OutputDir "$OutputFileName.xml"
$outputHTML = Join-Path $OutputDir "$OutputFileName.html"
#endregion

#region Environment Discovery
#Get Current date time
#Get current date time
$Date = Get-Date
$strDate = "$($Date.day)-$($Date.month)-$($Date.Year)"

If ($Username -and $Password)
{
	$DWServerDomain = (Get-WmiObject -ComputerName $DWServerName -Query "Select Domain from Win32_ComputerSystem" -Credential $PSCredential).Domain
	$DWServerNetBiosName = (Get-WmiObject -ComputerName $DWServerName -Query "Select name from Win32_ComputerSystem" -Credential $PSCredential).Name
} else {
	$DWServerDomain = (Get-WmiObject -ComputerName $DWServerName -Query "Select Domain from Win32_ComputerSystem").Domain
	$DWServerNetBiosName = (Get-WmiObject -ComputerName $DWServerName -Query "Select name from Win32_ComputerSystem").Name
}

$DWServerFQDN = "$DWServerNetBiosName`.$DWServerDomain"
#endregion

#region Define SQL Queries
Write-Verbose "Defining all SQL queries used in this script..."
#DW Dataset back log
$DWAggrSQLQuery = "Select (Case AggregationTypeId When 20 then 'Hourly' When 30 Then 'Daily' When 0 Then 'Raw' End) as AggregationType, Datasetid, (Select DatasetDefaultName From DataSet Where Datasetid = StandardDataSetAggregationHistory.Datasetid) as DatasetName ,  COUNT(*) as 'Count', MIN(AggregationDateTime) as 'First', MAX(AggregationDateTime) as 'Last' From StandardDataSetAggregationHistory Where LastAggregationDurationSeconds IS NULL group by AggregationTypeId, Datasetid"

#DW Staging Tables Row Counts
$AlertStagingCountSQL = "SELECT count(*) as 'count' from Alert.AlertStage"
$EventStagingCountSQL = "SELECT count(*) as 'count' from Event.EventStage"
$PerfStagingCountSQL = "SELECT count(*) as 'count' from Perf.PerformanceStage"
$StateStagingCountSQL = "SELECT count(*) as 'count' from State.StateStage"
$ManagedEntityStagingCountSQL = "select COUNT(*) as 'count' from ManagedEntityStage"

#SQL Version
$SQLVersionQuery = "select SubString(@@VERSION,1,Charindex(' - ',@@Version)-1) As 'Version'"

#SQL Property
$SQLPropertyQuery = "SELECT SERVERPROPERTY('productversion') as 'Version', SERVERPROPERTY ('productlevel') as 'ServicePack', SERVERPROPERTY ('edition') as 'Edition'"

#SQL Configuration
$SQLConfigQuery = "SELECT * FROM sys.configurations order by name"

#SQL Server Collation
$SQLCollationQuery = "SELECT CONVERT (varchar, SERVERPROPERTY('collation')) as ServerCollation"

#SQL Service accounts
$SQLServiceAccountQuery = "select * from sys.dm_server_services where servicename LIKE 'SQL Server (%'"

#DW Database property query
$DWDBPropertyQuery = "SELECT * FROM sys.databases WHERE name='$DWDatabaseName'"

#DW Database last backup
$DWDBLastBckpQuery = "SELECT top 1 * FROM dbo.backupset where database_name = '$DWDatabaseName' order by backup_start_date desc"

#DW Database Auto Growth
$DBGrowthQuery = @"
SELECT
S.[name] AS [LogicalName]
,S.[file_id] AS [FileID]
, S.[physical_name] AS [FileName]
,CAST(CAST(G.name AS VARBINARY(256)) AS sysname) AS [FileGroupName]
,CONVERT (varchar(10),(S.[size]/128)) AS [SizeMB]
,CASE WHEN S.[max_size]=-1 THEN 'Unlimited' ELSE CONVERT(VARCHAR(10),CONVERT(bigint,S.[max_size])/128) END AS [MaxSizeMB]
,CASE s.is_percent_growth WHEN 1 THEN CONVERT(VARCHAR(10),S.growth) +'%' ELSE Convert(VARCHAR(10),S.growth/128) + ' MB' END AS [Growth]
,Case WHEN S.[type]=0 THEN 'Data Only'
WHEN S.[type]=1 THEN 'Log Only'
WHEN S.[type]=2 THEN 'FILESTREAM Only'
WHEN S.[type]=3 THEN 'Informational purposes Only'
WHEN S.[type]=4 THEN 'Full-text '
END AS [Usage]
,DB_name(S.database_id) AS [DatabaseName]
FROM sys.master_files AS S
LEFT JOIN sys.filegroups AS G ON ((S.type = 2 OR S.type = 0)
AND (S.drop_lsn IS NULL)) AND (S.data_space_id=G.data_space_id) Where DB_name(S.database_id) = `'{0}`' ORDER by Usage
"@

#SQL Perf counters
#Buffer Manager\Buffer Cache hit ratio
$SQLPerfBCHRQuery = "SELECT * FROM sys.dm_os_performance_counters where object_name LIKE '%:Buffer Manager%' and counter_name LIKE 'Buffer cache hit ratio%'"

#Buffer Manager\Page Life Expectancy
$SQLPerfPLEQuery = "SELECT * FROM sys.dm_os_performance_counters where object_name LIKE '%:Buffer Manager%' and counter_name = 'Page Life Expectancy'"

#SQL DB Data free space
$SQLDWDBFreeSpaceQuery = "SELECT DB_NAME() AS DbName, name AS FileName, size/128.0 AS CurrentSizeMB, size/128.0 - CAST(FILEPROPERTY(name, 'SpaceUsed') AS INT)/128.0 AS FreeSpaceMB FROM sys.database_files where type_desc = 'ROWS'"

#SQL DB Log free space
$SQLDWLogFreeSpaceQuery = "SELECT DB_NAME() AS DbName, name AS FileName, size/128.0 AS CurrentSizeMB, size/128.0 - CAST(FILEPROPERTY(name, 'SpaceUsed') AS INT)/128.0 AS FreeSpaceMB FROM sys.database_files where type_desc = 'LOG'"

#OpsMgr DW DataSets List
$SQLDWDataSetListQuery = "EXEC StandardDatasetAggregationSizeList"
#endregion

#region Data Collection
#Get DW SQL Server Operating System and Hardware configuration
If ($Username -and $Password)
{
	$objWMIComputerSystem = Get-WmiObject -Computer $DWServerName -Query "Select * From Win32_ComputerSystem" -Credential $PSCredential
	$objWMIOS = Get-WmiObject -Computer $DWServerName -Query "Select * From Win32_OperatingSystem" -Credential $PSCredential
	$objWMICPU = Get-WmiObject -Computer $DWServerName -Query "Select * From Win32_Processor" -Credential $PSCredential
} else {
	$objWMIComputerSystem = Get-WmiObject -Computer $DWServerName -Query "Select * From Win32_ComputerSystem"
	$objWMIOS = Get-WmiObject -Computer $DWServerName -Query "Select * From Win32_OperatingSystem"
	$objWMICPU = Get-WmiObject -Computer $DWServerName -Query "Select * From Win32_Processor"
}

$DWServerManufacturer = $objWMIComputerSystem.Manufacturer
$DWServerModel = $objWMIComputerSystem.Model
$DWServerOSCaption = $objWMIOS.Caption
$DWServerOSVersion = $objWMIOS.Version
$DWServerOSArchitecture = $objWMIOS.OSArchitecture
$DWServerTotalMemoryMB = [math]::Round($objWMIOS.TotalVisibleMemorySize/1024)
If ($objWMICPU.GetType().BaseType.Name -eq 'Array')
{
	$DWServerCPUSpeed = $objWMICPU[0].maxclockspeed
	$DWServerCPUAddressWidth = $objWMICPU[0].AddressWidth
	$DWServerCPUCoreNumber = 0
	$DWServerCPULogicalProcessorNumber = 0
	Foreach ($item in $objWMICPU)
	{
		$DWServerCPUCoreNumber = $DWServerCPUCoreNumber + $item.numberOfCores
		$DWServerCPULogicalProcessorNumber = $DWServerCPULogicalProcessorNumber + $item.NumberOfLogicalProcessors
	}
} else {
	$DWServerCPUSpeed = $objWMICPU.maxclockspeed
	$DWServerCPUAddressWidth = $objWMICPU.AddressWidth
	$DWServerCPUCoreNumber = $objWMICPU.numberOfCores
	$DWServerCPULogicalProcessorNumber = $objWMICPU.NumberOfLogicalProcessors
}



$MGVersion = $MG.Version.ToString()
$MGName = $MG.Name

$RMSE = $Admin.GetRootManagementServer().Name
$MGSummary = $Admin.GetSummary()
$iWindowsAgentCount = $MGSummary.AgentManagedComputerCount
$iAgentlessManagedCount = $MGSummary.RemotelyManagedComputerCount
$iNetworkDevicesCount = $Admin.GetAllRemotelyManagedDevices().Count

$GWServersCount = @($Admin.GetAllManagementServers() | Where-Object {$_.IsGateway -eq $true}).Count
#Get Unix computer count
$UnixClass = $MG.GetMonitoringClasses("Microsoft.Unix.Computer")[0]
$UnixComputers = $MG.GetMonitoringObjects($UnixClass)
$iUnixComputerCount = $UnixComputers.Count
#Get Operational DB info
$hklm = 2147483650
$OpsMgrSetupKey = "SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\Setup"
$RegOpsDBNamevalue = "DatabaseName"
$RegOpsDBServerNameValue = "DatabaseServerName"
If ($Username -and $Password)
{
	$OpsMgrMSRegWmi = get-wmiobject -list "StdRegProv" -namespace root\default -Computername $SDK -Credential $PSCredential
} else {
	$OpsMgrMSRegWmi = get-wmiobject -list "StdRegProv" -namespace root\default -Computername $SDK
}

$OpsDBName = $OpsMgrMSRegWmi.GetStringValue($hklm, $OpsMgrSetupKey, $RegOpsDBNamevalue).sValue
$OpsDBServerAddress = $OpsMgrMSRegWmi.GetStringValue($hklm, $OpsMgrSetupKey, $RegOpsDBServerNameValue).sValue
$OpsDBSQLAddress = $OpsDBServerAddress.split("\")[1]
$OpsDBSQLPort = $OpsDBServerAddress.split(",")[1]
$OpsDBServerName = $OpsDBServerAddress.split(",")[0].split("\")[0]
Write-Verbose "OpsDB Server Name: $OpsDBServerName"
$OpsDBSQLAddress = $OpsDBServerAddress.split("\")[1]
$OpsDBSQLPort = $OpsDBServerAddress.split(",")[1]
If ($OpsDBSQLAddress -ne $null)
{
	$OpsDBSQLInstance = $OpsDBSQLAddress.split(",")[0]
	Write-Verbose "OpsDB SQL Instance: $OpsDBSQLInstance"
} else {
	Write-Verbose "OpsDB SQL Instance: The Default Instance"
}
	
	
If ($OpsDBSQLPort -ne $null)
{
	Write-Verbose "OpsDB SQL Server Port: $OpsDBSQLPort"
} else {
	Write-Verbose "OpsDB SQL Server Port: Default or dynamic port"
}
#Get Total number of SDK connections
Write-Verbose "Getting total SDK Connection count."
$iTotalSDKCount = 0
Foreach ($MgmtServer in $arrMgmtServers)
{
	$MSFQDN = $MgmtServer.PrincipalName
	#$SDKPID = (Get-Process -ComputerName $MSFQDN -Name Microsoft.Mom.Sdk.ServiceHost).ID
    try {
	    #$Counter = Get-Counter "\\$MSFQDN`\OpsMgr SDK Service(microsoft.mom.sdk.servicehost-pid$SDKPID)\Client Connections"
		If ($Username -and $Password)
		{
			$SDKPID = (Get-WmiObject -ComputerName $MSFQDN -Query "Select * from win32_process where name = 'Microsoft.Mom.Sdk.ServiceHost.exe'" -Credential $PSCredential).ProcessId
			$Counter = Get-WmiObject -ComputerName $MSFQDN -Query "Select * from win32_PerfFormattedData_OpsMgrSDKService_OpsMgrSDKService Where Name = 'microsoft.mom.sdk.servicehost-pid$SDKPID'" -Credential $PSCredential
		} else {
			$SDKPID = (Get-WmiObject -ComputerName $MSFQDN -Query "Select * from win32_process where name = 'Microsoft.Mom.Sdk.ServiceHost.exe'").ProcessId
			$Counter = Get-WmiObject -ComputerName $MSFQDN -Query "Select * from win32_PerfFormattedData_OpsMgrSDKService_OpsMgrSDKService Where Name = 'microsoft.mom.sdk.servicehost-pid$SDKPID'"
		}
		$ConnectionValue = $Counter.ClientConnections
	    #Write-Verbose "$MSFQDN`: $ConnectionValue"
	    $iTotalSDKCount = $iTotalSDKCount + $ConnectionValue
    } catch {
        Write-Error "Unable to connect to $MSFQDN!"
    }
}
#Create a custom object to be used when generating HTML report
$objHTMLManagementGroupInfo = New-Object psobject
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Management Group Name" -Value $MGName
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Management Group Version" -Value $MGVersion
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Root Management Server Emulator" -Value $RMSE
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Operational DB SQL Server" -Value $OpsDBServerName
If ($OpsDBSQLInstance -ne $null)
{
	Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Operational DB SQL Instance" -Value $OpsDBSQLInstance
} else {
	Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Operational DB SQL Instance" -Value "MSSQLSERVER"
}
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Operational DB SQL Port" -Value $OpsDBSQLPort
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Operational DB Database Name" -Value $OpsDBName
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Data Warehouse SQL Server" -Value $DWServerName
If ($DWSQLInstance -ne $null)
{
	Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Data Warehouse SQL Instance" -Value $DWSQLInstance
} else {
	Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Data Warehouse SQL Instance" -Value "MSSQLSERVER"
}
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Data Warehouse SQL Port" -Value $DWSQLPort
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Data Warehouse SQL Database" -Value $DWDatabaseName
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Number of Management Servers" -Value $MgmtServersCount
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Number of Gateway Servers" -Value $GWServersCount
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Number of Windows Agents" -Value $iWindowsAgentCount
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Number of Unix/Linux Agents" -Value $iUnixComputerCount
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Number of Network Devices" -Value $iNetworkDevicesCount
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Number of Agentless Managed Computers" -Value $iAgentlessManagedCount
Add-Member -InputObject $objHTMLManagementGroupInfo -MemberType NoteProperty "Current SDK Connection Count" -Value $iTotalSDKCount

#Get SQL Server details
Write-Verbose "Get SQL Server version"
If ($Username -and $Password)
{
	$txtSQLVersion = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLVersionQuery -Username $Username -Password $Password)[0].Version
} else {
	$txtSQLVersion = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLVersionQuery)[0].Version
}

Write-Verbose "Get parameterised SQL Server properties"

If ($Username -and $Password)
{
	$objSQLServerProperty = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLPropertyQuery -Username $Username -Password $Password)[0]
} else {
	$objSQLServerProperty = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLPropertyQuery)[0]
}

Write-Verbose "Get SQL Server collation"
If ($Username -and $Password)
{
	$SQLServerCollation = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLCollationQuery -Username $Username -Password $Password)[0].ServerCollation
} else {
	$SQLServerCollation = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLCollationQuery)[0].ServerCollation
}

#Get SQL minimum and maximum memory
Write-Verbose "Getting SQL Server memory configuration"
If ($Username -and $Password)
{
	$SQLConfig = Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLConfigQuery -Username $Username -Password $Password
} else {
	$SQLConfig = Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLConfigQuery
}
$SQLMinMemory = ($SQLConfig | Where-Object {$_.name -ieq 'min server memory (MB)'}).value
$SQLMaxMemory = ($SQLConfig | Where-Object {$_.name -ieq 'max server memory (MB)'}).value
#Create a custom object to be used when generating HTML report
$objHTMLDWSQLServerInfo = New-Object psobject
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Server Manufacturer" -Value $DWServerManufacturer
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Server Model" -Value $DWServerModel
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Operating System Name" -Value $DWServerOSCaption
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Operating System Version" -Value $DWServerOSVersion
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Operating System Architecture" -Value $DWServerOSArchitecture
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Total Memory (MB)" -Value $DWServerTotalMemoryMB
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "CPU Speed (MHz)" -Value $DWServerCPUSpeed
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "CPU Address Width" -Value $DWServerCPUAddressWidth
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Number of CPU Cores" -Value $DWServerCPUCoreNumber
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Number of Logical Processors" -Value $DWServerCPULogicalProcessorNumber
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "SQL Version" -Value $txtSQLVersion
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Version Number" -Value $objSQLServerProperty.Version
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Edition" -Value $objSQLServerProperty.Edition
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Service Pack" -Value $objSQLServerProperty.ServicePack
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "SQL Server Collation" -Value $SQLServerCollation
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Minimum Memory (MB)" -Value $SQLMinMemory
Add-Member -InputObject $objHTMLDWSQLServerInfo -MemberType NoteProperty "Maximum Memory (MB)" -Value $SQLMaxMemory

#Check SQL service account rights
Write-Verbose "Check user rights for SQL service account (for performance optimisation)"
#Get SQL Service account
If ($Username -and $Password)
{
	$SQLServiceAccount = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLServiceAccountQuery -Username $Username -Password $Password)[0].service_account
} else {
	$SQLServiceAccount = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLServiceAccountQuery)[0].service_account
}
Write-Verbose "SQL Service account: `"$SQLServiceAccount`""
#Checking SQL Service Account rights on DW DB Server
If ($Username -and $Password)
{
	#Check if the SQL Service Account has "Perform volume maintenance tasks" rights
	Write-Verbose "Check if the SQL Service Account has `"Perform volume maintenance tasks`" rights using alternative credential (Username: '$Username')."
	$bHasVolMaintRights = Check-UserPrivilege -UserAccount $SQLServiceAccount -Privilege "SeManageVolumePrivilege" -ComputerName $DWServerName -Username $Username -Password $Password
	#Check if the SQL Service Account has "Lock Pages in Memory" rights
	Write-Verbose "Check if the SQL Service Account has `"Lock Pages in Memory`" rights using alternative credential (Username: '$Username')."
	$bHasLPIMRights = Check-UserPrivilege -UserAccount $SQLServiceAccount -Privilege "SeLockMemoryPrivilege" -ComputerName $DWServerName -Username $Username -Password $Password
} else {
	#Check if the SQL Service Account has "Perform volume maintenance tasks" rights
	Write-Verbose "Check if the SQL Service Account has `"Perform volume maintenance tasks`" rights."
	$bHasVolMaintRights = Check-UserPrivilege -UserAccount $SQLServiceAccount -Privilege "SeManageVolumePrivilege" -ComputerName $DWServerName
	#Check if the SQL Service Account has "Lock Pages in Memory" rights
	Write-Verbose "Check if the SQL Service Account has `"Lock Pages in Memory`" rights."
	$bHasLPIMRights = Check-UserPrivilege -UserAccount $SQLServiceAccount -Privilege "SeLockMemoryPrivilege" -ComputerName $DWServerName
}
#Create a custom object to be used when generating HTML report
$objHTMLSQLSerivceAccountConfig = New-Object psobject
Add-Member -InputObject $objHTMLSQLSerivceAccountConfig -MemberType NoteProperty "Service Account Name" -Value $SQLServiceAccount
Add-Member -InputObject $objHTMLSQLSerivceAccountConfig -MemberType NoteProperty "Perform volume maintenance tasks Rights" -Value $bHasVolMaintRights
Add-Member -InputObject $objHTMLSQLSerivceAccountConfig -MemberType NoteProperty "Lock Pages in Memory Rights" -Value $bHasLPIMRights

#Get TempDB Configuration
Write-Verbose "Checking tempdb configuration"
$TempDBGrowthQuery = [string]::format($DBGrowthQuery,"tempdb")
If ($Username -and $Password)
{
	$TempDBConfig = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $TempDBGrowthQuery -Username $Username -Password $Password)
} else {
	$TempDBConfig = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $TempDBGrowthQuery)
}
#Create an array list to store custom objects to be used when generating HTML report
$arrHTMLTempDBConfig = New-Object System.Collections.ArrayList
Foreach ($item in $TempDBConfig)
{
	$ojbHTMLTempDBConfig = New-Object psobject
	Add-Member -InputObject $ojbHTMLTempDBConfig -MemberType NoteProperty "Logical Name" -Value $item.LogicalName
	Add-Member -InputObject $ojbHTMLTempDBConfig -MemberType NoteProperty "File ID" -Value $item.FileID
	Add-Member -InputObject $ojbHTMLTempDBConfig -MemberType NoteProperty "File Name" -Value $item.FileName
	Add-Member -InputObject $ojbHTMLTempDBConfig -MemberType NoteProperty "File Group Name" -Value $item.FileGroupName
	Add-Member -InputObject $ojbHTMLTempDBConfig -MemberType NoteProperty "Size MB" -Value $item.SizeMB
	Add-Member -InputObject $ojbHTMLTempDBConfig -MemberType NoteProperty "Max Size MB" -Value $item.MaxSizeMB
	Add-Member -InputObject $ojbHTMLTempDBConfig -MemberType NoteProperty "Growth" -Value $item.Growth
	Add-Member -InputObject $ojbHTMLTempDBConfig -MemberType NoteProperty "Usage" -Value $item.Usage
	[void]$arrHTMLTempDBConfig.Add($ojbHTMLTempDBConfig)
}

#Get DW DB details
Write-Verbose "Get SCOM Datawarehouse DB properties"

If ($Username -and $Password)
{
	$objDWDBProperty = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $DWDBPropertyQuery -Username $Username -Password $Password)[0]
} else {
	$objDWDBProperty = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $DWDBPropertyQuery)[0]
}

#Get DW DB Growth Configuration
$DWDBGrowthQuery = [string]::format($DBGrowthQuery,$DWDatabaseName)
If ($Username -and $Password)
{
	$DWDBGrowth = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $DWDBGrowthQuery -Username $Username -Password $Password)
} else {
	$DWDBGrowth = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $DWDBGrowthQuery)
}
$DWDBDataGrowth = $DWDBGrowth | Where-Object {$_.Usage -ieq 'data only'}
$DWDBLogGrowth = $DWDBGrowth | Where-Object {$_.Usage -ieq 'log only'}

#Get SQL maintenance tasks details for DW DB
#Last backup
Write-Verbose "Get details of the last backup for SCOM Datawarehouse DB"
If ($Username -and $Password)
{
	$DWDBLastBckp = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database msdb -Query $DWDBLastBckpQuery -Username $Username -Password $Password)[0]
} else {
	$DWDBLastBckp = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database msdb -Query $DWDBLastBckpQuery)[0]
}
If ($DWDBLastBckp)
{
	$DWBackupSizeMB = $DWDBLastBckp.backup_size /1MB
	$DWBackupStartDate = $DWDBLastBckp.backup_start_date.ToString()
	$DWBackupFinishDate = $DWDBLastBckp.backup_finish_date.ToString()
	$DWBackupExpiryDate = $DWDBLastBckp.expiration_date.ToString()
} else {
	#DW DB has never been backed up
	$DWBackupStartDate = "None"
	$DWBackupFinishDate = "N/A"
	$DWBackupExpiryDate = "N/A"
	$DWBackupSizeMB = 0
}


#Get SQL free space
Write-Verbose "Getting OpsMgr Data Warehouse database free space"
If ($Username -and $Password)
{
	$SQLDWDBSize = Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $SQLDWDBFreeSpaceQuery -Username $Username -Password $Password
	$SQLDWLogSize = Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $SQLDWLogFreeSpaceQuery -Username $Username -Password $Password
} else {
	$SQLDWDBSize = Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $SQLDWDBFreeSpaceQuery
	$SQLDWLogSize = Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $SQLDWLogFreeSpaceQuery
}

$TotalDBSizeMB = 0
$TotalDBFreeSpaceMB = 0
Foreach ($item in $SQLDWDBSize)
{
	$TotalDBSizeMB = $TotalDBSizeMB + $item.CurrentSizeMB
	$TotalDBFreeSpaceMB = $TotalDBFreeSpaceMB + $item.FreeSpaceMB
}

$TotalLogSizeMB = 0
$TotalLogFreeSpaceMB = 0
Foreach ($item in $SQLDWLogSize)
{
	$TotalLogSizeMB = $TotalLogSizeMB + $item.CurrentSizeMB
	$TotalLogFreeSpaceMB = $TotalLogFreeSpaceMB + $item.FreeSpaceMB
}
$CultureInfo = Get-Culture
$CultureInfo.NumberFormat.PercentDecimalDigits = 2
$TotalDBFreeSpacePercent = ($TotalDBFreeSpaceMB/$TotalDBSizeMB).Tostring("P")
$TotalLogFreeSpacePercent = ($TotalLogFreeSpaceMB/$TotalLogSizeMB).Tostring("P")
#Create a custom object to be used when generating HTML report
$objHTMLDWDatabaseInfo = New-Object psobject
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Database Name" -Value $objDWDBProperty.name
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Database Creation Date" -Value $objDWDBProperty.create_date.ToString()
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Database Collation" -Value $objDWDBProperty.collation_name
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Database Recovery Mode" -Value $objDWDBProperty.recovery_model_desc
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Database State" -Value $objDWDBProperty.state_desc
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Broker Enabled" -Value $objDWDBProperty.is_broker_enabled.ToString()
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Auto Shrink Enabled" -Value $objDWDBProperty.is_auto_shrink_on.ToString()
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Data File Current Size MB" -Value $DWDBDataGrowth.SizeMB
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Data File Max Size MB" -Value $DWDBDataGrowth.MaxSizeMB
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Data Growth Setting" -Value $DWDBDataGrowth.Growth
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Data Free Space MB" -Value $TotalDBFreeSpaceMB.ToString()
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Data Free Space Percent" -Value $TotalDBFreeSpacePercent.ToString()
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Log File Current Size MB" -Value $DWDBLogGrowth.SizeMB
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Log File Max Size MB" -Value $DWDBLogGrowth.MaxSizeMB
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Log Growth Setting" -Value $DWDBLogGrowth.Growth
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Log Free Space MB" -Value $TotalLogFreeSpaceMB.ToString()
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Log Free Space Percent" -Value $TotalLogFreeSpacePercent.Tostring()
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Last Backup Start Date" -Value $DWBackupStartDate
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Last Backup Finish Date" -Value $DWBackupFinishDate
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Last Backup Expiry Date" -Value $DWBackupExpiryDate
Add-Member -InputObject $objHTMLDWDatabaseInfo -MemberType NoteProperty "Last Backup Size MB" -Value $DWBackupSizeMB.ToString()

#Get DW Dataset retention settings
Write-Verbose "Getting current OpsMgr Data Warehouse datasets retention settings"
If ($Username -and $Password)
{
	$DWDataSets = Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $SQLDWDataSetListQuery -Username $Username -Password $Password
} else {
	$DWDataSets = Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $SQLDWDataSetListQuery
}
#Since I cannot select from Stored Procedure results without creating a temp table, I'll filter and manipulate the result (datasets) in PowerShell now.
$arrDatasetDetails = New-Object System.Collections.ArrayList
foreach ($item in $DWDataSets)
{
	$objDatasetDetails = New-Object psobject
	$SizeMB = [int]($item.sizekb / [math]::pow(1024,1))
	$SizePercentOfTotalSize = [math]::Round($item.SizePercentOfTotalSize)
	Add-Member -InputObject $objDatasetDetails -MemberType NoteProperty -Name 'DatasetName' -Value $item.DatasetName
	Add-Member -InputObject $objDatasetDetails -MemberType NoteProperty -Name 'AggregationTypeName' -Value $item.AggregationTypeName
	Add-Member -InputObject $objDatasetDetails -MemberType NoteProperty -Name 'MaxDataAgeDays' -Value $item.MaxDataAgeDays
	Add-Member -InputObject $objDatasetDetails -MemberType NoteProperty -Name 'RowCount' -Value $item.RowCount
	Add-Member -InputObject $objDatasetDetails -MemberType NoteProperty -Name 'SizeKB' -Value $item.Sizekb
	Add-Member -InputObject $objDatasetDetails -MemberType NoteProperty -Name 'SizeMB' -Value $SizeMB
	Add-Member -InputObject $objDatasetDetails -MemberType NoteProperty -Name 'SizePercentOfTotalSize' -Value $SizePercentOfTotalSize
	[void]$arrDatasetDetails.Add($objDatasetDetails)
}
#Create an array list to store custom objects to be used when generating HTML report
$arrHTMLDatasetDetails = New-Object System.Collections.ArrayList
Foreach ($item in $arrDatasetDetails)
{
	#Create a custom object to be used when generating HTML report
	$objHTMLDatasetDetails = New-Object psobject
	Add-Member -InputObject $objHTMLDatasetDetails -MemberType NoteProperty "Dataset Name" -Value $item.DatasetName
	Add-Member -InputObject $objHTMLDatasetDetails -MemberType NoteProperty "Aggregation Name" -Value $item.AggregationTypeName
	Add-Member -InputObject $objHTMLDatasetDetails -MemberType NoteProperty "Max Age" -Value $item.MaxDataAgeDays
	Add-Member -InputObject $objHTMLDatasetDetails -MemberType NoteProperty "Row Count" -Value $item.RowCount
	Add-Member -InputObject $objHTMLDatasetDetails -MemberType NoteProperty "Current Size (KB)" -Value $item.SizeKB
	Add-Member -InputObject $objHTMLDatasetDetails -MemberType NoteProperty "Current Size (MB)" -Value $item.SizeMB
	Add-Member -InputObject $objHTMLDatasetDetails -MemberType NoteProperty "Percentage of Total Size" -Value $item.SizePercentOfTotalSize
	[Void]$arrHTMLDatasetDetails.Add($objHTMLDatasetDetails)
}

#Get DW aggregation Backlog
Write-Verbose "Getting OpsMgr Data Warehouse Datasets Aggregation Backlog"
If ($Username -and $Password)
{
	$outStandingAggr = Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $DWAggrSQLQuery -Username $Username -Password $Password
} else {
	$outStandingAggr = Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $DWAggrSQLQuery
}
#Create an array list to store custom objects to be used when generating HTML report
$arrHTMLAggrBacklogs = New-Object System.Collections.ArrayList
Foreach ($item in $outStandingAggr)
{
	#Create a custom object to be used when generating HTML report
	$objHTMLAggrBacklogs = New-Object psobject
	Add-Member -InputObject $objHTMLAggrBacklogs -MemberType NoteProperty "Aggregation Type" -Value $item.AggregationType
	Add-Member -InputObject $objHTMLAggrBacklogs -MemberType NoteProperty "Count" -Value $item.Count
	Add-Member -InputObject $objHTMLAggrBacklogs -MemberType NoteProperty "Dataset Id" -Value $item.Datasetid
	Add-Member -InputObject $objHTMLAggrBacklogs -MemberType NoteProperty "Dataset Name" -Value $item.DatasetName
	Add-Member -InputObject $objHTMLAggrBacklogs -MemberType NoteProperty "First Record Date" -Value $item.First
	Add-Member -InputObject $objHTMLAggrBacklogs -MemberType NoteProperty "Last Record Date" -Value $item.Last
	[Void]$arrHTMLAggrBacklogs.Add($objHTMLAggrBacklogs)
}

#Get DW staging tables row counts
Write-Verbose "Getting OpsMgr Data Warehouse staging tables row counts"
If ($Username -and $Password)
{
	$AlertStagingCount = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $AlertStagingCountSQL -Username $Username -Password $Password)[0].count
	$EventStagingCount = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $EventStagingCountSQL -Username $Username -Password $Password)[0].count
	$PerfStagingCount = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $PerfStagingCountSQL -Username $Username -Password $Password)[0].count
	$StateStagingCount = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $StateStagingCountSQL -Username $Username -Password $Password)[0].count
	$ManagedEntityStagingCount = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $ManagedEntityStagingCountSQL -Username $Username -Password $Password)[0].count
} else {
	$AlertStagingCount = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $AlertStagingCountSQL)[0].count
	$EventStagingCount = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $EventStagingCountSQL)[0].count
	$PerfStagingCount = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $PerfStagingCountSQL)[0].count
	$StateStagingCount = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $StateStagingCountSQL)[0].count
	$ManagedEntityStagingCount = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database $DWDatabaseName -Query $ManagedEntityStagingCountSQL)[0].count
}
#Create a custom object to be used when generating HTML report
$objHTMLSQLDWStagingTableRowCount = New-Object psobject
Add-Member -InputObject $objHTMLSQLDWStagingTableRowCount -MemberType NoteProperty "Alert.AlertStage Table Row Count" -Value $AlertStagingCount
Add-Member -InputObject $objHTMLSQLDWStagingTableRowCount -MemberType NoteProperty "Event.EventStage Table Row Count" -Value $EventStagingCount
Add-Member -InputObject $objHTMLSQLDWStagingTableRowCount -MemberType NoteProperty "Perf.PerformanceStage Table Row Count" -Value $PerfStagingCount
Add-Member -InputObject $objHTMLSQLDWStagingTableRowCount -MemberType NoteProperty "State.StateStage Table Row Count" -Value $StateStagingCount
Add-Member -InputObject $objHTMLSQLDWStagingTableRowCount -MemberType NoteProperty "ManagedEntityStage Table Row Count" -Value $ManagedEntityStagingCount

#Get SQL perf counters
Write-Verbose "Getting SQL performance counters"
If ($Username -and $Password)
{
	$SQLPerfBCHR = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLPerfBCHRQuery -Username $Username -Password $Password)
	$SQLPerfPLE = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLPerfPLEQuery -Username $Username -Password $Password)[0]
} else {
	$SQLPerfBCHR = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLPerfBCHRQuery)
	$SQLPerfPLE = (Invoke-SQLQuery -SQLServer $DWServerName -SQLInstance $DWSQLInstance -SQLPort $DWSQLPort -Database master -Query $SQLPerfPLEQuery)[0]
}
#Calculate SQL perf counters
#Buffer Cache Hit Ratio
Write-Verbose "Calclating Buffer Cache Hit Ratio"
$BCHR = ($SQLPerfBCHR | Where-Object {$_.counter_name.trimend() -imatch "^Buffer cache hit ratio$"}).cntr_value
$BCHRB =($SQLPerfBCHR | Where-Object {$_.counter_name.trimend() -imatch "^Buffer cache hit ratio base$"}).cntr_value
Write-Verbose "Buffer Cache Hit Ratio: $BCHR"
Write-Verbose "Buffer Cache Hit Ratio Base: $BCHRB"
$SQLBCHRValue = 100 * $BCHR/$BCHRB
$SQLBCHRValue = "$SQLBCHRValue`%"
Write-Verbose "Buffer Hit Ratio %: $SQLBCHRValue"
#Create an array list to store custom objects to be used when generating HTML report
$arrHTMLSQLPerfCounters = New-Object System.Collections.ArrayList
$objHTMLSQLPerfBCHR = New-Object psobject
Add-Member -InputObject $objHTMLSQLPerfBCHR -MemberType NoteProperty "Counter Display Name" -Value "SQLServer:Buffer Manager\Buffer cache hit ratio"
Add-Member -InputObject $objHTMLSQLPerfBCHR -MemberType NoteProperty "Object Name" -Value "SQLServer:Buffer Manager"
Add-Member -InputObject $objHTMLSQLPerfBCHR -MemberType NoteProperty "Counter Name" -Value "Buffer cache hit ratio"
Add-Member -InputObject $objHTMLSQLPerfBCHR -MemberType NoteProperty "Value" -Value $SQLBCHRValue
[Void]$arrHTMLSQLPerfCounters.Add($objHTMLSQLPerfBCHR)
$objHTMLSQLPerfPLE = New-Object psobject
Add-Member -InputObject $objHTMLSQLPerfPLE -MemberType NoteProperty "Counter Display Name" -Value "SQLServer:Buffer Manager\Page Life Expectancy"
Add-Member -InputObject $objHTMLSQLPerfPLE -MemberType NoteProperty "Object Name" -Value "SQLServer:Buffer Manager"
Add-Member -InputObject $objHTMLSQLPerfPLE -MemberType NoteProperty "Counter Name" -Value "Page life expectancy"
Add-Member -InputObject $objHTMLSQLPerfPLE -MemberType NoteProperty "Value" -Value $SQLPerfPLE.cntr_value
[Void]$arrHTMLSQLPerfCounters.Add($objHTMLSQLPerfPLE)

#Get OS Perf counters
If ($username -and $password)
{
	#\LogicalDisk(_total)\Avg. Disk sec/Read
	$LDAvgDiskSecRead = (Get-WmiObject -ComputerName $DWServerName -Query "Select * from win32_PerfFormattedData_PerfDisk_LogicalDisk Where Name = '_Total'" -Credential $PSCredential).AvgDisksecPerRead
	#\LogicalDisk(_total)\Avg. Disk sec/Write
	$LDAvgDiskSecWrite = (Get-WmiObject -ComputerName $DWServerName -Query "Select * from win32_PerfFormattedData_PerfDisk_LogicalDisk Where Name = '_Total'" -Credential $PSCredential).AvgDisksecPerWrite
	#\Processor Information(_Total)\% Processor Time
	$ProcInfoProcessorTime = (Get-WmiObject -ComputerName $DWServerName -Query "Select * from win32_PerfFormattedData_PerfOS_Processor Where Name = '_Total'" -Credential $PSCredential).PercentProcessorTime
} else {
	#\LogicalDisk(_total)\Avg. Disk sec/Read
	$LDAvgDiskSecRead = (Get-WmiObject -ComputerName $DWServerName -Query "Select * from win32_PerfFormattedData_PerfDisk_LogicalDisk Where Name = '_Total'").AvgDisksecPerRead
	#\LogicalDisk(_total)\Avg. Disk sec/Write
	$LDAvgDiskSecWrite = (Get-WmiObject -ComputerName $DWServerName -Query "Select * from win32_PerfFormattedData_PerfDisk_LogicalDisk Where Name = '_Total'").AvgDisksecPerWrite
	#\Processor Information(_Total)\% Processor Time
	$ProcInfoProcessorTime = (Get-WmiObject -ComputerName $DWServerName -Query "Select * from win32_PerfFormattedData_PerfOS_Processor Where Name = '_Total'").PercentProcessorTime
}
Write-Verbose "Getting Operating System performance counters"
#Create an array list to store custom objects to be used when generating HTML report
$arrHTMLOSPerfCounters = New-Object System.Collections.ArrayList
$objHTMLLDAvgDiskSecRead = New-Object psobject
Add-Member -InputObject $objHTMLLDAvgDiskSecRead -MemberType NoteProperty "Counter Display Name" -Value "LogicalDisk(_total)\Avg. disk sec/Read"
Add-Member -InputObject $objHTMLLDAvgDiskSecRead -MemberType NoteProperty "Object Name" -Value "LogicalDisk(_total)"
Add-Member -InputObject $objHTMLLDAvgDiskSecRead -MemberType NoteProperty "Counter Name" -Value "Avg. disk sec/Read"
Add-Member -InputObject $objHTMLLDAvgDiskSecRead -MemberType NoteProperty "Value" -Value $LDAvgDiskSecRead
[Void]$arrHTMLOSPerfCounters.Add($objHTMLLDAvgDiskSecRead)
$objHTMLLDAvgDiskSecWrite = New-Object psobject
Add-Member -InputObject $objHTMLLDAvgDiskSecWrite -MemberType NoteProperty "Counter Display Name" -Value "LogicalDisk(_total)\Avg. disk sec/Write"
Add-Member -InputObject $objHTMLLDAvgDiskSecWrite -MemberType NoteProperty "Object Name" -Value "LogicalDisk(_total)"
Add-Member -InputObject $objHTMLLDAvgDiskSecWrite -MemberType NoteProperty "Counter Name" -Value "Avg. disk sec/Write"
Add-Member -InputObject $objHTMLLDAvgDiskSecWrite -MemberType NoteProperty "Value" -Value $LDAvgDiskSecWrite
[Void]$arrHTMLOSPerfCounters.Add($objHTMLLDAvgDiskSecWrite)
$objHTMLProcInfoProcessorTime = New-Object psobject
Add-Member -InputObject $objHTMLProcInfoProcessorTime -MemberType NoteProperty "Counter Display Name" -Value "Processor Information(_Total)\% Processor Time"
Add-Member -InputObject $objHTMLProcInfoProcessorTime -MemberType NoteProperty "Object Name" -Value "Processor Information(_Total)"
Add-Member -InputObject $objHTMLProcInfoProcessorTime -MemberType NoteProperty "Counter Name" -Value "% Processor Time"
Add-Member -InputObject $objHTMLProcInfoProcessorTime -MemberType NoteProperty "Value" -Value $ProcInfoProcessorTime
[Void]$arrHTMLOSPerfCounters.Add($objHTMLProcInfoProcessorTime)

#Get relative events from "Operations Manager" log from each management server
Write-Verbose "Getting related event log entries in the 'Operations Manager' log from each management server (Event ID 2115, 8000 and 31550-31559)."
$EventsXMLFilter = '<QueryList><Query Id="0" Path="Operations Manager"><Select Path="Operations Manager">*[System[(EventID=2115 or EventID=8000 or  (EventID &gt;= 31550 and EventID &lt;= 31559) )]]</Select></Query></QueryList>'
$strMainHTMLReportEventSection = @"
<h2>Data Warehouse Related Events From Management Servers</h2>
<p>Event ID: 2115, 8000 and 31550-31559</p>
<table> <colgroup><col/><col/><col/><col/></colgroup>
<tr><th>Management Server</th><th>Number of Events Captured</th><th>Events HTML Report</th><th>Event CSV Export</th> 
"@
Foreach ($Mgmtserver in $arrMgmtServers)
{
	$MgmtServerName = $MgmtServer.Name
	$MgmtServerComputerName = $MgmtServer.ComputerName
	Write-Verbose "Exporting events from $MgmtServerName."
	
	If ($Username -and $Password)
	{
		$Events = Get-WinEvent -Computername $MgmtServerName -FilterXml $EventsXMLFilter -Credential $PSCredential
	} else {
		$Events = Get-WinEvent -Computername $MgmtServerName -FilterXml $EventsXMLFilter
	}
	$iEventCount = $Events.count
	Write-Verbose "Found $iEventCount events from $MgmtServerName."
	#Export to HTML and CSV
	$EventsOutputFileName = "DWRelatedEvents-$MgmtServerComputerName-$now"
	$outputXML = Join-Path $OutputDir "$OutputFileName.xml"
	$EventsOutputHTML = Join-Path $OutputDir "$EventsOutputFileName.html"
	$EventsOutputCSV = Join-Path $OutputDir "$EventsOutputFileName.csv"
	$EventsHTMLHeader = @()
	$EventsHTMLHeader += "<style>"
	$EventsHTMLHeader += "BODY{font-family:Calibri,sans-serif; font-size: small; font color: black}"
	$EventsHTMLHeader += "H1{text-align:center;color:#102F47}"
	$EventsHTMLHeader += "H2{color:#102F47}"
	$EventsHTMLHeader += "TABLE{border-width: 1px;text-align: left;border-style: solid;border-color: black;border-collapse: collapse;}"
	$EventsHTMLHeader += "TH{border-width: 1px;padding: 4px;border-style: solid;border-color: black;background-color:#102F47;color:white}"
	$EventsHTMLHeader += "TD{border-width: 1px;padding: 4px;border-style: solid;border-color: black;}"
	$EventsHTMLHeader += "</style>"

	$EventsHTMLBody = @()
	$EventsHTMLBody +="<h1>OpsMgr 2012 Data Warehouse Related Events</h1>"
	$EventsHTMLBody +="<p align='center'>Report Date:  $Now</p>"
	$EventsHTMLBody +="<p align='center'>Event IDs: 2115, 8000 and 31550-31559</p>"
	$EventsHTMLBody += "<h2>$MgmtServerName</h2>"
	$EventsHTMLBody += "<table> <colgroup><col/><col/><col/><col/><col/></colgroup>"
	$EventsHTMLBody += "<tr><th>Time Created</th><th>Event ID</th><th>Event Level</th><th>Machine Name</th><th>Message</th></tr>"
	Foreach ($event in $events)
	{
		Switch ($event.LevelDisplayName)
		{
			"Warning" {$fontColour = "#FF9900"}
			"Error" {$fontColour = "Red"}
			default {$fontColour = "Black"}
		}
		$EventTimeCreated = $event.TimeCreated
		$EventId = $event.Id
		$EventLevelDisplayName = $event.LevelDisplayName
		$EventMachineName = $event.MachineName
		$EventMessage = $event.Message
		$EventsHTMLBody += "<tr><td>$EventTimeCreated</td><td>$EventId</td><td><font color='$fontColour'>$EventLevelDisplayName</font></td><td>$EventMachineName</td><td>$EventMessage</td></tr>"
		#$EventsHTMLBody += $events | select TimeCreated, Id, LevelDisplayName, MachineName, Message | ConvertTo-Html -Fragment
	}
	$EventsHTMLBody += "</table>"
	#$EventsHTMLBody += $events | select TimeCreated, Id, LevelDisplayName, MachineName, Message | ConvertTo-Html -Fragment
	ConvertTo-HTML -head $EventsHTMLHeader -body "$EventsHTMLBody" | Out-File $EventsOutputHTML
	[void]$arrOutputFiles.Add($EventsOutputHTML)
	$events | select TimeCreated, Id, LevelDisplayName, MachineName, Message | Export-Csv -Path $EventsOutputCSV -NoTypeInformation
	[void]$arrOutputFiles.Add($EventsOutputCSV)
	$strMainHTMLReportEventSection = $strMainHTMLReportEventSection  + "<tr><td>$MgmtServerName</td><td>$iEventCount</td><td><a href='$EventsOutputFileName.html'>$EventsOutputFileName.html</a></td><td><a href='$EventsOutputFileName.csv'>$EventsOutputFileName.csv</a></td></tr>"
}
$strMainHTMLReportEventSection = $strMainHTMLReportEventSection + "</table>"
#endregion

#region Generate Report
Write-Verbose "Loading Result template XML"
$xml = [xml](Get-Content $resultTemplateXML)

Write-Verbose "Saving Management Group Information"
$xml.Data.ManagementGroupInfo.MGName = $MGName
$xml.Data.ManagementGroupInfo.MGVersion = $MGVersion
$xml.Data.ManagementGroupInfo.RMSE = $RMSE
$xml.Data.ManagementGroupInfo.OpsDBServerName = $OpsDBServerName
If ($OpsDBSQLInstance -ne $NULL)
{
	$xml.Data.ManagementGroupInfo.OpsDBSQLInstance = $OpsDBSQLInstance
} else {
	$xml.Data.ManagementGroupInfo.OpsDBSQLInstance = "MSSQLSERVER"
}
If ($OpsDBSQLPort -ne $NULL)
{
	$xml.Data.ManagementGroupInfo.OpsDBSQLPort = $OpsDBSQLPort.Tostring()
}
$xml.Data.ManagementGroupInfo.OpsDBName = $OpsDBName
$xml.Data.ManagementGroupInfo.DWServerName = $DWServerName
If ($DWSQLInstance -ne $NULL)
{
	$xml.Data.ManagementGroupInfo.DWSQLInstance = $DWSQLInstance
} else {
	$xml.Data.ManagementGroupInfo.DWSQLInstance = "MSSQLSERVER"
}
If ($DWSQLPort -ne $null)
{
	$xml.Data.ManagementGroupInfo.DWSQLPort = $DWSQLPort.Tostring()
}
$xml.Data.ManagementGroupInfo.DWDatabaseName = $DWDatabaseName
$xml.Data.ManagementGroupInfo.MgmtServerCount = $MgmtServersCount.ToString()
$xml.Data.ManagementGroupInfo.GWServerCount = $GWServersCount.ToString()
$xml.Data.ManagementGroupInfo.WindowsAgentCount = $iWindowsAgentCount.ToString()
$xml.Data.ManagementGroupInfo.UnixAgentCount = $iUnixComputerCount.ToString()
$xml.Data.ManagementGroupInfo.NetworkDeviceCount = $iNetworkDevicesCount.ToString()
$xml.Data.ManagementGroupInfo.AgenlessComputerCount = $iAgentlessManagedCount.ToString()
$xml.Data.ManagementGroupInfo.CurrentSDKConnections= $iTotalSDKCount.Tostring()

Write-Verbose "Saving Data Warehouse SQL Server Information"
$xml.Data.DWSQLServerInfo.ServerHardwareManufacturer = $DWServerManufacturer
$xml.Data.DWSQLServerInfo.ServerHardwareModel = $DWServerModel
$xml.Data.DWSQLServerInfo.OSCaption = $DWServerOSCaption
$xml.Data.DWSQLServerInfo.OSVersion = $DWServerOSVersion
$xml.Data.DWSQLServerInfo.OSArchitecture = $DWServerOSArchitecture
$xml.Data.DWSQLServerInfo.ServerTotalMemoryMB = $DWServerTotalMemoryMB.ToString()
$xml.Data.DWSQLServerInfo.CPUSpeed = $DWServerCPUSpeed.ToString()
$xml.Data.DWSQLServerInfo.CPUAddressWidth = $DWServerCPUAddressWidth.ToString()
$xml.Data.DWSQLServerInfo.CPUCoreNumber = $DWServerCPUCoreNumber.ToString()
$xml.Data.DWSQLServerInfo.CPULogicalProcessorNumber = $DWServerCPULogicalProcessorNumber.ToString()
$xml.Data.DWSQLServerInfo.SQLVersion = $txtSQLVersion
$xml.Data.DWSQLServerInfo.SQLVersionNumber = $objSQLServerProperty.Version.Tostring()
$xml.Data.DWSQLServerInfo.SQLEdition = $objSQLServerProperty.Edition
$xml.Data.DWSQLServerInfo.SQLSP = $objSQLServerProperty.ServicePack
$xml.Data.DWSQLServerInfo.SQLCollation = $SQLServerCollation
$xml.Data.DWSQLServerInfo.SQLMinMemory = $SQLMinMemory.Tostring()
$xml.Data.DWSQLServerInfo.SQLMaxMemory = $SQLMaxMemory.Tostring()

Write-Verbose 'Saving TempDB Configuration'
$DWSQLTempDBNode = $xml.Data.SelectNodes("DWSQLTempDB")[0]
Foreach ($item in $TempDBConfig)
{
	$NewTempDBFile = $DWSQLTempDBNode.AppendChild($xml.CreateElement("TempDBFile"))
	Foreach ($Property in (Get-Member -InputObject $item -MemberType NoteProperty))
	{
		$PropertyName = $Property.Name
        [void]$NewTempDBFile.AppendChild($xml.CreateElement($PropertyName))
		$NewTempDBFile.$PropertyName = $item.$PropertyName.Tostring()
	}
}
Write-Verbose "Saving SQL Service Account Configuration"
$xml.Data.SQLSerivceAccountConfig.ServiceAccountName = $SQLServiceAccount
$xml.Data.SQLSerivceAccountConfig.PerformVolumeMaintenanceTasksPrivilege = $bHasVolMaintRights.ToString()
$xml.Data.SQLSerivceAccountConfig.LockPagesInMemoryPrivilege = $bHasLPIMRights.Tostring()

Write-Verbose "Saving Data Warehouse Database Information"
$xml.Data.DWDatabaseInfo.DBName = $objDWDBProperty.name
$xml.Data.DWDatabaseInfo.DBCreationDate = $objDWDBProperty.create_date.ToString()
$xml.Data.DWDatabaseInfo.DBCollation = $objDWDBProperty.collation_name
$xml.Data.DWDatabaseInfo.DBRecoveryMode = $objDWDBProperty.recovery_model_desc
$xml.Data.DWDatabaseInfo.DBState = $objDWDBProperty.state_desc
$xml.Data.DWDatabaseInfo.DBBrokerEnabled = $objDWDBProperty.is_broker_enabled.ToString()
$xml.Data.DWDatabaseInfo.DBAutoShrink = $objDWDBProperty.is_auto_shrink_on.ToString()
$xml.Data.DWDatabaseInfo.DBDataFileCurrentSizeMB = $DWDBDataGrowth.SizeMB.ToString()
$xml.Data.DWDatabaseInfo.DBDataFileMaxSizeMB = $DWDBDataGrowth.MaxSizeMB.ToString()
$xml.Data.DWDatabaseInfo.DBDataGrowthSetting = $DWDBDataGrowth.Growth.ToString()
$xml.Data.DWDatabaseInfo.DBDataFreeSpaceMB = $TotalDBFreeSpaceMB.ToString()
$xml.Data.DWDatabaseInfo.DBDataFreeSpacePercent = $TotalDBFreeSpacePercent.ToString()
$xml.Data.DWDatabaseInfo.DBLogFileCurrentSizeMB = $DWDBLogGrowth.SizeMB.ToString()
$xml.Data.DWDatabaseInfo.DBLogFileMaxSizeMB = $DWDBLogGrowth.MaxSizeMB.ToString()
$xml.Data.DWDatabaseInfo.DBLogGrowthSetting = $DWDBLogGrowth.Growth.ToString()
$xml.Data.DWDatabaseInfo.DBLogFreeSpaceMB = $TotalLogFreeSpaceMB.ToString()
$xml.Data.DWDatabaseInfo.DBLogFreeSpacePercent = $TotalLogFreeSpacePercent.Tostring()
$xml.Data.DWDatabaseInfo.DBLastBackupStartDate = $DWBackupStartDate
$xml.Data.DWDatabaseInfo.DBLastBackupFinishDate = $DWBackupFinishDate
$xml.Data.DWDatabaseInfo.DBLastBackupExpiryDate = $DWBackupExpiryDate
$xml.Data.DWDatabaseInfo.DBLastBackupSizeMB = $DWBackupSizeMB.ToString()

Write-Verbose "Saving Data Warehouse datasets configuration"
$DatasetDetailsNode = $xml.Data.SelectNodes("DatasetDetails")[0]
Foreach ($Dataset in $arrDatasetDetails)
{
	$NewDatasetDetail = $DatasetDetailsNode.AppendChild($xml.CreateElement("DatasetDetail"))
	Foreach ($Property in (Get-Member -InputObject $Dataset -MemberType NoteProperty))
	{
		$PropertyName = $Property.Name
        [void]$NewDatasetDetail.AppendChild($xml.CreateElement($PropertyName))
		$NewDatasetDetail.$PropertyName = $Dataset.$PropertyName.Tostring()
	}
}

Write-Verbose "Saving Data Warehouse dataset aggregation backlogs information"
$DatasetAggregationBacklogsNode = $xml.Data.SelectNodes("DatasetAggregationBacklogs")[0]
Foreach ($aggr in $outStandingAggr)
{
	$NewAggr = $DatasetAggregationBacklogsNode.AppendChild($xml.CreateElement("DatasetAggregationBacklog"))
	Foreach ($Property in (Get-Member -InputObject $aggr -MemberType NoteProperty))
	{
		$PropertyName = $Property.Name
        [void]$NewAggr.AppendChild($xml.CreateElement($PropertyName))
		$NewAggr.$PropertyName = $aggr.$PropertyName.Tostring()
	}
}
Write-Verbose "Saving Data Warehouse Database Staging Tables Row Count"
$xml.Data.StagingTables.AlertStagingRowCount = $AlertStagingCount.ToString()
$xml.Data.StagingTables.EventStagingRowCount = $EventStagingCount.ToString()
$xml.Data.StagingTables.PerfStagingRowCount = $PerfStagingCount.ToString()
$xml.Data.StagingTables.StateStagingRowCount = $StateStagingCount.ToString()
$xml.Data.StagingTables.ManagedEntityStagingRowCount = $ManagedEntityStagingCount.ToString()

Write-Verbose "Saving SQL performance counters"
$SQLPerfCountersNode = $xml.Data.SelectNodes("SQLPerfCounters")[0]
#SQLServer:Buffer Manager\Buffer cache hit ratio
$SQLBCHRCounter = $SQLPerfCountersNode.AppendChild($xml.CreateElement("SQLPerfCounter"))
[Void]$SQLBCHRCounter.AppendChild($xml.CreateElement("CounterDisplayName"))
[Void]$SQLBCHRCounter.AppendChild($xml.CreateElement("ObjectName"))
[Void]$SQLBCHRCounter.AppendChild($xml.CreateElement("CounterName"))
[Void]$SQLBCHRCounter.AppendChild($xml.CreateElement("Value"))
$SQLBCHRCounter.CounterDisplayName = "SQLServer:Buffer Manager\Buffer cache hit ratio"
$SQLBCHRCounter.ObjectName = "SQLServer:Buffer Manager"
$SQLBCHRCounter.CounterName = "Buffer cache hit ratio"
$SQLBCHRCounter.Value = $SQLBCHRValue
#SQLServer:Buffer Manager\Page life expectancy
$SQLPLECounter = $SQLPerfCountersNode.AppendChild($xml.CreateElement("SQLPerfCounter"))
[Void]$SQLPLECounter.AppendChild($xml.CreateElement("CounterDisplayName"))
[Void]$SQLPLECounter.AppendChild($xml.CreateElement("ObjectName"))
[Void]$SQLPLECounter.AppendChild($xml.CreateElement("CounterName"))
[Void]$SQLPLECounter.AppendChild($xml.CreateElement("Value"))
$SQLPLECounter.CounterDisplayName = "SQLServer:Buffer Manager\Page Life Expectancy"
$SQLPLECounter.ObjectName = $SQLPerfPLE.object_name.trim()
$SQLPLECounter.CounterName = $SQLPerfPLE.counter_name.trim()
$SQLPLECounter.Value = $SQLPerfPLE.cntr_value.ToString()

Write-Verbose "Saving Windows OS performance counters"
$OSPerfCountersNode = $xml.Data.SelectNodes("OSPerfCounters")[0]
#\LogicalDisk(_total)\Avg. Disk sec/Read
$OSLDAvgDiskSecReadCounter = $OSPerfCountersNode.AppendChild($xml.CreateElement("OSPerfCounter"))
[Void]$OSLDAvgDiskSecReadCounter.AppendChild($xml.CreateElement("CounterDisplayName"))
[Void]$OSLDAvgDiskSecReadCounter.AppendChild($xml.CreateElement("ObjectName"))
[Void]$OSLDAvgDiskSecReadCounter.AppendChild($xml.CreateElement("CounterName"))
[Void]$OSLDAvgDiskSecReadCounter.AppendChild($xml.CreateElement("Value"))
$OSLDAvgDiskSecReadCounter.CounterDisplayName = "LogicalDisk(_total)\Avg. disk sec/Read"
$OSLDAvgDiskSecReadCounter.ObjectName = "LogicalDisk(_total)"
$OSLDAvgDiskSecReadCounter.CounterName = "Avg. disk sec/Read"
$OSLDAvgDiskSecReadCounter.Value = $LDAvgDiskSecRead.ToString()
#\LogicalDisk(_total)\Avg. Disk sec/Write
$OSLDAvgDiskSecWriteCounter = $OSPerfCountersNode.AppendChild($xml.CreateElement("OSPerfCounter"))
[Void]$OSLDAvgDiskSecWriteCounter.AppendChild($xml.CreateElement("CounterDisplayName"))
[Void]$OSLDAvgDiskSecWriteCounter.AppendChild($xml.CreateElement("ObjectName"))
[Void]$OSLDAvgDiskSecWriteCounter.AppendChild($xml.CreateElement("CounterName"))
[Void]$OSLDAvgDiskSecWriteCounter.AppendChild($xml.CreateElement("Value"))
$OSLDAvgDiskSecWriteCounter.CounterDisplayName = "LogicalDisk(_total)\Avg. disk sec/Write"
$OSLDAvgDiskSecWriteCounter.ObjectName = "LogicalDisk(_total)"
$OSLDAvgDiskSecWriteCounter.CounterName = "Avg. disk sec/Write"
$OSLDAvgDiskSecWriteCounter.Value = $LDAvgDiskSecWrite.ToString()
#\Processor Information(_Total)\% Processor Time
$OSProcInfoProcessorTimeCounter = $OSPerfCountersNode.AppendChild($xml.CreateElement("OSPerfCounter"))
[Void]$OSProcInfoProcessorTimeCounter.AppendChild($xml.CreateElement("CounterDisplayName"))
[Void]$OSProcInfoProcessorTimeCounter.AppendChild($xml.CreateElement("ObjectName"))
[Void]$OSProcInfoProcessorTimeCounter.AppendChild($xml.CreateElement("CounterName"))
[Void]$OSProcInfoProcessorTimeCounter.AppendChild($xml.CreateElement("Value"))
$OSProcInfoProcessorTimeCounter.CounterDisplayName = "Processor Information(_Total)\% Processor Time"
$OSProcInfoProcessorTimeCounter.ObjectName = "Processor Information(_Total)"
$OSProcInfoProcessorTimeCounter.CounterName = "% Processor Time"
$OSProcInfoProcessorTimeCounter.Value = $ProcInfoProcessorTime.ToString()

#Save XML
$xml.save($outputXML)
[void]$arrOutputFiles.Add($outputXML)

#Save Report to HTML
#HTML Header
$HTMLHeader = @()
$HTMLHeader += "<style>"
$HTMLHeader += "BODY{font-family:Calibri,sans-serif; font-size: small;}"
$HTMLHeader += "H1{text-align:center;color:#102F47}"
$HTMLHeader += "H2{color:#102F47}"
$HTMLHeader += "H3{color:#102F47}"
$HTMLHeader += "TABLE{border-width: 1px;text-align: center;border-style: solid;border-color: black;border-collapse: collapse;}"
$HTMLHeader += "TH{border-width: 1px;padding: 4px;border-style: solid;border-color: black;background-color:#102F47;color:white}"
$HTMLHeader += "TD{border-width: 1px;padding: 4px;border-style: solid;border-color: black;}"
$HTMLHeader += "</style>"

#HTML Body
$HTMLBody = @()
$HTMLBody +="<h1>OpsMgr 2012 Data Warehouse Health Check Report - $MGName</h1>"
$HTMLBody +="<p align='center'>Report Date:  $Now</p>"
$HTMLBody += "<hr>"
#MG Info
$HTMLBody += "<h2>Management Group Information</h2>"
$HTMLBody += $objHTMLManagementGroupInfo | select * | ConvertTo-Html -As List -Fragment

#DW SQL Info
$HTMLBody += "<h2>Data Warehouse SQL Server Information</h2>"
$HTMLBody += $objHTMLDWSQLServerInfo | select * | ConvertTo-Html -As List -Fragment
$HTMLBody += "<h3>Note:</h3>"
$HTMLBody += "<p>Please compare your Operations Manager configuration with Microsoft's supported configuration:</P>"
$HTMLBody += "<p><a href='https://technet.microsoft.com/en-au/library/jj656649.aspx'>System Requirements for System Center Operations Manager 2012 (7.0.8560.0)</a></P>"
$HTMLBody += "<p><a href='https://technet.microsoft.com/en-au/library/jj656654.aspx'>System Requirements for System Center Operations Manager 2012 SP1 (Version 7.1.9538.0)</a></P>"
$HTMLBody += "<p><a href='https://technet.microsoft.com/en-au/library/dn249696.aspx'>System Requirements for System Center Operations Manager 2012 R2 (Version 7.1.10226.0)</a></P>"
$HTMLBody += "<hr>"

#DW DB Info
$HTMLBody += "<h2>Data Warehouse SQL Database Information</h2>"
$HTMLBody += $objHTMLDWDatabaseInfo | select * | ConvertTo-Html -As List -Fragment
$HTMLBody += "<h3>Recommendations:</h3>"
$HTMLBody += "<p>It is recommended the OpsMgr Data Warehouse DB is configured as the following:</P>"
$HTMLBody += "<ul>"
$HTMLBody += "<li>Recovery Mode: leave it as default (SIMPLE)</li>"
$HTMLBody += "<li>Make sure it is been backed up on a schedule</li>"
$HTMLBody += "<li>Other than CHECKDB, do not configure any other maintenance tasks against this database</li>"
$HTMLBody += "<li>Use the <a href='http://blogs.technet.com/b/momteam/archive/2012/04/02/operations-manager-2012-sizing-helper-tool.aspx'>OpsMgr 2012 Sizing Helper Tool</a> to estimate the space required for the Data Warehouse Database.</li>"
$HTMLBody += "<li>If Auto-Growth is enabled on the Data Warehouse database, please make sure the auto increase amount is configured to increase a reasonable amount. This will ensure the auto grow will not take place too frequently, and help reduce the data base file fragmentation.</li>"
$HTMLBody += "<li>Make sure Auto-Shrink is disabled.</li>"
$HTMLBody += "</ul>"
$HTMLBody += "<p>Please refer to Kevin Holman's article <a href='http://blogs.technet.com/b/kevinholman/archive/2008/04/12/what-sql-maintenance-should-i-perform-on-my-opsmgr-databases.aspx'>What SQL maintenance should I perform on my OpsMgr databases?</a> for more information.</P>"
$HTMLBody += "<hr>"

#DW SQL TempDB Config
$HTMLBody += "<h2>Data Warehouse SQL Server TempDB Configuration</h2>"
$HTMLBody += $arrHTMLTempDBConfig | select * | ConvertTo-Html -Fragment
$HTMLBody += "<h3>Recommendations:</h3>"
$HTMLBody += "<p>Please create separate equally sized files for the tempdb and size them appropriately. For better performance, please place them on faster disks.</P>"
$HTMLBody += "<p>Please refer to Page 19 of the <a href='https://gallery.technet.microsoft.com/SQL-Server-guide-for-8584c403'>SQL Server guide for System Center 2012</a> for detailed tempdb configuration recommendation.</P>"
$HTMLBody += "<hr>"

#DW SQL Service Account Config
$HTMLBody += "<h2>Data Warehouse SQL Service Account Configuration</h2>"
$HTMLBody += $objHTMLSQLSerivceAccountConfig | select * | ConvertTo-Html -Fragment
$HTMLBody += "<h3>Recommendations:</h3>"
$HTMLBody += "<p>It is recommended that the SQL Service Account has 'Performance Volume maintenance tasks' and 'Lock Pages in Memory' privilege on the SQL server.</P>"
$HTMLBody += "<p>Please refer to Page 16 of the <a href='https://gallery.technet.microsoft.com/SQL-Server-guide-for-8584c403'>SQL Server guide for System Center 2012</a> for details.</P>"
$HTMLBody += "<h3>Note:</h3>"
$HTMLBody += "<p>This health check script only checks direct privilege assignments. It <b>does not</b> check if the service account is a member of each groups that have been assigned to the privileges listed above. If this report indicates the service account does not have the privilege, you may manually check if the service account is a member of any groups listed in the local policy(gpedit.msc), under 'Computer Configuration\Windows Settings\Security Settings\Local Policies\User Rights Assignments'</P>"
$HTMLBody += "<hr>"

#Dataset Details
$HTMLBody += "<h2>Data Warehouse Data Sets Information</h2>"
$HTMLBody += $arrHTMLDatasetDetails | select * | ConvertTo-Html -Fragment
$HTMLBody += "<h3>Recommendations:</h3>"
$HTMLBody += "<ul>"
$HTMLBody += "<li>Use <a href='http://blogs.technet.com/b/momteam/archive/2008/05/14/data-warehouse-data-retention-policy-dwdatarp-exe.aspx'>dwdatarp.exe</a> to configure data retention according to your business needs. reduce the retention peroids from default if the business does not require longer retention periods (for reporting).</li>"
$HTMLBody += "<li>Use overrides to disable the performance and event collection rules that you do not wish to run reports against.</li>"
$HTMLBody += "<li>Wherever you can, use overrides to increase the interval on how often the performance and event collection rules. This would help to reduce the amount of raw data stored in the Data Warehouse DB.</li>"
$HTMLBody += "</ul>"
$HTMLBody += "<h3>Additional Information:</h3>"
$HTMLBody += "<ul>"
$HTMLBody += "<li><a href='http://blogs.technet.com/b/kevinholman/archive/2010/01/05/understanding-and-modifying-data-warehouse-retention-and-grooming.aspx'>Understanding and modifying Data Warehouse retention and grooming</a> By Kevin Holman</li>"
$HTMLBody += "<li><a href='http://blogs.technet.com/b/stefan_stranger/archive/2009/08/15/everything-you-wanted-to-know-about-opsmgr-data-warehouse-grooming-but-were-afraid-to-ask.aspx'>Everything you wanted to know about OpsMgr Data Warehouse Grooming but were afraid to ask</a> By Stefan Stranger</li>"
$HTMLBody += "</ul>"
$HTMLBody += "<hr>"

#Dataset Aggregation Backlogs
$HTMLBody += "<h2>Data Warehouse Data Sets Aggregation Backlog</h2>"
$HTMLBody += $arrHTMLAggrBacklogs | select * | ConvertTo-Html -Fragment
$HTMLBody += "<h3>Additional Information:</h3>"
$HTMLBody += "<ul>"
$HTMLBody += "<li><a href='https://michelkamp.wordpress.com/2012/04/10/scom-dwh-aggregations-data-loose-tip-and-tricks/'>SCOM DWH aggregations data loose Tip and Tricks</a> By Michel Kamp</li>"
$HTMLBody += "<li><a href='https://michelkamp.wordpress.com/2013/03/24/get-a-grip-on-the-dwh-aggregations/'>Get a grip on the DWH aggregations</a> By Michel Kamp</li>"
$HTMLBody += "<li><a href='https://michelkamp.wordpress.com/2012/03/23/dude-where-my-availability-report-data-from-the-scom-dwh/'>Dude where is my Availability Report data from the SCOM DWH ??</a> By Michel Kamp</li>"
$HTMLBody += "<li><a href='http://ok-sandbox.com/2014/04/scom-standard-dataset-maintenance-workflow/'>SCOM Standard Dataset maintenance workflow</a> By Oleg Kapustin</li>"
$HTMLBody += "</ul>"
$HTMLBody += "<hr>"

#DW Staging Tables Row Count
$HTMLBody += "<h2>Data Warehouse Database Staging Tables Row Count</h2>"
$HTMLBody += $objHTMLSQLDWStagingTableRowCount | select * | ConvertTo-Html -As List -Fragment
$HTMLBody += "<h3>Additional Information:</h3>"
$HTMLBody += "<ul>"
$HTMLBody += "<li><a href='http://www.bictt.com/blogs/bictt.php/2014/10/10/case-of-the-fast-growing'>Case of the fast growing SCOM datawarehouse db and logs</a> By Bob Cornelissen</li>"
$HTMLBody += "</ul>"
$HTMLBody += "<hr>"

#SQL Perf Counters
$HTMLBody += "<h2>Key SQL Performance Counters</h2>"
$HTMLBody += $arrHTMLSQLPerfCounters | select * | ConvertTo-Html -Fragment

#OS Perf Counters
$HTMLBody += "<h2>Key Operating System Performance Counters</h2>"
$HTMLBody += $arrHTMLOSPerfCounters | select * | ConvertTo-Html -Fragment

$HTMLBody += "<h3>Additional Information:</h3>"
$HTMLBody += "<p>Please refer to Page 32 of the <a href='https://gallery.technet.microsoft.com/SQL-Server-guide-for-8584c403'>SQL Server guide for System Center 2012</a> for detailed explaination of the key SQL and OS performance counters.</P>"
$HTMLBody += "<hr>"

#Event Log export
$HTMLBody += $strMainHTMLReportEventSection
$HTMLBody += "<h3>Additional Troubleshooting Resources:</h3>"
$HTMLBody += "<ul>"
$HTMLBody += "<li>Event ID 2115: <a href='http://blogs.technet.com/b/kevinholman/archive/2008/04/21/event-id-2115-a-bind-data-source-in-management-group.aspx'>Event ID 2115 A Bind Data Source in Management Group</a></li>"
$HTMLBody += "<li>Event ID 2115, 31552: <a href='https://support.microsoft.com/en-au/kb/2573329'>Troubleshooting Blank Reports in System Center Operations Manager</a></li>"
$HTMLBody += "<li>Event ID 31552: <a href='http://www.systemcentercentral.com/fix-failed-to-store-data-in-the-data-warehouse-due-to-a-exception-sqlexception-timeout-expired/'>FIX: Failed to store data in the Data Warehouse due to a Exception ‘SqlException': Timeout expired.</a></li>"
$HTMLBody += "<li>Event ID 31552: <a href='http://blogs.technet.com/b/kevinholman/archive/2010/08/30/the-31552-event-or-why-is-my-data-warehouse-server-consuming-so-much-cpu.aspx'>The 31552 event, or why is my data warehouse server consuming so much CPU?</a></li>"
$HTMLBody += "<li>Event ID 31552: <a href='https://michelkamp.wordpress.com/2012/01/05/howto-failed-to-store-data-in-the-data-warehouse-arithmetic-overflow-error-converting-expression-to-data-type-float/'>[HOWTO] Failed to store data in the Data Warehouse : Arithmetic overflow error converting expression to data type float.</a></li>"
$HTMLBody += "<li>Event ID 31553: <a href='http://blogs.technet.com/b/operationsmgr/archive/2011/09/06/standard-dataset-maintenance-troubleshooter-for-system-center-operations-manager-2007.aspx'>Standard Dataset Maintenance troubleshooter for System Center Operations Manager 2007</a></li>"
$HTMLBody += "</ul>"

#Save HTML
ConvertTo-HTML -head $HTMLHeader -body "$HTMLBody" | Out-File $outputHTML
[void]$arrOutputFiles.Add($outputHTML)
Write-Output "Done."
Write-Output "The following files are generated by this script:"
Foreach ($file in $arrOutputFiles)
{
	Write-Output "  - $file"
}
If ($OpenReport)
{#Open report
	Write-Verbose "Opening HTML Report"
	$OpenReportCmd = "$env:SystemRoot\System32\rundll32.exe url.dll,FileProtocolHandler $outputHTML"
	Invoke-Expression $OpenReportCmd
}
#endregion