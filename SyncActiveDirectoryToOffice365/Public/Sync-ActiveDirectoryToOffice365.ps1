<#
.SYNOPSIS
    Sync Active Directory Contact, Group & User objects to Office 365
.DESCRIPTION
    Sync Active Directory Contact, Group & User objects to Office 365
    Syncs basic object properties, as well as Group membership
    WARNING: Assumes Active Directory is the source of truth
.EXAMPLE
    PS C:\> Sync-ActiveDirectoryToOffice365 -BaseOU "OU=MyOU,DC=MyDomain,DC=com" `
        -DomainControllerFQDN "MyDC.MyDomain.com" `
        -EmailDomain "MyDomain.com" `
        -ObjectsToSync Contacts, Groups, Users `
        -CredentialForOffice365 (Get-Credential) `
        -CreateNewGroupsAsSecurityGroups `
        -ConvertExistingDistributionGroupsToSecurityGroups `
        -EnableModernAuth;
.NOTES
    Author: Robert D. Biddle
    https://github.com/RobBiddle
    https://github.com/RobBiddle/SyncActiveDirectoryToOffice365
    SyncActiveDirectoryToOffice365  Copyright (C) 2017  Robert D. Biddle
    This program comes with ABSOLUTELY NO WARRANTY; for details type `"help Sync-ActiveDirectoryToOffice365 -full`".
    This is free software, and you are welcome to redistribute it
    under certain conditions; for details type `"help Sync-ActiveDirectoryToOffice365 -full`".
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
    The GNU General Public License does not permit incorporating your program
    into proprietary programs.  If your program is a subroutine library, you
    may consider it more useful to permit linking proprietary applications with
    the library.  If this is what you want to do, use the GNU Lesser General
    Public License instead of this License.  But first, please read
    <http://www.gnu.org/philosophy/why-not-lgpl.html>.
#>

function Sync-ActiveDirectoryToOffice365 {
    [CmdletBinding(DefaultParametersetName = "Set 1")]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # Active Directory Domain Controller FQDN
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $false, ParameterSetName = "Set 1")]
        [Parameter(HelpMessage = "FQDN of Active Directory Domain Controller")]
        [String]
        $DomainControllerFQDN,
    
        # Credential for Office365
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $false, ParameterSetName = "Set 1")]
        [Parameter(HelpMessage = "PSCredential object for Office 365")]
        [PSCredential]
        $CredentialForOffice365,
    
        # Sync AD Objects in Organizational Unit
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $false, ParameterSetName = "Set 1")]
        [Parameter(HelpMessage = "Sync AD Objects in Organizational Unit")]
        [String]
        $BaseOU,
    
        # Sync AD Objects associated with specified Email Domain
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $false, ParameterSetName = "Set 1")]
        [Parameter(HelpMessage = "Sync AD Objects associated with specified Email Domain")]
        [String]
        $EmailDomain,
    
        # Mail Enabled Objects to Sync
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $false, ParameterSetName = "Set 1")]
        [Parameter(HelpMessage = "Mail Enabled Objects to Sync")]
        [ValidateSet('Contacts', 'Groups', 'Users')]
        [string[]]
        $ObjectsToSync,

        # Create New Office 365 DistributionGroups As Mail Enabled Security Type Groups
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $false, ParameterSetName = "Set 1")]
        [Parameter(HelpMessage = "Create New Office 365 DistributionGroups As Mail Enabled Security Type Groups")]
        [switch]
        $CreateNewGroupsAsSecurityGroups,

        # Convert Existing Office 365 DistributionGroups To Mail Enabled Security Groups
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $false, ParameterSetName = "Set 1")]
        [Parameter(HelpMessage = "Convert Existing Office 365 DistributionGroups To Mail Enabled SecurityGroups")]
        [switch]
        $ConvertExistingDistributionGroupsToSecurityGroups,

        # Enable Modern Authentication (OAuth2)
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $false, ParameterSetName = "Set 1")]
        [Parameter(HelpMessage = "Enable Modern Authentication i.e. OAuth2")]
        [switch]
        $EnableModernAuth
    )
    # Elevated Console is required to create new EventLog Source
    If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
                [Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
        Break
    }

    # This is required to catch errors from imported cmdlets :-/
    $Global:ErrorActionPreference = 'Stop'
    
    # Trap Block to catch anything outside of a try/catch
    trap {
        $ErrorMessage = $_.Exception.Message
        Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened :-(  `nError was: $ErrorMessage"
        continue
    }

    # Force ObjectsToSync to be an Array, even if only one of the validation set options is entered
    $ObjectsToSync = @($ObjectsToSync)

    # Set Base OU if not specified
    if (-NOT $BaseOU) {
        $SplitDomainControllerFQDN = $DomainControllerFQDN -split '\.'
        $BaseOU = ""
        1..($SplitDomainControllerFQDN.count -1) | ForEach-Object {
            if ($_ -eq 1) {
                $BaseOU = "DC=$($SplitDomainControllerFQDN[$_])"
            }
            else {
                $BaseOU = "$BaseOU,DC=$($SplitDomainControllerFQDN[$_])"
            }
        }
    }

    # Connect to Office365
    Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365`n Connecting to Msol-Service for $EmailDomain"
    Connect-MsolService -Credential $CredentialForOffice365

    # Connect to Exchange Online
    Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365`n Connecting to Exchange Online for $EmailDomain"
    $EOSession = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
        -Credential $CredentialForOffice365 `
        -Authentication Basic `
        -AllowRedirection
    Import-PSSession $EOSession -AllowClobber -Verbose:$false -ErrorAction Stop -WarningAction SilentlyContinue

    # Enable Office365 Modern Authentication
    if ($EnableModernAuth) {
        if (-NOT ((Get-OrganizationConfig).OAuth2ClientProfileEnabled)) {
            Set-OrganizationConfig -OAuth2ClientProfileEnabled $true
        }
    }
    
    # Check for ActiveDirectory Module
    if (-NOT (Get-Module ActiveDirectory -ErrorAction SilentlyContinue)) {
        # Try to import ActiveDirectory
        Import-Module ActiveDirectory -Force -Verbose:$false -ErrorAction Stop -WarningAction SilentlyContinue
    }
    
    if ($ObjectsToSync -contains 'Users') {
        Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365`n Getting Users for $EmailDomain"
        $UsersToSync = Get-UsersToSync -DomainControllerFQDN $DomainControllerFQDN -EmailDomain $EmailDomain
        Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365`n Syncing Users for $EmailDomain"
        Sync-UsersToOffice365 -UsersToSync $UsersToSync
    }
    
    if ($ObjectsToSync -contains 'Contacts') {
        Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365`n Getting Contacts for $EmailDomain"
        $ContactsToSync = Get-ContactsToSync -BaseOU $BaseOU -DomainControllerFQDN $DomainControllerFQDN -EmailDomain $EmailDomain
        Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365`n Syncing Contacts for $EmailDomain"
        Sync-ContactsToOffice365 -ContactsToSync $ContactsToSync
    }

    if ($ObjectsToSync -contains 'Groups') {
        $GetGroupParams = @{
            DomainControllerFQDN = $DomainControllerFQDN
            EmailDomain          = $EmailDomain
        }
        if ($ConvertExistingDistributionGroupsToSecurityGroups) {
            $GetGroupParams += @{ConvertExistingDistributionGroupsToSecurityGroups = $true}
        }
        Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365`n Getting Groups for $EmailDomain"
        $GroupsToSync = Get-GroupsToSync @GetGroupParams

        $SyncGroupParams = @{
            GroupsToSync = $GroupsToSync
        }
        if ($CreateNewGroupsAsSecurityGroups) {
            $SyncGroupParams += @{CreateNewGroupsAsSecurityGroups = $true}
        }
        Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365`n Syncing Groups for $EmailDomain"
        Sync-GroupsToOffice365 @SyncGroupParams
    }

    # Clean up PSSession
    Get-PSSession | Remove-PSSession
    
}
