<#
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

function Get-GroupsToSync {
    param (
        $DomainControllerFQDN, 
        $EmailDomain,
        [switch]
        $ConvertExistingDistributionGroupsToSecurityGroups
    )
    # Get DistributionGroups from Office 365
    $Groups365 = Get-DistributionGroup

    # Get DistributionGroups from Active Directory
    $filter = "objectClass -eq `"group`" -and mail -like `"*$($EmailDomain)`""
    $GroupsAD = Get-ADObject -Filter $filter -Properties * -Server $DomainControllerFQDN

    # Add ExistsInOffice365 Property to GroupsAD object
    $GroupsAD | ForEach-Object {
        $_ | Add-Member -NotePropertyName "ExistsInOffice365" -NotePropertyValue $false -Force
    }

    # Add SyncComplete property to objects and set to $false
    $GroupsAD | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name SyncComplete -Value $false -force -ErrorAction SilentlyContinue
    }

    # Determine Groups to Sync
    $GroupsToSync = @()
    foreach ($ADGroup in $GroupsAD) {
        $365Group = $Groups365 | Where-Object PrimarySmtpAddress -Like $ADGroup.Mail
        
        # If no match on primary mail properties, add to GroupsToSync Array
        if (-NOT $365Group) {
            $GroupsToSync += $ADGroup
            Continue
        }
        elseif ($ConvertExistingDistributionGroupsToSecurityGroups) {
            if (-NOT (($365Group).RecipientTypeDetails -match 'Security')) {
                try {
                    Remove-DistributionGroup -Identity $365Group.Name -BypassSecurityGroupManagerCheck -Confirm:$false
                    $EventMessage = "SyncActiveDirectoryToOffice365 `nREMOVED Group: $($365Group.Name)`n"
                    $EventMessage += "This was done becase -ConvertExistingDistributionGroupsToSecurityGroups was specified and the Group was not a Security Group."
                    Write-ScriptEvent -EntryType Warning -EventId 187 -Message $EventMessage
                    $GroupsToSync += $ADGroup
                    Continue
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nAttempted to delete Group: $($365Group.Name) `nError was: $ErrorMessage"
                    $ADGroup.ExistsInOffice365 = $true # Group still exists if it fails to delete
                    $GroupsToSync += $ADGroup
                    Continue
                }
            }
        }
        
        $ADGroup.ExistsInOffice365 = $true
        # Verify properties
        if ($365Group.Name -notlike $ADGroup.Name) {
            $GroupsToSync += $ADGroup
            Continue
        }
        if ($365Group.DisplayName -notlike $ADGroup.DisplayName) {
            $GroupsToSync += $ADGroup
            Continue
        }
        if ($365Group.WindowsEmailAddress -notlike $ADGroup.mail) {
            $GroupsToSync += $ADGroup
            Continue
        }
        if (Compare-Object -ReferenceObject $ADGroup.ProxyAddresses -DifferenceObject $365Group.EmailAddresses) {
            $GroupsToSync += $ADGroup
            Continue
        }
        # Verify Group Membership
        $365GroupMembers = Get-DistributionGroupMember -Identity $ADGroup.Name -ErrorAction SilentlyContinue
        $ADGroupMembers = $ADGroup.member | Get-ADObject -Properties mail -ErrorAction SilentlyContinue
        $MembershipChecks = @()
        foreach ($365member in $365GroupMembers) {
            if ( ($365member.PrimarySmtpAddress ) -and ($365member.PrimarySmtpAddress -notin $ADGroupMembers.mail) ) {
                $GroupsToSync += $ADGroup
                $MembershipChecks += $false
            }
        }
        foreach ($ADmember in $ADGroupMembers) {
            if ( ($ADmember.mail) -and ($ADmember.mail -notin $365GroupMembers.PrimarySmtpAddress)) {
                $GroupsToSync += $ADGroup
                $MembershipChecks += $false
            }
        }
        if ($MembershipChecks -contains $false) {
            Continue
        }
        else {
            # If all properties match
            $ADGroup.SyncComplete = $true
            $GroupsToSync += $ADGroup
        }

    }

    # Return Objects
    $GroupsToSync = $GroupsToSync | Select-Object -Unique
    $GroupsToSync
}
