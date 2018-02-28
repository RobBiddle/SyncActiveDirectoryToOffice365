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
function Sync-GroupsToOffice365 {
    param (
        $GroupsToSync, 
        [switch]
        $CreateNewGroupsAsSecurityGroups
    )
    foreach ($ADGroup in $GroupsToSync) {

        # If Group is not in Office 365, create Group
        if ($ADGroup.ExistsInOffice365 -eq $false) {
            try {
                $NewGroupParams = @{
                    Name               = $ADGroup.Name
                    DisplayName        = $ADGroup.DisplayName
                    PrimarySmtpAddress = $ADGroup.mail
                }
                if ($CreateNewGroupsAsSecurityGroups) {
                    $NewGroupParams += @{Type = "Security"}
                }
                
                $EventMessage = "SyncActiveDirectoryToOffice365 `nAdding DistributionGroup: $($ADGroup.Name)`n"
                $EventMessage += "Object ADGroup.ExistsInOffice365 = $($ADGroup.ExistsInOffice365)`n"
                $EventMessage += "Object ADGroup.SyncComplete = $($ADGroup.SyncComplete)`n"
                Write-ScriptEvent -EntryType Information -EventId 411 -Message $EventMessage
                New-DistributionGroup @NewGroupParams
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened during New-DistributionGroup  `nError was: $ErrorMessage"
            }

            $365Group = $null
            try {
                $365Group = Get-DistributionGroup $ADGroup.Name
            }
            catch {
                $error.clear()
            }

            if ($365Group) {
                $ADGroup.ExistsInOffice365 = $true    
            }

        }
        # If Group exists in Office 365, sync properties
        if ( ($ADGroup.ExistsInOffice365 -eq $true) -and ($ADGroup.SyncComplete -eq $false) ) {
            # If Group is in Office 365, set attributes
            $365Group = $null
            $365Group = Get-DistributionGroup -Identity $ADGroup.Name -ErrorAction SilentlyContinue
            $365GroupMembers = Get-DistributionGroupMember -Identity $ADGroup.Name -ErrorAction SilentlyContinue
            $ADGroupMembers = $ADGroup.member | Get-ADObject -Properties mail -ErrorAction SilentlyContinue
            # Sync Name
            if ($365Group.Name -notlike $ADGroup.Name) {
                try {
                    Set-DistributionGroup -Identity $365Group.Identity `
                        -Name $ADGroup.Name;
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened while setting group name  `nError was: $ErrorMessage"
                }

            }
            if ($365Group.DisplayName -notlike $ADGroup.DisplayName) {
                try {
                    Set-DistributionGroup -Identity $365Group.Identity `
                        -DisplayName $ADGroup.DisplayName;
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened while setting group displayname  `nError was: $ErrorMessage"
                }

            }
            # Sync proxyAddresses
            foreach ($365ProxyAddress in $365Group.EmailAddresses) {
                if ( ($365ProxyAddress) -and ($365ProxyAddress -notin $ADGroup.proxyAddresses)) {
                    try {
                        Set-DistributionGroup -Identity $365Group.Identity -EmailAddresses @{Remove = $365ProxyAddress}
                    }
                    catch {
                        $ErrorMessage = $_.Exception.Message
                        Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened :-(  `nError was: $ErrorMessage"
                    }

                }
            }
            foreach ($ADproxyAddress in $ADGroup.proxyAddresses) {
                if ( ($ADproxyAddress) -and ($ADproxyAddress -notin $365Group.EmailAddresses)) {
                    try {
                        Set-DistributionGroup -Identity $365Group.Identity -EmailAddresses @{Add = $ADproxyAddress}
                        Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365 `nAdded proxy address: $ADproxyAddress `nTo Group: $($365Group.Identity)"
                    }
                    catch {
                        $ErrorMessage = $_.Exception.Message
                        Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nAttempted to add an email address $ADproxyAddress `nto Group: $($365Group.Identity) `nError was: $ErrorMessage"
                    }
                }
            }

            # Sync Group Members
            foreach ($365member in $365GroupMembers) {
                if ( ($365member.PrimarySmtpAddress ) -and ($365member.PrimarySmtpAddress -notin $ADGroupMembers.mail) ) {
                    try {
                        Remove-DistributionGroupMember -Identity $365Group.Identity -Member $365member.PrimarySmtpAddress
                        Write-ScriptEvent -EntryType Warning -EventId 411 -Message "SyncActiveDirectoryToOffice365 Removed user: $($365member.PrimarySmtpAddress) `nfrom Office 365 Group: $($365Group.Identity)"
                    }
                    catch {
                        $ErrorMessage = $_.Exception.Message
                        Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened while removing user: $($365member.PrimarySmtpAddress) `nfrom Office 365 Group: $($365Group.Identity) `nError was: $ErrorMessage"
                    }
                    
                }
            }

            foreach ($ADmember in $ADGroupMembers) {
                if ( ($ADmember.mail) -and ($ADmember.mail -notin $365GroupMembers.PrimarySmtpAddress)) {
                    try {
                        Add-DistributionGroupMember -Identity $365Group.Identity -Member $ADmember.mail
                        Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365 `nAdded user: $($ADmember.mail) `nTo Office 365 Group: $($365Group.Identity)"
                    }
                    catch {
                        $ErrorMessage = $_.Exception.Message
                        Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened while adding user: $($ADmember.mail) `nTo Office 365 Group: $($365Group.Identity)  `nError was: $ErrorMessage"
                    }
                        
                }
            }

        }
    
    }
    
}
