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

function Sync-UsersToOffice365 ($UsersToSync) {
    # Trap Block to catch anything outside of a try/catch
    trap {
        $ErrorMessage = $_.Exception.Message
        Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened :-(  `nError was: $ErrorMessage"
        continue
    }
    Add-Type -AssemblyName System.Web # Provides support for generating random passwords
    $StringForEventMessage = $null
    foreach ($ADUser in $UsersToSync) {

        # If User is not in Office 365, create User
        if ($ADUser.ExistsInOffice365 -eq $false) {
            try {
                Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365 `nCreating new Office 365 user: $($ADUser.UserPrincipalName)"
                $password = [System.Web.Security.Membership]::GeneratePassword(16, 0)
                New-MsolUser -TenantId $currentTenantId `
                    -UserPrincipalName "$($ADUser.UserPrincipalName)" `
                    -DisplayName "$($ADUser.DisplayName)" `
                    -FirstName "$($ADUser.GivenName)" `
                    -LastName "$($ADUser.Surname)" `
                    -ImmutableId $ADUser.immutableID `
                    -Password $password ;
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened during New-MsolUser :-(  `nError was: $ErrorMessage"
            }

            $365User = $null
            try {
                $365User = Get-MsolUser -UserPrincipalName $ADUser.UserPrincipalName
            }
            catch {
                Continue
            }

            if ($365User) {
                $ADUser.ExistsInOffice365 = $true
            }
        }

        # Sync properties
        if ( ($ADUser.ExistsInOffice365 -eq $true) -and ($ADUser.SyncComplete -eq $false)  ) {
            # If User is in Office 365, set attributes
            $365User = $null
            $365User = Get-MsolUser -UserPrincipalName $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
            $365Mailbox = Get-Mailbox -Identity $365User.UserPrincipalName -ErrorAction SilentlyContinue

            if ($365User.BlockCredential -eq $ADUser.Enabled) {
                if (($365User.BlockCredential -eq $true) -and ($ADUser.Enabled -eq $true )   ) {
                    $EventMessage = "Enabling User: $($365User.DisplayName)"
                    Set-MsolUser -UserPrincipalName $365User.userPrincipalName `
                        -BlockCredential $false;
                }
                elseif (($365User.BlockCredential -eq $false) -and ($ADUser.Enabled -eq $false )   ) {
                    $EventMessage = "Disabling User: $($365User.DisplayName)"
                    Set-MsolUser -UserPrincipalName $365User.userPrincipalName `
                        -BlockCredential $true;
                }
                Write-ScriptEvent -EntryType Warning -EventId 411 -Message $EventMessage
            }

            if ($365User.FirstName -notlike $ADUser.givenName) {
                try {
                    Set-MsolUser -ObjectId $365User.ObjectId `
                        -UserPrincipalName $365User.userPrincipalName `
                        -FirstName $ADUser.givenName;
                    $EventMessage = "Updated the FirstName for 365User: $($365User.DisplayName) / ADUser: $($ADUser.UserPrincipalName)"
                    Write-ScriptEvent -EntryType Information -EventId 411 -Message $EventMessage
                }
                catch {
                    $EventMessage = "Something Unexpected happened to SyncActiveDirectoryToOffice365 while attempting to update the FirstName for user: $($365User.DisplayName)  Error was: $ErrorMessage"
                    $ErrorMessage = $_.Exception.Message
                    Write-ScriptEvent -EntryType Error -EventId 187 -Message $EventMessage
                }

            }

            if ($365User.LastName -notlike $ADUser.sn) {
                try {
                    Set-MsolUser -ObjectId $365User.ObjectId `
                        -UserPrincipalName $365User.userPrincipalName `
                        -LastName $ADUser.sn;
                    $EventMessage = "Updated the LastName for 365User: $($365User.DisplayName) / ADUser: $($ADUser.UserPrincipalName)"
                    Write-ScriptEvent -EntryType Information -EventId 411 -Message $EventMessage
                }
                catch {
                    $EventMessage = "Something Unexpected happened to SyncActiveDirectoryToOffice365 while attempting to update the LastName for user: $($365User.DisplayName)  Error was: $ErrorMessage"
                    $ErrorMessage = $_.Exception.Message
                    Write-ScriptEvent -EntryType Error -EventId 187 -Message $EventMessage
                }

            }

            if ($365User.DisplayName -notlike $ADUser.DisplayName) {
                try {
                    Set-MsolUser -ObjectId $365User.ObjectId `
                        -UserPrincipalName $365User.userPrincipalName `
                        -DisplayName $ADUser.DisplayName;
                    $EventMessage = "Updated the DisplayName for 365User: $($365User.DisplayName) / ADUser: $($ADUser.UserPrincipalName)"
                    Write-ScriptEvent -EntryType Information -EventId 411 -Message $EventMessage
                }
                catch {
                    $EventMessage = "Something Unexpected happened to SyncActiveDirectoryToOffice365 while attempting to update the DisplayName for user: $($365User.DisplayName)  Error was: $ErrorMessage"
                    $ErrorMessage = $_.Exception.Message
                    Write-ScriptEvent -EntryType Error -EventId 187 -Message $EventMessage
                }

            }

            if ($365User.immutableID -notlike $ADUser.immutableID) {
                # Setting the ImmutableID is only possible if the Tenant is not currently using Federated Authentication
                if ( (Get-MsolDomain -DomainName "$(($ADUser.UserPrincipalName -split '@')[1])").Authentication -eq 'Federated') {
                    $ErrorMessage = "SyncActiveDirectoryToOffice365 `nImmutableId for 365User: $($365User.DisplayName) / ADUser: $($ADUser.UserPrincipalName) DOES NOT MATCH!`n"
                    $ErrorMessage += "Domain: $(($ADUser.UserPrincipalName -split '@')[1]) is currently FEDERATED`n"
                    $ErrorMessage += "The Domain must be set back to MANAGED in order to update this ImmutableId."
                    Write-ScriptEvent -EntryType Error -EventId 911 -Message $ErrorMessage
                }
                else {
                    try {
                        Set-MsolUser -ObjectId $365User.ObjectId `
                            -UserPrincipalName $ADUser.UserPrincipalName `
                            -ImmutableId $ADUser.immutableID;
                        $EventMessage = "Updated the ImmutableId for User: $($365User.DisplayName)"
                        Write-ScriptEvent -EntryType Information -EventId 411 -Message $EventMessage
                    }
                    catch {
                        $EventMessage = "Something Unexpected happened to SyncActiveDirectoryToOffice365 while attempting to update the immutableId for user: $($365User.DisplayName)  Error was: $ErrorMessage"
                        $ErrorMessage = $_.Exception.Message
                        Write-ScriptEvent -EntryType Error -EventId 187 -Message $EventMessage
                    }

                }
            }

            # Sync proxyAddresses - $365User.proxyAddresses is empty prior to mailbox creation
            if ($365User.proxyAddresses.count -gt 0) {
                $365ProxyAddressesToRemove = @()
                foreach ($365ProxyAddress in $365Mailbox.EmailAddresses) {
                    if ( ($365ProxyAddress -imatch 'SMTP') -and ($365ProxyAddress -notin $ADUser.proxyAddresses)) {
                        try {
                            Set-MailBox -Identity $365User.userPrincipalName -EmailAddresses @{Remove = $365ProxyAddress} -Confirm:$false
                            $365ProxyAddressesToRemove += $365ProxyAddress
                        }
                        catch {
                            $ErrorMessage = $_.Exception.Message
                            Write-ScriptEvent -EntryType Error -EventId 187 -Message "Something Unexpected happened to SyncActiveDirectoryToOffice365 while attempting to remove proxy address: $ADProxyAddress from user: $($365User.DisplayName)  Error was: $ErrorMessage"
                        }
                    }
                }
                if ($365ProxyAddressesToRemove.count -gt 0) {
                    foreach ($proxyaddress in $365ProxyAddressesToRemove) {
                        $StringForEventMessage += $proxyaddress
                    }
                    $StringForEventMessage += " From User: $($365User.DisplayName) `n"
                }

                # Convert Existing LegacyExchangeDN values to X500 Addresses and add to proxyAddresses
                if ($ADUser.LegacyExchangeDN) {
                    $X500LegacyExchangeDN = Convert-LegacyExchangeDNToX500 -Address $ADUser.LegacyExchangeDN
                    if ($X500LegacyExchangeDN -notin $ADUser.proxyAddresses ) {
                        $ADUser.proxyAddresses += $X500LegacyExchangeDN
                    }
                }

                foreach ($ADproxyAddress in $ADUser.proxyAddresses) {
                    if ($ADproxyAddress -notin $365Mailbox.EmailAddresses) {
                        try {
                            Set-MailBox -Identity $365User.userPrincipalName -EmailAddresses @{Add = $ADproxyAddress} -Confirm:$false
                            Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365 `nAdded proxy address: $ADProxyAddress `nTo user: $($365User.DisplayName)"
                        }
                        catch {
                            $ErrorMessage = $_.Exception.Message
                            Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nAttempted to add proxy address: $ADProxyAddress `nTo user: $($365User.DisplayName) `nError was: $ErrorMessage"
                        }
                        
                    }
                }
            }
            
        }
    }
    
    if ( (Measure-Object -InputObject $StringForEventMessage -Line).Lines -ge 2 ) {
        Write-ScriptEvent -EntryType Warning -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nREMOVED proxy addresses:`n $StringForEventMessage"
    }
    
}
