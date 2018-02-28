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

function Sync-ContactsToOffice365 ($ContactsToSync) {
    foreach ($ADContact in $ContactsToSync) {

        # If Contact is not in Office 365, create contact
        if ($ADContact.ExistsInOffice365 -eq $false) {
            try {
                Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365 `nAdding MailContact: $($ADContact.Name)"
                New-MailContact -Name $ADContact.Name `
                    -LastName $ADContact.sn `
                    -FirstName $ADContact.givenName `
                    -DisplayName $ADContact.DisplayName `
                    -ExternalEmailAddress $ADContact.mail;
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened :-(  `nError was: $ErrorMessage"
            }

            $365Contact = $null
            try {
                $365Contact = Get-MailContact $ADContact.Name
            }
            catch {
                $error.clear()
            }
            
            if ($365Contact) {
                $ADContact.ExistsInOffice365 = $true    
            }
        }

        # Sync properties
        if ($ADContact.ExistsInOffice365 -eq $true) {
            # If Contact is in Office 365, set attributes
            $365Contact = $null
            $365Contact = Get-MailContact $ADContact.Name -ErrorAction SilentlyContinue
            if ($365Contact.FirstName -notlike $ADContact.givenName) {
                try {
                    Set-Contact -Identity $365Contact.Identity `
                        -FirstName $ADContact.givenName;
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened :-(  `nError was: $ErrorMessage"
                }

            }

            if ($365Contact.LastName -notlike $ADContact.sn) {
                try {
                    Set-Contact -Identity $365Contact.Identity `
                        -LastName $ADContact.sn;
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened :-(  `nError was: $ErrorMessage"
                }

            }

            if ($365Contact.DisplayName -notlike $ADContact.DisplayName) {
                try {
                    Set-Contact -Identity $365Contact.Identity `
                        -DisplayName $ADContact.DisplayName;
                }
                catch {
                    $ErrorMessage = $_.Exception.Message
                    Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened :-(  `nError was: $ErrorMessage"
                }

            }

            # Sync proxyAddresses
            foreach ($365ProxyAddress in $365Contact.EmailAddresses) {
                if ($ADContact.proxyAddresses -inotcontains $365ProxyAddress) {
                    try {
                        Set-MailContact -Identity $365Contact.Identity -EmailAddresses @{Remove = $365ProxyAddress} -Confirm:$false
                        Write-ScriptEvent -EntryType Warning -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nREMOVED proxy address: $ADproxyAddress `nFrom MailContact: $($365Contact.Identity)"
                    }
                    catch {
                        $ErrorMessage = $_.Exception.Message
                        Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened :-(  `nError was: $ErrorMessage"
                    }

                }
            }

            foreach ($ADproxyAddress in $ADContact.proxyAddresses) {
                if ($365Contact.EmailAddresses -inotcontains $ADproxyAddress) {
                    try {
                        Set-MailContact -Identity $365Contact.Identity -EmailAddresses @{Add = $ADproxyAddress} -Confirm:$false
                        Write-ScriptEvent -EntryType Information -EventId 411 -Message "SyncActiveDirectoryToOffice365 `nAdded proxy address: $ADproxyAddress `nto MailContact: $($365Contact.Identity)"
                    }
                    catch [System.Management.Automation.RemoteException] {
                        $ErrorMessage = $_.Exception.Message
                        Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nAttempted to add $ADproxyAddress `nTo MailContact: $($365Contact.Identity)  `nError was: $ErrorMessage"
                    }
                    catch {
                        $ErrorMessage = $_.Exception.Message
                        Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened :-(  `nError was: $ErrorMessage"
                    }

                }
            }

        }
    
    }
    
}
