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

function Get-ContactsToSync ($BaseOU, $DomainControllerFQDN, $EmailDomain) {

    # Trap Block to catch anything outside of a try/catch
    trap {
        $ErrorMessage = $_.Exception.Message
        Write-ScriptEvent -EntryType Error -EventId 187 -Message "SyncActiveDirectoryToOffice365 `nSomething Unexpected happened :-(  `nError was: $ErrorMessage"
        continue
    }

    # Get Contacts from Office 365
    $Contacts365 = Get-MailContact

    # Get Contacts from Active Directory
    $filter = 'objectClass -eq "contact"'
    $ContactsAD = Get-ADObject -SearchBase $BaseOU -Filter $filter -Properties * -Server $DomainControllerFQDN

    # Add ExistsInOffice365 Property to ContactsAD object
    $ContactsAD | ForEach-Object {
        $_ | Add-Member -NotePropertyName "ExistsInOffice365" -NotePropertyValue $false -Force -ErrorAction SilentlyContinue
    }

    # Add SyncComplete property to objects and set to $false
    $ContactsAD | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name SyncComplete -Value $false -Force -ErrorAction SilentlyContinue
    }

    # Determine Contacts to Sync
    $ContactsToSync = @()
    foreach ($ADcontact in $ContactsAD) {
        $365Contact = $Contacts365 | Where-Object PrimarySmtpAddress -Like $ADcontact.Mail
        
        # If no match on primary mail properties, add to ContactsToSync Array
        if (-NOT $365Contact) {
            $ContactsToSync += $ADcontact
            Continue
        }
        else {
            $ADcontact.ExistsInOffice365 = $true            
        }

        # Verify properties
        if ($365Contact.Name -notlike $ADcontact.Name) {
            $ContactsToSync += $ADcontact
            Continue
        }
        if ($365Contact.DisplayName -notlike $ADcontact.DisplayName) {
            $ContactsToSync += $ADcontact
            Continue
        }
        if (Compare-Object -ReferenceObject $ADcontact.ProxyAddresses -DifferenceObject $365Contact.EmailAddresses) {
            $ContactsToSync += $ADcontact
            Continue
        }
        # If all properties match
        $ADcontact.SyncComplete = $true
        $ContactsToSync += $ADcontact
    }

    # Return Objects
    $ContactsToSync
}
