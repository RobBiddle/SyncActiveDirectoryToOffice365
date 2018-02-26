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

function Get-UsersToSync ($DomainControllerFQDN, $EmailDomain) {
    $Users365 = Get-MsolUser
    $filter = "userPrincipalName -like `"*$($EmailDomain)`""
    $UsersAD = Get-ADUser -Properties * -Filter $filter -Server $DomainControllerFQDN
    $UsersToSync = $UsersAD
    # Add immutableID property to objects and populate values
    $UsersToSync | ForEach-Object {
        $guid = $_.ObjectGUID
        $immutableID = [System.Convert]::ToBase64String($guid.tobytearray())
        $_ | Add-Member -MemberType NoteProperty -Name ImmutableId -Value $immutableID -force -ErrorAction SilentlyContinue
    }

    # Add ExistsInOffice365 property to objects and set to $false
    $UsersToSync | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name ExistsInOffice365 -Value $false -force -ErrorAction SilentlyContinue
    }
    # Change SyncedTo365 value to true for objects that exist in both AD & 365
    $UsersToSync | Where-Object UserPrincipalName -In $Users365.UserPrincipalName | ForEach-Object {
        $_.ExistsInOffice365 = $true
    }
    $UsersToSync | Where-Object ImmutableId -In $Users365.ImmutableId | ForEach-Object {
        $_.ExistsInOffice365 = $true
    }

    # Add SyncComplete property to objects and set to $false
    $UsersToSync | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name SyncComplete -Value $false -force -ErrorAction SilentlyContinue
    }

    # Change SyncComplete value to true after verifying synced attributes
    $UsersToSync | Where-Object UserPrincipalName -In $Users365.UserPrincipalName | ForEach-Object {
        $ADUser = $_
        $365User = $Users365 | Where-Object UserPrincipalName -eq $ADUser.UserPrincipalName
        try {
            $365Mailbox = Get-Mailbox -Identity $365User.UserPrincipalName
        }
        catch {
            if ($error[0].CategoryInfo.Reason -match 'ManagementObjectNotFoundException') {
                Write-ScriptEvent -EntryType Warning -EventId 404 -Message "SyncActiveDirectoryToOffice365 Unable to find a Mailbox for user: $($365User.DisplayName)"
                Continue
            }
        }
        
        # Check basic properties
        if ( ($ADUser.GivenName -like $365User.FirstName) `
                -and ($ADUser.Surname -like $365User.LastName) `
                -and ($ADUser.DisplayName -like $365User.DisplayName) `
                -and ($ADUser.ImmutableId -like $365User.ImmutableId)
        ) {
            $BasicPropertiesMatch = $true
        }
        else {
            $BasicPropertiesMatch = $false
        }
        # Check Proxy Addresses
        if ($365User.proxyAddresses.count -gt 0) {
            if (Compare-Object -ReferenceObject $ADUser.proxyAddresses -DifferenceObject $365Mailbox.EmailAddresses -ErrorAction SilentlyContinue) {
                $proxyComparison = $false
            }
            else {
                $proxyComparison = $true
            }

            if ( ($BasicPropertiesMatch -eq $true) -and ($proxyChecks -eq $false) ) {
                $ADUser.SyncComplete = $true
            }

        }
        else {
            if ($BasicPropertiesMatch -eq $true) {
                $ADUser.SyncComplete = $true
            }

        }

    }
    # Return Objects
    $UsersToSync
}
