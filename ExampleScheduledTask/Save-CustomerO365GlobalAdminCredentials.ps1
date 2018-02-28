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
# Script to generate encrypted credential stores for Sync-ActiveDirectoryToOffice365 
Write-Host "In order for this to work properly the following must be true:" -ForegroundColor Yellow
Write-Host "1. This script should be running in a PowerShell console via 'Run as a different user' using the AD Service Account which will run the ScheduledTask for Sync-ActiveDirectoryToOffice365" -ForegroundColor Yellow
Write-Host "2. You should be in the same directory as SyncActiveDirectoryToOffice365_ScheduledTask.ps1" -ForegroundColor Yellow
Write-Host "If the above requirements have not been met hit Ctrl-C to exit this script!" -ForegroundColor Red
Pause

$CustomerNumber = Read-Host "Enter Customer Number"

$O365UserName = Read-Host "Enter O365 Global Admin UserName" -AsSecureString
$O365Password = Read-Host "Enter O365 Global Admin Password" -AsSecureString

$O365UserName | ConvertFrom-SecureString | Out-File .\$($CustomerNumber)_EncryptedO365UserName.txt -Force
$O365Password | ConvertFrom-SecureString | Out-File .\$($CustomerNumber)_EncryptedO365Password.txt -Force
