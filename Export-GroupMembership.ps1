<#
.SYNOPSIS
    Exports the current membership of a mail-enabled group in Exchange Online to a timestamped CSV.

.DESCRIPTION
    Retrieves all members of a specified group (mail-enabled security group, distribution list,
    or Microsoft 365 group) and exports them to a CSV formatted for direct reuse with the
    Add-BulkUsersToMailSecurityGroup and Remove-BulkUsersFromMailSecurityGroup scripts.

    Includes display name, UPN, primary SMTP, and recipient type for each member.

.PARAMETER GroupIdentity
    The group's primary SMTP address, alias, or DisplayName.

.PARAMETER ExportPath
    Folder where the export CSV will be written. Defaults to script directory.

.PARAMETER LogPath
    Folder where the run log will be written. Defaults to script directory.

.EXAMPLE
    .\Export-GroupMembership.ps1 -GroupIdentity "ProjectAlpha@contoso.com"

.EXAMPLE
    .\Export-GroupMembership.ps1 -GroupIdentity "ProjectAlpha@contoso.com" -ExportPath "C:\Backups\Groups"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$GroupIdentity,   # <-- REPLACE at runtime: e.g. "GroupAlias@yourdomain.com"

    [Parameter(Mandatory = $false)]
    [string]$ExportPath = (Split-Path -Parent $MyInvocation.MyCommand.Path),

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "ExportGroupMembership_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host "=== Export Group Membership ===" -ForegroundColor Cyan
Write-Host "Group: