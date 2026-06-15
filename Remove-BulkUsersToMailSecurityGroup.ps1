<#
.SYNOPSIS
    Bulk-removes users listed in a CSV from a mail-enabled security group in Exchange Online.

.DESCRIPTION
    Reads UserPrincipalName values from a CSV, validates each user, skips users
    not currently in the group, and logs all actions to a transcript file.
    Prompts for confirmation before live execution unless -Force or -WhatIf is used.

.PARAMETER GroupIdentity
    The mail-enabled security group's primary SMTP address, alias, or DisplayName.

.PARAMETER CsvPath
    Full path to the CSV file. Must contain a 'UserPrincipalName' column.

.PARAMETER LogPath
    Folder where the run log will be written. Defaults to script directory.

.PARAMETER Force
    Skip the interactive confirmation prompt. Use with caution in production.

.EXAMPLE
    .\Remove-BulkUsersFromMailSecurityGroup.ps1 -GroupIdentity "ProjectAlpha@contoso.com" -CsvPath "C:\Temp\users.csv" -WhatIf

.EXAMPLE
    .\Remove-BulkUsersFromMailSecurityGroup.ps1 -GroupIdentity "ProjectAlpha@contoso.com" -CsvPath "C:\Temp\users.csv"
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$GroupIdentity,   # <-- REPLACE at runtime: e.g. "GroupAlias@yourdomain.com"

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,         # <-- REPLACE at runtime: e.g. "C:\Temp\users.csv"

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path),

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "RemoveFromGroup_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host "=== Bulk Remove Users from Mail-Enabled Security Group ===" -ForegroundColor Cyan
Write-Host "Group:    $GroupIdentity"
Write-Host "CSV:      $CsvPath"
Write-Host "Log file: $logFile`n"

# ---------- Connect ----------
try {
    if (-not (Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' })) {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    } else {
        Write-Host "Already connected to Exchange Online." -ForegroundColor Green
    }
}
catch {
    Write-Error "Failed to connect to Exchange Online: $_"
    Stop-Transcript | Out-Null
    return
}

# ---------- Validate group ----------
try {
    $group = Get-DistributionGroup -Identity $GroupIdentity -ErrorAction Stop
    if ($group.RecipientTypeDetails -ne 'MailUniversalSecurityGroup') {
        Write-Warning "Target group type is '$($group.RecipientTypeDetails)', not a mail-enabled security group. Aborting."
        Stop-Transcript | Out-Null
        return
    }
    Write-Host "Group validated: $($group.DisplayName) [$($group.PrimarySmtpAddress)]`n" -ForegroundColor Green
}
catch {
    Write-Error "Group '$GroupIdentity' not found or not accessible: $_"
    Stop-Transcript | Out-Null
    return
}

# ---------- Load CSV ----------
$users = Import-Csv -Path $CsvPath
if (-not ($users | Get-Member -Name 'UserPrincipalName' -MemberType NoteProperty)) {
    Write-Error "CSV must contain a 'UserPrincipalName' column."
    Stop-Transcript | Out-Null
    return
}

# Cache existing membership for fast non-member checks
$existingMembers = Get-DistributionGroupMember -Identity $GroupIdentity -ResultSize Unlimited |
                   Select-Object -ExpandProperty PrimarySmtpAddress

# ---------- Confirmation gate ----------
$userCount = ($users | Where-Object { -not [string]::IsNullOrWhiteSpace($_.UserPrincipalName) }).Count

if (-not $Force -and -not $WhatIfPreference) {
    Write-Host "`nYou are about to remove up to $userCount user(s) from '$($group.DisplayName)'." -ForegroundColor Yellow
    $confirm = Read-Host "Type 'YES' to proceed (anything else cancels)"
    if ($confirm -ne 'YES') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        Stop-Transcript | Out-Null
        return
    }
}

# ---------- Process ----------
$results = [System.Collections.Generic.List[object]]::new()

foreach ($row in $users) {
    $upn = $row.UserPrincipalName.Trim()
    if ([string]::IsNullOrWhiteSpace($upn)) { continue }

    $status = 'Unknown'
    $detail = ''

    try {
        $recipient = Get-Recipient -Identity $upn -ErrorAction Stop

        if ($existingMembers -notcontains $recipient.PrimarySmtpAddress) {
            $status = 'Skipped'
            $detail = 'Not a current member'
        }
        elseif ($PSCmdlet.ShouldProcess($upn, "Remove from $GroupIdentity")) {
            Remove-DistributionGroupMember `
                -Identity $GroupIdentity `
                -Member $upn `
                -BypassSecurityGroupManagerCheck `
                -Confirm:$false `
                -ErrorAction Stop
            $status = 'Removed'
            $detail = 'Success'
        }
        else {
            $status = 'WhatIf'
            $detail = 'Would remove (dry run)'
        }
    }
    catch {
        $status = 'Failed'
        $detail = $_.Exception.Message
    }

    $results.Add([PSCustomObject]@{
        UserPrincipalName = $upn
        Status            = $status
        Detail            = $detail
    })

    Write-Host ("[{0,-7}] {1}  -  {2}" -f $status, $upn, $detail)
}

# ---------- Summary ----------
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
$results | Group-Object Status | ForEach-Object {
    Write-Host ("{0,-8}: {1}" -f $_.Name, $_.Count)
}

# Export results next to the log
$resultsCsv = Join-Path $LogPath "RemoveFromGroup_Results_$timestamp.csv"
$results | Export-Csv -Path $resultsCsv -NoTypeInformation
Write-Host "`nResults exported: $resultsCsv" -ForegroundColor Green

Stop-Transcript | Out-Null