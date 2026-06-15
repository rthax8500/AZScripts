<#
.SYNOPSIS
    Bulk-adds users from a CSV to a mail-enabled security group in Exchange Online.

.DESCRIPTION
    Reads UserPrincipalName values from a CSV, validates each user exists,
    skips users already in the group, and logs all actions to a transcript file.

.PARAMETER GroupIdentity
    The mail-enabled security group's primary SMTP address, alias, or DisplayName.

.PARAMETER CsvPath
    Full path to the CSV file. Must contain a 'UserPrincipalName' column.

.PARAMETER LogPath
    Folder where the run log will be written. Defaults to script directory.

.EXAMPLE
    .\Add-BulkUsersToMailSecurityGroup.ps1 -GroupIdentity "ProjectAlpha@contoso.com" -CsvPath "C:\Temp\users.csv"

.EXAMPLE
    .\Add-BulkUsersToMailSecurityGroup.ps1 -GroupIdentity "ProjectAlpha@contoso.com" -CsvPath "C:\Temp\users.csv" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$GroupIdentity,   # <-- REPLACE at runtime: e.g. "GroupAlias@yourdomain.com"

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,         # <-- REPLACE at runtime: e.g. "C:\Temp\users.csv"

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "AddToGroup_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host "=== Bulk Add Users to Mail-Enabled Security Group ===" -ForegroundColor Cyan
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

# Cache existing membership for fast duplicate checks
$existingMembers = Get-DistributionGroupMember -Identity $GroupIdentity -ResultSize Unlimited |
                   Select-Object -ExpandProperty PrimarySmtpAddress

# ---------- Process ----------
$results = [System.Collections.Generic.List[object]]::new()

foreach ($row in $users) {
    $upn = $row.UserPrincipalName.Trim()
    if ([string]::IsNullOrWhiteSpace($upn)) { continue }

    $status = 'Unknown'
    $detail = ''

    try {
        $recipient = Get-Recipient -Identity $upn -ErrorAction Stop

        if ($existingMembers -contains $recipient.PrimarySmtpAddress) {
            $status = 'Skipped'
            $detail = 'Already a member'
        }
        elseif ($PSCmdlet.ShouldProcess($upn, "Add to $GroupIdentity")) {
            Add-DistributionGroupMember -Identity $GroupIdentity -Member $upn -ErrorAction Stop
            $status = 'Added'
            $detail = 'Success'
        }
        else {
            $status = 'WhatIf'
            $detail = 'Would add (dry run)'
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
$resultsCsv = Join-Path $LogPath "AddToGroup_Results_$timestamp.csv"
$results | Export-Csv -Path $resultsCsv -NoTypeInformation
Write-Host "`nResults exported: $resultsCsv" -ForegroundColor Green

Stop-Transcript | Out-Null