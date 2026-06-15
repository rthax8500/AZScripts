<#
.SYNOPSIS
    Bulk-grants mailbox permissions (FullAccess, SendAs, SendOnBehalf) from a CSV mapping.

.DESCRIPTION
    Reads a CSV containing Mailbox/User/Permission rows and applies the specified
    permission to each pair. Supports all three common mailbox access types in one run.

    Includes -WhatIf support, confirmation prompt, transcript logging, and a results CSV.

.PARAMETER CsvPath
    Full path to the CSV file. Required columns: Mailbox, User.
    Optional column: Permission (FullAccess, SendAs, SendOnBehalf). Defaults to FullAccess.

.PARAMETER AutoMapping
    For FullAccess only. When True, Outlook will automatically add the shared mailbox
    to the user's profile. Default is True. Set False for high-volume shared mailboxes
    where auto-mapping would clutter user profiles.

.PARAMETER LogPath
    Folder where the run log will be written. Defaults to script directory.

.PARAMETER Force
    Skip the interactive confirmation prompt.

.EXAMPLE
    .\Grant-MailboxFullAccess.ps1 -CsvPath "C:\Temp\permissions.csv" -WhatIf

.EXAMPLE
    .\Grant-MailboxFullAccess.ps1 -CsvPath "C:\Temp\permissions.csv" -AutoMapping $false
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,         # <-- REPLACE at runtime: e.g. "C:\Temp\permissions.csv"

    [Parameter(Mandatory = $false)]
    [bool]$AutoMapping = $true,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path),

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "GrantMailboxAccess_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host "=== Grant Mailbox Access (Bulk) ===" -ForegroundColor Cyan
Write-Host "CSV:          $CsvPath"
Write-Host "AutoMapping:  $AutoMapping (applies to FullAccess only)"
Write-Host "Log file:     $logFile`n"

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

# ---------- Load CSV ----------
$rows = Import-Csv -Path $CsvPath

# Required columns
foreach ($required in 'Mailbox', 'User') {
    if (-not ($rows | Get-Member -Name $required -MemberType NoteProperty)) {
        Write-Error "CSV must contain a '$required' column."
        Stop-Transcript | Out-Null
        return
    }
}
$hasPermissionCol = [bool]($rows | Get-Member -Name 'Permission' -MemberType NoteProperty)

# Validate Permission values up front
$validPerms = @('FullAccess', 'SendAs', 'SendOnBehalf')
if ($hasPermissionCol) {
    $badRows = $rows | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.Permission) -and
        $validPerms -notcontains $_.Permission
    }
    if ($badRows) {
        Write-Error "Invalid Permission value(s) in CSV. Allowed: $($validPerms -join ', ')"
        $badRows | ForEach-Object { Write-Host ("  Bad row -> Mailbox={0} User={1} Permission={2}" -f $_.Mailbox, $_.User, $_.Permission) }
        Stop-Transcript | Out-Null
        return
    }
}

# ---------- Confirmation gate ----------
$rowCount = ($rows | Where-Object {
    -not [string]::IsNullOrWhiteSpace($_.Mailbox) -and
    -not [string]::IsNullOrWhiteSpace($_.User)
}).Count

Write-Host "Permission grants to process: $rowCount`n" -ForegroundColor Yellow
if (-not $Force -and -not $WhatIfPreference) {
    $confirm = Read-Host "Type 'YES' to proceed (anything else cancels)"
    if ($confirm -ne 'YES') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        Stop-Transcript | Out-Null
        return
    }
}

# ---------- Process ----------
$results = [System.Collections.Generic.List[object]]::new()

foreach ($row in $rows) {
    $mailbox    = $row.Mailbox.Trim()
    $user       = $row.User.Trim()
    $permission = if ($hasPermissionCol -and -not [string]::IsNullOrWhiteSpace($row.Permission)) {
        $row.Permission.Trim()
    } else {
        'FullAccess'
    }

    if ([string]::IsNullOrWhiteSpace($mailbox) -or [string]::IsNullOrWhiteSpace($user)) { continue }

    $status = 'Unknown'
    $detail = ''

    try {
        # Validate both ends exist
        $null = Get-Mailbox   -Identity $mailbox -ErrorAction Stop
        $null = Get-Recipient -Identity $user    -ErrorAction Stop

        switch ($permission) {

            'FullAccess' {
                # Check for existing identical permission
                $existing = Get-MailboxPermission -Identity $mailbox -User $user -ErrorAction SilentlyContinue |
                            Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited -and $_.Deny -eq $false }
                if ($existing) {
                    $status = 'Skipped'; $detail = 'FullAccess already granted'
                }
                elseif ($PSCmdlet.ShouldProcess("$user -> $mailbox", "Grant FullAccess (AutoMapping=$AutoMapping)")) {
                    Add-MailboxPermission -Identity $mailbox -User $user `
                        -AccessRights FullAccess -InheritanceType All -AutoMapping:$AutoMapping `
                        -ErrorAction Stop | Out-Null
                    $status = 'Granted'; $detail = "FullAccess (AutoMapping=$AutoMapping)"
                }
                else {
                    $status = 'WhatIf'; $detail = 'Would grant FullAccess'
                }
            }

            'SendAs' {
                $existing = Get-RecipientPermission -Identity $mailbox -Trustee $user -ErrorAction SilentlyContinue |
                            Where-Object { $_.AccessRights -contains 'SendAs' }
                if ($existing) {
                    $status = 'Skipped'; $detail = 'SendAs already granted'
                }
                elseif ($PSCmdlet.ShouldProcess("$user -> $mailbox", "Grant SendAs")) {
                    Add-RecipientPermission -Identity $mailbox -Trustee $user `
                        -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                    $status = 'Granted'; $detail = 'SendAs'
                }
                else {
                    $status = 'WhatIf'; $detail = 'Would grant SendAs'
                }
            }

            'SendOnBehalf' {
                $mbx = Get-Mailbox -Identity $mailbox
                $alreadyDelegated = $false
                if ($mbx.GrantSendOnBehalfTo) {
                    $userRecipient = Get-Recipient -Identity $user
                    $alreadyDelegated = $mbx.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() } | Where-Object { $_ -eq $userRecipient.Identity.ToString() }
                }
                if ($alreadyDelegated) {
                    $status = 'Skipped'; $detail = 'SendOnBehalf already granted'
                }
                elseif ($PSCmdlet.ShouldProcess("$user -> $mailbox", "Grant SendOnBehalf")) {
                    Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo @{Add=$user} -ErrorAction Stop
                    $status = 'Granted'; $detail = 'SendOnBehalf'
                }
                else {
                    $status = 'WhatIf'; $detail = 'Would grant SendOnBehalf'
                }
            }
        }
    }
    catch {
        $status = 'Failed'; $detail = $_.Exception.Message
    }

    $results.Add([PSCustomObject]@{
        Mailbox    = $mailbox
        User       = $user
        Permission = $permission
        Status     = $status
        Detail     = $detail
    })

    Write-Host ("[{0,-7}] {1,-22} {2} -> {3}  ({4})" -f $status, $permission, $user, $mailbox, $detail)
}

# ---------- Summary ----------
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
$results | Group-Object Status | ForEach-Object {
    Write-Host ("{0,-8}: {1}" -f $_.Name, $_.Count)
}

Write-Host "`nBy permission type:" -ForegroundColor Yellow
$results | Group-Object Permission | ForEach-Object {
    Write-Host ("  {0,-15} {1}" -f $_.Name, $_.Count)
}

$resultsCsv = Join-Path $LogPath "GrantMailboxAccess_Results_$timestamp.csv"
$results | Export-Csv -Path $resultsCsv -NoTypeInformation -Encoding UTF8
Write-Host "`nResults exported: $resultsCsv" -ForegroundColor Green

Stop-Transcript | Out-Null