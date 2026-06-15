<#
.SYNOPSIS
    Bulk-sets or disables Out-of-Office (Automatic Replies) for users listed in a CSV.

.DESCRIPTION
    Reads UserPrincipalName values from a CSV and applies Automatic Replies settings
    to each mailbox. Supports scheduled or always-on OOO, separate internal/external
    messages, external audience scope, and a Disable action to turn OOO off in bulk.

    The CSV may optionally include InternalMessage and ExternalMessage columns to
    override the parameter values per-user.

.PARAMETER CsvPath
    Full path to the CSV file. Must contain a 'UserPrincipalName' column.
    Optional columns: 'InternalMessage', 'ExternalMessage'.

.PARAMETER Action
    Enable (set OOO) or Disable (turn OOO off). Default is Enable.

.PARAMETER InternalMessage
    OOO message shown to internal senders. Required when Action is Enable
    unless every CSV row provides its own InternalMessage.

.PARAMETER ExternalMessage
    OOO message shown to external senders. If omitted, falls back to InternalMessage.

.PARAMETER ExternalAudience
    Who external messages are sent to. Valid values: None, Known, All. Default is All.
    'Known' restricts external replies to senders in the user's contacts.

.PARAMETER StartTime
    Optional. Start date/time for scheduled OOO. Omit for "always on".

.PARAMETER EndTime
    Optional. End date/time for scheduled OOO. Required if StartTime is provided.

.PARAMETER LogPath
    Folder where the run log will be written. Defaults to script directory.

.PARAMETER Force
    Skip the interactive confirmation prompt.

.EXAMPLE
    # Holiday OOO from CSV with one shared message, scheduled
    .\Set-OutOfOfficeBulk.ps1 `
        -CsvPath "C:\Temp\holiday_users.csv" `
        -InternalMessage "Office is closed Dec 24-26. Will reply on the 27th." `
        -StartTime "2026-12-24 00:00" -EndTime "2026-12-26 23:59"

.EXAMPLE
    # Disable OOO for a list
    .\Set-OutOfOfficeBulk.ps1 -CsvPath "C:\Temp\users.csv" -Action Disable

.EXAMPLE
    # Per-user messages from CSV, dry run
    .\Set-OutOfOfficeBulk.ps1 -CsvPath "C:\Temp\peruser.csv" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,         # <-- REPLACE at runtime: e.g. "C:\Temp\users.csv"

    [Parameter(Mandatory = $false)]
    [ValidateSet('Enable', 'Disable')]
    [string]$Action = 'Enable',

    [Parameter(Mandatory = $false)]
    [string]$InternalMessage,

    [Parameter(Mandatory = $false)]
    [string]$ExternalMessage,

    [Parameter(Mandatory = $false)]
    [ValidateSet('None', 'Known', 'All')]
    [string]$ExternalAudience = 'All',

    [Parameter(Mandatory = $false)]
    [datetime]$StartTime,

    [Parameter(Mandatory = $false)]
    [datetime]$EndTime,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path),

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "SetOOOBulk_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host "=== Set Out-of-Office (Bulk) ===" -ForegroundColor Cyan
Write-Host "Action:    $Action"
Write-Host "CSV:       $CsvPath"
if ($Action -eq 'Enable') {
    Write-Host "Audience:  $ExternalAudience"
    if ($StartTime) { Write-Host "Schedule:  $StartTime  ->  $EndTime" }
    else            { Write-Host "Schedule:  Always on (no time bound)" }
}
Write-Host "Log file:  $logFile`n"

# ---------- Validate parameter combinations ----------
if ($Action -eq 'Enable') {
    if ($StartTime -and -not $EndTime) {
        Write-Error "EndTime is required when StartTime is specified."
        Stop-Transcript | Out-Null
        return
    }
    if ($StartTime -and $EndTime -and $EndTime -le $StartTime) {
        Write-Error "EndTime must be later than StartTime."
        Stop-Transcript | Out-Null
        return
    }
}

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
$users = Import-Csv -Path $CsvPath
if (-not ($users | Get-Member -Name 'UserPrincipalName' -MemberType NoteProperty)) {
    Write-Error "CSV must contain a 'UserPrincipalName' column."
    Stop-Transcript | Out-Null
    return
}

$hasInternalCol = [bool]($users | Get-Member -Name 'InternalMessage' -MemberType NoteProperty)
$hasExternalCol = [bool]($users | Get-Member -Name 'ExternalMessage' -MemberType NoteProperty)

# ---------- Validate Enable has a message source ----------
if ($Action -eq 'Enable') {
    $missingMessage = $false
    if (-not $InternalMessage) {
        # Every row must supply its own InternalMessage if no parameter default
        if (-not $hasInternalCol) {
            $missingMessage = $true
        } else {
            foreach ($u in $users) {
                if ([string]::IsNullOrWhiteSpace($u.InternalMessage)) {
                    $missingMessage = $true; break
                }
            }
        }
    }
    if ($missingMessage) {
        Write-Error "Action is Enable but no InternalMessage is set (parameter or CSV column)."
        Stop-Transcript | Out-Null
        return
    }
}

# ---------- Confirmation gate ----------
$userCount = ($users | Where-Object { -not [string]::IsNullOrWhiteSpace($_.UserPrincipalName) }).Count

Write-Host "Targeted mailboxes: $userCount`n" -ForegroundColor Yellow
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

foreach ($row in $users) {
    $upn = $row.UserPrincipalName.Trim()
    if ([string]::IsNullOrWhiteSpace($upn)) { continue }

    $status = 'Unknown'
    $detail = ''

    try {
        # Validate mailbox exists
        $null = Get-Mailbox -Identity $upn -ErrorAction Stop

        if ($Action -eq 'Disable') {
            if ($PSCmdlet.ShouldProcess($upn, "Disable OOO")) {
                Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Disabled -ErrorAction Stop
                $status = 'Disabled'; $detail = 'Success'
            } else {
                $status = 'WhatIf'; $detail = 'Would disable OOO'
            }
        }
        else {
            # Resolve message contents (CSV row > parameter)
            $rowInternal = if ($hasInternalCol -and -not [string]::IsNullOrWhiteSpace($row.InternalMessage)) { $row.InternalMessage } else { $InternalMessage }
            $rowExternal = if ($hasExternalCol -and -not [string]::IsNullOrWhiteSpace($row.ExternalMessage)) { $row.ExternalMessage } elseif ($ExternalMessage) { $ExternalMessage } else { $rowInternal }

            $autoReplyState = if ($StartTime) { 'Scheduled' } else { 'Enabled' }

            $params = @{
                Identity         = $upn
                AutoReplyState   = $autoReplyState
                InternalMessage  = $rowInternal
                ExternalMessage  = $rowExternal
                ExternalAudience = $ExternalAudience
                ErrorAction      = 'Stop'
            }
            if ($StartTime) { $params['StartTime'] = $StartTime; $params['EndTime'] = $EndTime }

            if ($PSCmdlet.ShouldProcess($upn, "Set OOO ($autoReplyState, audience=$ExternalAudience)")) {
                Set-MailboxAutoReplyConfiguration @params
                $status = 'Enabled'; $detail = $autoReplyState
            } else {
                $status = 'WhatIf'; $detail = "Would set OOO ($autoReplyState)"
            }
        }
    }
    catch {
        $status = 'Failed'; $detail = $_.Exception.Message
    }

    $results.Add([PSCustomObject]@{
        UserPrincipalName = $upn
        Action            = $Action
        Status            = $status
        Detail            = $detail
    })

    Write-Host ("[{0,-9}] {1}  -  {2}" -f $status, $upn, $detail)
}

# ---------- Summary ----------
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
$results | Group-Object Status | ForEach-Object {
    Write-Host ("{0,-9}: {1}" -f $_.Name, $_.Count)
}

$resultsCsv = Join-Path $LogPath "SetOOOBulk_Results_$timestamp.csv"
$results | Export-Csv -Path $resultsCsv -NoTypeInformation -Encoding UTF8
Write-Host "`nResults exported: $resultsCsv" -ForegroundColor Green

Stop-Transcript | Out-Null