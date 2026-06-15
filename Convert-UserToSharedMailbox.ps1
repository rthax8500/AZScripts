<#
.SYNOPSIS
    Converts user mailboxes to shared mailboxes for offboarding workflows.

.DESCRIPTION
    Bulk converts user mailboxes to shared mailboxes, with optional delegate access
    grants, auto-reply configuration, and GAL visibility control. Includes pre-flight
    size checks, snapshot artifacts for rollback reference, and the standard
    transcript log + results CSV pattern.

    Standard offboarding workflow:
      User leaves -> Convert mailbox to Shared (this script) -> Grant manager access
      -> Set "no longer here" auto-reply -> Disable user account separately.

.PARAMETER UsersCsvPath
    Required. CSV with 'UserPrincipalName' column listing mailboxes to convert.

.PARAMETER DelegateCsvPath
    Optional. CSV with 'UserPrincipalName' and 'Delegate' columns, mapping each
    mailbox to one or more delegates who should receive Full Access.

.PARAMETER AutoReplyMessage
    Optional. If provided, sets an always-on auto-reply with this message before
    conversion. Both internal and external senders receive it.

.PARAMETER HideFromGAL
    When set, hides the converted mailbox from the Global Address List. Recommended
    for offboarded-employee mailboxes that aren't true team mailboxes.

.PARAMETER LogPath
    Folder where logs and snapshots will be written. Defaults to script directory.

.PARAMETER Force
    Skip the interactive confirmation prompt.

.EXAMPLE
    .\Convert-UserToSharedMailbox.ps1 `
        -UsersCsvPath "C:\Temp\offboarded.csv" `
        -DelegateCsvPath "C:\Temp\delegates.csv" `
        -AutoReplyMessage "This employee has left the company. Please contact your account manager." `
        -HideFromGAL `
        -WhatIf

.EXAMPLE
    .\Convert-UserToSharedMailbox.ps1 -UsersCsvPath "C:\Temp\offboarded.csv"
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$UsersCsvPath,        # <-- REPLACE at runtime: e.g. "C:\Temp\offboarded.csv"

    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$DelegateCsvPath,

    [Parameter(Mandatory = $false)]
    [string]$AutoReplyMessage,

    [Parameter(Mandatory = $false)]
    [switch]$HideFromGAL,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path),

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "ConvertToShared_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host "=== Convert User Mailbox to Shared ===" -ForegroundColor Cyan
Write-Host "Users CSV:    $UsersCsvPath"
if ($DelegateCsvPath) { Write-Host "Delegate CSV: $DelegateCsvPath" }
if ($AutoReplyMessage) { Write-Host "Auto-reply:   Will be set" }
Write-Host "Hide from GAL: $HideFromGAL"
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

# ---------- Helper: parse Exchange size ----------
function Convert-ExchangeSizeToBytes {
    param([string]$SizeString)
    if ([string]::IsNullOrWhiteSpace($SizeString)) { return 0 }
    if ($SizeString -match '\(([\d,]+) bytes\)') { return [int64]($Matches[1] -replace ',', '') }
    return 0
}

# ---------- Load CSVs ----------
$users = Import-Csv -Path $UsersCsvPath
if (-not ($users | Get-Member -Name 'UserPrincipalName' -MemberType NoteProperty)) {
    Write-Error "Users CSV must contain a 'UserPrincipalName' column."
    Stop-Transcript | Out-Null
    return
}

$delegateMap = @{}
if ($DelegateCsvPath) {
    $delegateRows = Import-Csv -Path $DelegateCsvPath
    foreach ($col in 'UserPrincipalName', 'Delegate') {
        if (-not ($delegateRows | Get-Member -Name $col -MemberType NoteProperty)) {
            Write-Error "Delegate CSV must contain '$col' column."
            Stop-Transcript | Out-Null
            return
        }
    }
    foreach ($d in $delegateRows) {
        $key = $d.UserPrincipalName.Trim().ToLower()
        if (-not $delegateMap.ContainsKey($key)) { $delegateMap[$key] = @() }
        $delegateMap[$key] += $d.Delegate.Trim()
    }
    Write-Host "Loaded delegates for $($delegateMap.Count) mailbox(es).`n" -ForegroundColor Yellow
}

# ---------- Confirmation gate ----------
$count = ($users | Where-Object { -not [string]::IsNullOrWhiteSpace($_.UserPrincipalName) }).Count
Write-Host "Mailboxes to convert: $count" -ForegroundColor Yellow
Write-Host "WARNING: Conversion is reversible (Set-Mailbox -Type Regular) but the user account license must be reattached if you reverse it." -ForegroundColor Yellow

if (-not $Force -and -not $WhatIfPreference) {
    $confirm = Read-Host "`nType 'CONVERT' to proceed (anything else cancels)"
    if ($confirm -ne 'CONVERT') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        Stop-Transcript | Out-Null
        return
    }
}

# ---------- Process ----------
$results   = [System.Collections.Generic.List[object]]::new()
$snapshots = [System.Collections.Generic.List[object]]::new()

foreach ($row in $users) {
    $upn = $row.UserPrincipalName.Trim()
    if ([string]::IsNullOrWhiteSpace($upn)) { continue }

    Write-Host "`n--- Processing $upn ---" -ForegroundColor Cyan

    $status = 'Unknown'
    $detail = ''
    $sizeGB = $null
    $delegatesGranted = 0
    $delegatesFailed  = 0

    try {
        # Pre-flight: validate mailbox
        $mbx = Get-Mailbox -Identity $upn -ErrorAction Stop

        if ($mbx.RecipientTypeDetails -ne 'UserMailbox') {
            $status = 'Skipped'
            $detail = "Already $($mbx.RecipientTypeDetails); no conversion needed"
            Write-Host "  $detail" -ForegroundColor Yellow
        }
        else {
            # Get size for pre-flight
            $stats = Get-MailboxStatistics -Identity $mbx.Identity -ErrorAction Stop
            $sizeBytes = Convert-ExchangeSizeToBytes -SizeString $stats.TotalItemSize.ToString()
            $sizeGB    = [math]::Round($sizeBytes / 1GB, 2)

            Write-Host "  Current size: $sizeGB GB"

            if ($sizeGB -gt 100) {
                $status = 'Failed'
                $detail = "Mailbox is $sizeGB GB - exceeds 100 GB shared mailbox limit. Archive/cleanup required first."
                Write-Host "  $detail" -ForegroundColor Red
            }
            else {
                if ($sizeGB -gt 50) {
                    Write-Host "  NOTE: $sizeGB GB exceeds the free 50 GB tier. An Exchange Online Plan 2 license must remain attached." -ForegroundColor Yellow
                }

                # Snapshot the original state
                $snapshots.Add([PSCustomObject]@{
                    UserPrincipalName    = $upn
                    DisplayName          = $mbx.DisplayName
                    PrimarySmtpAddress   = $mbx.PrimarySmtpAddress
                    OriginalType         = $mbx.RecipientTypeDetails
                    SizeGB               = $sizeGB
                    EmailAddresses       = ($mbx.EmailAddresses -join '; ')
                    HiddenFromAddressList = $mbx.HiddenFromAddressListsEnabled
                    LitigationHoldEnabled = $mbx.LitigationHoldEnabled
                    ArchiveStatus        = $mbx.ArchiveStatus
                    SnapshotTimestamp    = (Get-Date)
                })

                # Set auto-reply BEFORE conversion (cleaner)
                if ($AutoReplyMessage) {
                    if ($PSCmdlet.ShouldProcess($upn, "Set always-on auto-reply")) {
                        try {
                            Set-MailboxAutoReplyConfiguration -Identity $upn `
                                -AutoReplyState Enabled `
                                -InternalMessage $AutoReplyMessage `
                                -ExternalMessage $AutoReplyMessage `
                                -ExternalAudience All `
                                -ErrorAction Stop
                            Write-Host "  Auto-reply set" -ForegroundColor Green
                        }
                        catch {
                            Write-Warning "  Auto-reply failed (continuing): $_"
                        }
                    }
                }

                # The actual conversion
                if ($PSCmdlet.ShouldProcess($upn, "Convert to Shared mailbox")) {
                    Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
                    Write-Host "  Converted to Shared" -ForegroundColor Green
                    $status = 'Converted'
                    $detail = "Size: $sizeGB GB"
                }
                else {
                    $status = 'WhatIf'
                    $detail = "Would convert ($sizeGB GB)"
                    continue   # Skip delegate/GAL steps in WhatIf mode for this user
                }

                # Hide from GAL if requested
                if ($HideFromGAL) {
                    if ($PSCmdlet.ShouldProcess($upn, "Hide from GAL")) {
                        try {
                            Set-Mailbox -Identity $upn -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                            Write-Host "  Hidden from GAL" -ForegroundColor Green
                        }
                        catch {
                            Write-Warning "  Hide-from-GAL failed: $_"
                        }
                    }
                }

                # Grant delegates
                $key = $upn.ToLower()
                if ($delegateMap.ContainsKey($key)) {
                    foreach ($delegate in $delegateMap[$key]) {
                        if ($PSCmdlet.ShouldProcess("$delegate -> $upn", "Grant FullAccess")) {
                            try {
                                # Check for existing permission first
                                $existing = Get-MailboxPermission -Identity $upn -User $delegate -ErrorAction SilentlyContinue |
                                            Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited }
                                if ($existing) {
                                    Write-Host "  Delegate already has FullAccess: $delegate" -ForegroundColor Yellow
                                }
                                else {
                                    Add-MailboxPermission -Identity $upn -User $delegate `
                                        -AccessRights FullAccess -InheritanceType All -AutoMapping:$true `
                                        -ErrorAction Stop | Out-Null
                                    Write-Host "  Granted FullAccess to: $delegate" -ForegroundColor Green
                                    $delegatesGranted++
                                }
                            }
                            catch {
                                Write-Warning "  Delegate grant failed for $delegate : $_"
                                $delegatesFailed++
                            }
                        }
                    }
                }
            }
        }
    }
    catch {
        $status = 'Failed'
        $detail = $_.Exception.Message
        Write-Host "  ERROR: $detail" -ForegroundColor Red
    }

    $results.Add([PSCustomObject]@{
        UserPrincipalName = $upn
        Status            = $status
        OriginalSizeGB    = $sizeGB
        Detail            = $detail
        DelegatesGranted  = $delegatesGranted
        DelegatesFailed   = $delegatesFailed
    })
}

# ---------- Summary ----------
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
$results | Group-Object Status | ForEach-Object {
    Write-Host ("{0,-10}: {1}" -f $_.Name, $_.Count)
}

$totalDelegates = ($results | Measure-Object -Property DelegatesGranted -Sum).Sum
Write-Host ("Delegates granted: {0}" -f $totalDelegates)

# Export results and snapshot
$resultsCsv  = Join-Path $LogPath "ConvertToShared_Results_$timestamp.csv"
$snapshotCsv = Join-Path $LogPath "ConvertToShared_Snapshot_$timestamp.csv"
$results   | Export-Csv -Path $resultsCsv  -NoTypeInformation -Encoding UTF8
if ($snapshots.Count -gt 0) {
    $snapshots | Export-Csv -Path $snapshotCsv -NoTypeInformation -Encoding UTF8
    Write-Host "`nSnapshot exported: $snapshotCsv" -ForegroundColor Green
}
Write-Host "Results exported:  $resultsCsv" -ForegroundColor Green

Stop-Transcript | Out-Null