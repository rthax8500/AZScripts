<#
.SYNOPSIS
    Generates a mailbox size and quota usage report for all mailboxes in Exchange Online.

.DESCRIPTION
    Read-only audit script. Retrieves mailbox statistics, quota settings, and last logon
    data for every mailbox (or a filtered subset) and exports a categorized CSV report.

    Includes calculated quota usage percentage and a categorized quota status flag for
    quick filtering in Excel.

.PARAMETER MailboxType
    Filter by mailbox type. Valid values: All, UserMailbox, SharedMailbox, RoomMailbox,
    EquipmentMailbox. Default is All.

.PARAMETER MinimumSizeMB
    Optional. Only report mailboxes larger than this size in MB. Useful for finding
    "top consumers" without seeing every empty mailbox. Default is 0 (all).

.PARAMETER LogPath
    Folder where the run log and report CSV will be written. Defaults to script directory.

.EXAMPLE
    .\Get-MailboxSizeReport.ps1

.EXAMPLE
    .\Get-MailboxSizeReport.ps1 -MailboxType UserMailbox -MinimumSizeMB 5000
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('All', 'UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox')]
    [string]$MailboxType = 'All',

    [Parameter(Mandatory = $false)]
    [int]$MinimumSizeMB = 0,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "MailboxSizeReport_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host "=== Mailbox Size Report ===" -ForegroundColor Cyan
Write-Host "Mailbox type filter: $MailboxType"
Write-Host "Minimum size (MB):   $MinimumSizeMB"
Write-Host "Log file:            $logFile`n"

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

# ---------- Helper: parse Exchange size strings ----------
# Exchange returns sizes as "12.34 GB (13,256,789,012 bytes)" - we want the bytes.
function Convert-ExchangeSizeToBytes {
    param([string]$SizeString)

    if ([string]::IsNullOrWhiteSpace($SizeString)) { return 0 }
    if ($SizeString -match '\(([\d,]+) bytes\)') {
        return [int64]($Matches[1] -replace ',', '')
    }
    return 0
}

# ---------- Helper: convert quota string ----------
# Quota properties return as "50 GB (53,687,091,200 bytes)" or "Unlimited".
function Convert-QuotaToBytes {
    param([string]$QuotaString)

    if ([string]::IsNullOrWhiteSpace($QuotaString) -or $QuotaString -eq 'Unlimited') {
        return $null
    }
    return Convert-ExchangeSizeToBytes -SizeString $QuotaString
}

# ---------- Get mailbox list ----------
Write-Host "Retrieving mailbox list..." -ForegroundColor Yellow
try {
    if ($MailboxType -eq 'All') {
        $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
    } else {
        $mailboxes = Get-Mailbox -RecipientTypeDetails $MailboxType -ResultSize Unlimited -ErrorAction Stop
    }
    Write-Host "Found $($mailboxes.Count) mailbox(es)." -ForegroundColor Green
}
catch {
    Write-Error "Failed to retrieve mailboxes: $_"
    Stop-Transcript | Out-Null
    return
}

# ---------- Process each mailbox ----------
$results = [System.Collections.Generic.List[object]]::new()
$i = 0

foreach ($mbx in $mailboxes) {
    $i++
    Write-Progress -Activity "Gathering statistics" -Status $mbx.PrimarySmtpAddress -PercentComplete (($i / $mailboxes.Count) * 100)

    try {
        $stats = Get-MailboxStatistics -Identity $mbx.Identity -ErrorAction Stop

        $sizeBytes        = Convert-ExchangeSizeToBytes -SizeString $stats.TotalItemSize.ToString()
        $deletedBytes     = Convert-ExchangeSizeToBytes -SizeString $stats.TotalDeletedItemSize.ToString()
        $sizeMB           = [math]::Round($sizeBytes / 1MB, 2)
        $sizeGB           = [math]::Round($sizeBytes / 1GB, 2)
        $deletedMB        = [math]::Round($deletedBytes / 1MB, 2)

        # Skip if under minimum size threshold
        if ($sizeMB -lt $MinimumSizeMB) { continue }

        # Resolve quotas: mailbox-level overrides if UseDatabaseQuotaDefaults is false
        $prohibitSendBytes    = Convert-QuotaToBytes -QuotaString $mbx.ProhibitSendQuota.ToString()
        $prohibitSendRecvBytes = Convert-QuotaToBytes -QuotaString $mbx.ProhibitSendReceiveQuota.ToString()
        $issueWarningBytes    = Convert-QuotaToBytes -QuotaString $mbx.IssueWarningQuota.ToString()

        # Calculate usage % (against ProhibitSendReceive if set, otherwise ProhibitSend)
        $quotaForCalc = if ($prohibitSendRecvBytes) { $prohibitSendRecvBytes } else { $prohibitSendBytes }
        $usagePercent = if ($quotaForCalc -and $quotaForCalc -gt 0) {
            [math]::Round(($sizeBytes / $quotaForCalc) * 100, 1)
        } else { $null }

        # Categorize quota status
        $quotaStatus = if (-not $quotaForCalc) {
            'Unlimited'
        } elseif ($prohibitSendRecvBytes -and $sizeBytes -ge $prohibitSendRecvBytes) {
            'SendReceiveDisabled'
        } elseif ($prohibitSendBytes -and $sizeBytes -ge $prohibitSendBytes) {
            'SendDisabled'
        } elseif ($issueWarningBytes -and $sizeBytes -ge $issueWarningBytes) {
            'Warning'
        } else {
            'Healthy'
        }

        $results.Add([PSCustomObject]@{
            DisplayName            = $mbx.DisplayName
            UserPrincipalName      = $mbx.UserPrincipalName
            PrimarySmtpAddress     = $mbx.PrimarySmtpAddress
            MailboxType            = $mbx.RecipientTypeDetails
            TotalSizeMB            = $sizeMB
            TotalSizeGB            = $sizeGB
            ItemCount              = $stats.ItemCount
            DeletedItemsMB         = $deletedMB
            DeletedItemCount       = $stats.DeletedItemCount
            QuotaUsagePercent      = $usagePercent
            QuotaStatus            = $quotaStatus
            IssueWarningQuotaGB    = if ($issueWarningBytes) { [math]::Round($issueWarningBytes / 1GB, 2) } else { 'Unlimited' }
            ProhibitSendQuotaGB    = if ($prohibitSendBytes) { [math]::Round($prohibitSendBytes / 1GB, 2) } else { 'Unlimited' }
            ProhibitSendRecvQuotaGB = if ($prohibitSendRecvBytes) { [math]::Round($prohibitSendRecvBytes / 1GB, 2) } else { 'Unlimited' }
            UseDatabaseQuotaDefaults = $mbx.UseDatabaseQuotaDefaults
            LastLogonTime          = $stats.LastLogonTime
            LastUserActionTime     = $stats.LastUserActionTime
            ArchiveStatus          = $mbx.ArchiveStatus
            LitigationHoldEnabled  = $mbx.LitigationHoldEnabled
        })
    }
    catch {
        Write-Warning "Failed to get statistics for $($mbx.PrimarySmtpAddress): $_"
        $results.Add([PSCustomObject]@{
            DisplayName            = $mbx.DisplayName
            UserPrincipalName      = $mbx.UserPrincipalName
            PrimarySmtpAddress     = $mbx.PrimarySmtpAddress
            MailboxType            = $mbx.RecipientTypeDetails
            TotalSizeMB            = $null
            TotalSizeGB            = $null
            ItemCount              = $null
            DeletedItemsMB         = $null
            DeletedItemCount       = $null
            QuotaUsagePercent      = $null
            QuotaStatus            = 'StatsRetrievalFailed'
            IssueWarningQuotaGB    = $null
            ProhibitSendQuotaGB    = $null
            ProhibitSendRecvQuotaGB = $null
            UseDatabaseQuotaDefaults = $null
            LastLogonTime          = $null
            LastUserActionTime     = $null
            ArchiveStatus          = $null
            LitigationHoldEnabled  = $null
        })
    }
}
Write-Progress -Activity "Gathering statistics" -Completed

# ---------- Export ----------
$reportCsv = Join-Path $LogPath "MailboxSizeReport_$timestamp.csv"
$results | Sort-Object TotalSizeGB -Descending | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8

# ---------- Summary ----------
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host ("Total mailboxes reported:    {0}" -f $results.Count)

$totalGB = [math]::Round((($results | Measure-Object -Property TotalSizeGB -Sum).Sum), 2)
Write-Host ("Combined storage usage:      {0} GB" -f $totalGB)

Write-Host "`nBy quota status:" -ForegroundColor Yellow
$results | Group-Object QuotaStatus | Sort-Object Count -Descending | ForEach-Object {
    Write-Host ("  {0,-22} {1}" -f $_.Name, $_.Count)
}

Write-Host "`nBy mailbox type:" -ForegroundColor Yellow
$results | Group-Object MailboxType | Sort-Object Count -Descending | ForEach-Object {
    Write-Host ("  {0,-22} {1}" -f $_.Name, $_.Count)
}

Write-Host "`nTop 10 largest mailboxes:" -ForegroundColor Yellow
$results | Sort-Object TotalSizeGB -Descending | Select-Object -First 10 | ForEach-Object {
    Write-Host ("  {0,-45} {1,8} GB  ({2})" -f $_.PrimarySmtpAddress, $_.TotalSizeGB, $_.QuotaStatus)
}

Write-Host "`nReport exported: $reportCsv" -ForegroundColor Green

Stop-Transcript | Out-Null