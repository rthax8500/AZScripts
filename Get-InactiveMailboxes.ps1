<#
.SYNOPSIS
    Identifies inactive mailboxes in Exchange Online based on last user activity.

.DESCRIPTION
    Read-only audit script. Reports mailboxes whose LastUserActionTime (or LastLogonTime
    fallback) exceeds a configurable inactivity threshold. Tiers results by inactivity
    age and includes mailbox size, license info, and account enabled state.

.PARAMETER InactiveDays
    Minimum number of days since last activity to include a mailbox in the report.
    Default is 90.

.PARAMETER MailboxType
    Filter by mailbox type. Valid values: All, UserMailbox, SharedMailbox, RoomMailbox,
    EquipmentMailbox. Default is UserMailbox (the typical license-reclamation target).

.PARAMETER IncludeNeverLoggedOn
    When set, also reports mailboxes with no recorded activity at all (null LastUserActionTime
    AND null LastLogonTime). Useful for finding provisioned-but-never-used accounts.

.PARAMETER LogPath
    Folder where the run log and report CSV will be written. Defaults to script directory.

.EXAMPLE
    .\Get-InactiveMailboxes.ps1

.EXAMPLE
    .\Get-InactiveMailboxes.ps1 -InactiveDays 180 -IncludeNeverLoggedOn

.EXAMPLE
    .\Get-InactiveMailboxes.ps1 -MailboxType SharedMailbox -InactiveDays 365
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [int]$InactiveDays = 90,

    [Parameter(Mandatory = $false)]
    [ValidateSet('All', 'UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox')]
    [string]$MailboxType = 'UserMailbox',

    [Parameter(Mandatory = $false)]
    [switch]$IncludeNeverLoggedOn,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "InactiveMailboxes_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

$threshold = (Get-Date).AddDays(-$InactiveDays)

Write-Host "=== Inactive Mailbox Report ===" -ForegroundColor Cyan
Write-Host "Mailbox type filter:    $MailboxType"
Write-Host "Inactivity threshold:   $InactiveDays days (older than $($threshold.ToString('yyyy-MM-dd')))"
Write-Host "Include never-logged-on: $IncludeNeverLoggedOn"
Write-Host "Log file:               $logFile`n"

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
function Convert-ExchangeSizeToBytes {
    param([string]$SizeString)
    if ([string]::IsNullOrWhiteSpace($SizeString)) { return 0 }
    if ($SizeString -match '\(([\d,]+) bytes\)') { return [int64]($Matches[1] -replace ',', '') }
    return 0
}

# ---------- Helper: tier classification ----------
function Get-InactivityTier {
    param([datetime]$LastActivity)
    $daysInactive = (New-TimeSpan -Start $LastActivity -End (Get-Date)).Days
    switch ($daysInactive) {
        { $_ -ge 365 } { return '365+ days' }
        { $_ -ge 180 } { return '180-364 days' }
        { $_ -ge 90  } { return '90-179 days'  }
        { $_ -ge 60  } { return '60-89 days'   }
        { $_ -ge 30  } { return '30-59 days'   }
        default        { return 'under 30 days' }
    }
}

# ---------- Get mailbox list ----------
Write-Host "Retrieving mailbox list..." -ForegroundColor Yellow
try {
    if ($MailboxType -eq 'All') {
        $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
    } else {
        $mailboxes = Get-Mailbox -RecipientTypeDetails $MailboxType -ResultSize Unlimited -ErrorAction Stop
    }
    Write-Host "Found $($mailboxes.Count) mailbox(es) of type $MailboxType.`n" -ForegroundColor Green
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
    Write-Progress -Activity "Checking activity" -Status $mbx.PrimarySmtpAddress -PercentComplete (($i / $mailboxes.Count) * 100)

    try {
        $stats = Get-MailboxStatistics -Identity $mbx.Identity -ErrorAction Stop

        # Use the most accurate signal available
        $lastActivity = if ($stats.LastUserActionTime) { $stats.LastUserActionTime }
                        elseif ($stats.LastLogonTime)  { $stats.LastLogonTime }
                        else                            { $null }

        $signalSource = if ($stats.LastUserActionTime) { 'LastUserActionTime' }
                        elseif ($stats.LastLogonTime)  { 'LastLogonTime' }
                        else                            { 'NeverLoggedOn' }

        # Decide whether this mailbox qualifies
        $include = $false
        $tier    = ''
        $daysInactive = $null

        if ($null -eq $lastActivity) {
            if ($IncludeNeverLoggedOn) {
                $include = $true
                $tier    = 'NeverLoggedOn'
            }
        }
        elseif ($lastActivity -lt $threshold) {
            $include = $true
            $tier    = Get-InactivityTier -LastActivity $lastActivity
            $daysInactive = (New-TimeSpan -Start $lastActivity -End (Get-Date)).Days
        }

        if (-not $include) { continue }

        $sizeBytes = Convert-ExchangeSizeToBytes -SizeString $stats.TotalItemSize.ToString()
        $sizeMB    = [math]::Round($sizeBytes / 1MB, 2)
        $sizeGB    = [math]::Round($sizeBytes / 1GB, 2)

        $results.Add([PSCustomObject]@{
            DisplayName            = $mbx.DisplayName
            UserPrincipalName      = $mbx.UserPrincipalName
            PrimarySmtpAddress     = $mbx.PrimarySmtpAddress
            MailboxType            = $mbx.RecipientTypeDetails
            AccountDisabled        = $mbx.AccountDisabled
            InactivityTier         = $tier
            DaysInactive           = $daysInactive
            LastActivity           = $lastActivity
            ActivitySignalSource   = $signalSource
            ItemCount              = $stats.ItemCount
            TotalSizeMB            = $sizeMB
            TotalSizeGB            = $sizeGB
            ArchiveStatus          = $mbx.ArchiveStatus
            LitigationHoldEnabled  = $mbx.LitigationHoldEnabled
            HiddenFromAddressList  = $mbx.HiddenFromAddressListsEnabled
            WhenMailboxCreated     = $mbx.WhenMailboxCreated
        })
    }
    catch {
        Write-Warning "Failed to get stats for $($mbx.PrimarySmtpAddress): $_"
    }
}
Write-Progress -Activity "Checking activity" -Completed

# ---------- Export ----------
$reportCsv = Join-Path $LogPath "InactiveMailboxes_$timestamp.csv"

if ($results.Count -eq 0) {
    Write-Host "`nNo inactive mailboxes found matching the criteria." -ForegroundColor Green
    "DisplayName,UserPrincipalName,PrimarySmtpAddress,MailboxType,AccountDisabled,InactivityTier,DaysInactive,LastActivity,ActivitySignalSource,ItemCount,TotalSizeMB,TotalSizeGB,ArchiveStatus,LitigationHoldEnabled,HiddenFromAddressList,WhenMailboxCreated" |
        Out-File -FilePath $reportCsv -Encoding UTF8
    Stop-Transcript | Out-Null
    return
}

$results | Sort-Object DaysInactive -Descending | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8

# ---------- Summary ----------
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host ("Total inactive mailboxes:    {0}" -f $results.Count)

$totalGB = [math]::Round((($results | Measure-Object -Property TotalSizeGB -Sum).Sum), 2)
Write-Host ("Combined storage (reclaimable): {0} GB" -f $totalGB)

Write-Host "`nBy inactivity tier:" -ForegroundColor Yellow
# Sort tiers in logical order rather than alphabetical
$tierOrder = @('NeverLoggedOn', '365+ days', '180-364 days', '90-179 days', '60-89 days', '30-59 days', 'under 30 days')
$tierOrder | ForEach-Object {
    $count = ($results | Where-Object { $_.InactivityTier -eq $_ }).Count
    if (($results | Where-Object InactivityTier -eq $_).Count -gt 0) {
        Write-Host ("  {0,-18} {1}" -f $_, ($results | Where-Object InactivityTier -eq $_).Count)
    }
}

Write-Host "`nBy account state:" -ForegroundColor Yellow
$results | Group-Object AccountDisabled | ForEach-Object {
    $label = if ($_.Name -eq 'True') { 'Disabled' } else { 'Enabled ' }
    Write-Host ("  {0}  {1}" -f $label, $_.Count)
}

Write-Host "`nReclamation flags:" -ForegroundColor Yellow
$onHold     = ($results | Where-Object LitigationHoldEnabled -eq $true).Count
$archived   = ($results | Where-Object ArchiveStatus -eq 'Active').Count
$hidden     = ($results | Where-Object HiddenFromAddressList -eq $true).Count
Write-Host ("  On litigation hold:     {0}  (DO NOT delete without legal review)" -f $onHold)
Write-Host ("  Has archive enabled:    {0}" -f $archived)
Write-Host ("  Hidden from GAL:        {0}" -f $hidden)

Write-Host "`nTop 10 most inactive:" -ForegroundColor Yellow
$results | Sort-Object DaysInactive -Descending | Select-Object -First 10 | ForEach-Object {
    $days = if ($_.DaysInactive) { "$($_.DaysInactive)d" } else { 'never' }
    Write-Host ("  {0,-45} {1,8}  ({2} GB)" -f $_.PrimarySmtpAddress, $days, $_.TotalSizeGB)
}

Write-Host "`nReport exported: $reportCsv" -ForegroundColor Green

Stop-Transcript | Out-Null