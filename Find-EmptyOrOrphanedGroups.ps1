<#
.SYNOPSIS
    Identifies empty groups, owner-less groups, and groups with orphaned/disabled owners across Exchange Online.

.DESCRIPTION
    Read-only audit script. Scans distribution lists, mail-enabled security groups,
    and Microsoft 365 groups in the tenant and reports any that are:
      - Empty (zero members)
      - No-owner (no manager/owner assigned)
      - Orphaned (owner account no longer exists)
      - DisabledOwner (owner exists but is disabled)

    Outputs a categorized CSV plus a transcript log. Makes no changes.

.PARAMETER LogPath
    Folder where the run log and report CSV will be written. Defaults to script directory.

.PARAMETER GroupTypeFilter
    Optional. Restrict scan to one type. Valid values: All, DistributionList,
    MailSecurityGroup, M365Group. Default is All.

.EXAMPLE
    .\Find-EmptyOrOrphanedGroups.ps1

.EXAMPLE
    .\Find-EmptyOrOrphanedGroups.ps1 -GroupTypeFilter M365Group
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path),

    [Parameter(Mandatory = $false)]
    [ValidateSet('All', 'DistributionList', 'MailSecurityGroup', 'M365Group')]
    [string]$GroupTypeFilter = 'All'
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "FindEmptyOrOrphanedGroups_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host "=== Find Empty / Orphaned Groups ===" -ForegroundColor Cyan
Write-Host "Filter:   $GroupTypeFilter"
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

# ---------- Build mailbox cache for owner validation ----------
# Pulling once is far faster than calling Get-Recipient per owner.
Write-Host "Caching recipient list for owner validation (this can take a while)..." -ForegroundColor Yellow
$recipientCache = @{}
try {
    Get-Recipient -ResultSize Unlimited -ErrorAction Stop | ForEach-Object {
        if ($_.PrimarySmtpAddress) {
            $recipientCache[$_.PrimarySmtpAddress.ToString().ToLower()] = $_
        }
    }
    Write-Host "Cached $($recipientCache.Count) recipients.`n" -ForegroundColor Green
}
catch {
    Write-Warning "Recipient cache failed; owner validation may be incomplete: $_"
}

# ---------- Helper: evaluate a single group ----------
function Test-GroupHealth {
    param(
        [object]$Group,
        [string]$Type,
        [array]$Members,
        [array]$OwnerSmtpList
    )

    $issues = @()

    if (-not $Members -or $Members.Count -eq 0) {
        $issues += 'Empty'
    }

    if (-not $OwnerSmtpList -or $OwnerSmtpList.Count -eq 0) {
        $issues += 'NoOwner'
    }
    else {
        foreach ($ownerSmtp in $OwnerSmtpList) {
            $key = $ownerSmtp.ToString().ToLower()
            if (-not $recipientCache.ContainsKey($key)) {
                $issues += 'OrphanedOwner'
                break
            }
            else {
                # Disabled detection: UserMailbox with no logon -> not reliable here,
                # but recipient-level disabled accounts return as 'DisabledUser'
                $rec = $recipientCache[$key]
                if ($rec.RecipientTypeDetails -eq 'DisabledUser') {
                    $issues += 'DisabledOwner'
                    break
                }
            }
        }
    }

    if ($issues.Count -eq 0) { return $null }

    [PSCustomObject]@{
        GroupName          = $Group.DisplayName
        PrimarySmtpAddress = $Group.PrimarySmtpAddress
        GroupType          = $Type
        Issue              = ($issues | Sort-Object -Unique) -join ', '
        MemberCount        = $Members.Count
        OwnerCount         = $OwnerSmtpList.Count
        OwnerDetail        = if ($OwnerSmtpList) { $OwnerSmtpList -join '; ' } else { '(none)' }
        WhenCreated        = $Group.WhenCreated
    }
}

# ---------- Scan: Distribution Lists & Mail-Enabled Security Groups ----------
$findings = [System.Collections.Generic.List[object]]::new()

if ($GroupTypeFilter -in @('All', 'DistributionList', 'MailSecurityGroup')) {
    Write-Host "Scanning distribution lists and mail-enabled security groups..." -ForegroundColor Yellow

    try {
        $dgFilter = switch ($GroupTypeFilter) {
            'DistributionList'   { "RecipientTypeDetails -eq 'MailUniversalDistributionGroup'" }
            'MailSecurityGroup'  { "RecipientTypeDetails -eq 'MailUniversalSecurityGroup'" }
            default              { $null }
        }

        $dgs = if ($dgFilter) {
            Get-DistributionGroup -ResultSize Unlimited -Filter $dgFilter -ErrorAction Stop
        } else {
            Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop
        }

        $i = 0
        foreach ($g in $dgs) {
            $i++
            Write-Progress -Activity "Scanning DLs / MESGs" -Status $g.DisplayName -PercentComplete (($i / $dgs.Count) * 100)

            $members = @(Get-DistributionGroupMember -Identity $g.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue)
            # ManagedBy is array of identities; resolve to SMTP via cache where possible
            $ownerSmtp = @()
            foreach ($m in $g.ManagedBy) {
                $found = $recipientCache.Values | Where-Object { $_.Identity -eq $m -or $_.DistinguishedName -eq $m } | Select-Object -First 1
                if ($found) { $ownerSmtp += $found.PrimarySmtpAddress.ToString() }
                else        { $ownerSmtp += $m.ToString() }   # keep raw value so OrphanedOwner detection still fires
            }

            $type = if ($g.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') { 'MailSecurityGroup' } else { 'DistributionList' }
            $finding = Test-GroupHealth -Group $g -Type $type -Members $members -OwnerSmtpList $ownerSmtp
            if ($finding) { $findings.Add($finding) }
        }
        Write-Progress -Activity "Scanning DLs / MESGs" -Completed
    }
    catch {
        Write-Warning "DL/MESG scan failed: $_"
    }
}

# ---------- Scan: M365 Groups ----------
if ($GroupTypeFilter -in @('All', 'M365Group')) {
    Write-Host "Scanning Microsoft 365 groups..." -ForegroundColor Yellow

    try {
        $ugs = Get-UnifiedGroup -ResultSize Unlimited -ErrorAction Stop
        $i = 0
        foreach ($g in $ugs) {
            $i++
            Write-Progress -Activity "Scanning M365 Groups" -Status $g.DisplayName -PercentComplete (($i / $ugs.Count) * 100)

            $members = @(Get-UnifiedGroupLinks -Identity $g.Identity -LinkType Members -ResultSize Unlimited -ErrorAction SilentlyContinue)
            $owners  = @(Get-UnifiedGroupLinks -Identity $g.Identity -LinkType Owners  -ResultSize Unlimited -ErrorAction SilentlyContinue)
            $ownerSmtp = $owners | Select-Object -ExpandProperty PrimarySmtpAddress

            $finding = Test-GroupHealth -Group $g -Type 'M365Group' -Members $members -OwnerSmtpList $ownerSmtp
            if ($finding) { $findings.Add($finding) }
        }
        Write-Progress -Activity "Scanning M365 Groups" -Completed
    }
    catch {
        Write-Warning "M365 Group scan failed: $_"
    }
}

# ---------- Output ----------
$reportCsv = Join-Path $LogPath "FindEmptyOrOrphanedGroups_Report_$timestamp.csv"

if ($findings.Count -eq 0) {
    Write-Host "`nNo problematic groups found. Tenant is clean!" -ForegroundColor Green
    "GroupName,PrimarySmtpAddress,GroupType,Issue,MemberCount,OwnerCount,OwnerDetail,WhenCreated" |
        Out-File -FilePath $reportCsv -Encoding UTF8
}
else {
    $findings | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8

    Write-Host "`n=== Summary ===" -ForegroundColor Cyan
    Write-Host ("Total problematic groups: {0}" -f $findings.Count)
    $findings | Group-Object Issue | Sort-Object Count -Descending | ForEach-Object {
        Write-Host ("  {0,-40} {1}" -f $_.Name, $_.Count)
    }
    Write-Host ""
    Write-Host ("By group type:")
    $findings | Group-Object GroupType | ForEach-Object {
        Write-Host ("  {0,-25} {1}" -f $_.Name, $_.Count)
    }
}

Write-Host "`nReport exported: $reportCsv" -ForegroundColor Green

Stop-Transcript | Out-Null