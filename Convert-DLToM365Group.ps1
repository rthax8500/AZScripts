<#
.SYNOPSIS
    Converts an eligible Distribution List to a Microsoft 365 Group with pre-flight checks and rollback snapshot.

.DESCRIPTION
    Wraps Microsoft's native Upgrade-DistributionGroup cmdlet with:
      - Eligibility validation before attempting upgrade
      - Membership snapshot exported to CSV for rollback purposes
      - Transcript logging
      - Post-upgrade verification

    Note: This is a one-way operation. The original DL is deleted and replaced
    with a new M365 Group bearing the same SMTP address.

.PARAMETER DistributionList
    The DL's primary SMTP address, alias, or DisplayName.

.PARAMETER LogPath
    Folder where the run log, rollback CSV, and verification CSV will be written.

.PARAMETER Force
    Skip the interactive confirmation prompt.

.EXAMPLE
    .\Convert-DLToM365Group.ps1 -DistributionList "ProjectAlpha@contoso.com" -WhatIf

.EXAMPLE
    .\Convert-DLToM365Group.ps1 -DistributionList "ProjectAlpha@contoso.com"
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$DistributionList,    # <-- REPLACE at runtime: e.g. "OldDL@yourdomain.com"

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path),

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "ConvertDLToM365Group_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host "=== Convert DL to Microsoft 365 Group ===" -ForegroundColor Cyan
Write-Host "DL:       $DistributionList"
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

# ---------- Validate DL exists and is eligible ----------
Write-Host "`n--- Pre-flight checks ---" -ForegroundColor Cyan

try {
    $dl = Get-DistributionGroup -Identity $DistributionList -ErrorAction Stop
}
catch {
    Write-Error "Distribution list '$DistributionList' not found: $_"
    Stop-Transcript | Out-Null
    return
}

$blockers = [System.Collections.Generic.List[string]]::new()

# Check 1: Must be a Distribution List, not a Mail-Enabled Security Group
if ($dl.RecipientTypeDetails -ne 'MailUniversalDistributionGroup') {
    $blockers.Add("Not a standard Distribution List (type: $($dl.RecipientTypeDetails)). Mail-enabled security groups cannot be upgraded.")
}
else {
    Write-Host "  [PASS] Group type is MailUniversalDistributionGroup"
}

# Check 2: Must be cloud-managed (not synced from on-prem AD)
if ($dl.IsDirSynced -eq $true) {
    $blockers.Add("DL is synced from on-prem AD (IsDirSynced = True). Only cloud-managed DLs can be upgraded.")
}
else {
    Write-Host "  [PASS] DL is cloud-managed"
}

# Check 3: Must have an owner
if (-not $dl.ManagedBy -or $dl.ManagedBy.Count -eq 0) {
    $blockers.Add("DL has no owner (ManagedBy is empty). Assign an owner before upgrading.")
}
else {
    Write-Host "  [PASS] DL has $($dl.ManagedBy.Count) owner(s)"
}

# Check 4: Must not be nested inside another DL
try {
    $parentDLs = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop |
        Where-Object {
            (Get-DistributionGroupMember -Identity $_.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue).PrimarySmtpAddress -contains $dl.PrimarySmtpAddress
        }
    if ($parentDLs) {
        $blockers.Add("DL is nested inside parent DL(s): $($parentDLs.DisplayName -join ', '). Remove from parent(s) before upgrading.")
    }
    else {
        Write-Host "  [PASS] DL is not nested inside another DL"
    }
}
catch {
    Write-Warning "  [SKIP] Could not check for parent DLs: $_"
}

# Check 5: Members must all be user mailboxes
try {
    $members = Get-DistributionGroupMember -Identity $dl.Identity -ResultSize Unlimited -ErrorAction Stop
    $invalidMembers = $members | Where-Object {
        $_.RecipientTypeDetails -notin @('UserMailbox', 'SharedMailbox', 'GuestMailUser')
    }
    if ($invalidMembers) {
        $blockers.Add("DL contains $($invalidMembers.Count) member(s) of unsupported types: $(($invalidMembers.RecipientTypeDetails | Sort-Object -Unique) -join ', ')")
    }
    else {
        Write-Host "  [PASS] All $($members.Count) member(s) are eligible types"
    }
}
catch {
    Write-Warning "  [SKIP] Could not enumerate members: $_"
    $members = @()
}

# ---------- Block on failures ----------
if ($blockers.Count -gt 0) {
    Write-Host "`n=== Eligibility blockers found ===" -ForegroundColor Red
    foreach ($b in $blockers) { Write-Host "  - $b" -ForegroundColor Red }
    Write-Host "`nResolve the above before attempting upgrade. Exiting." -ForegroundColor Yellow
    Stop-Transcript | Out-Null
    return
}

Write-Host "`nAll eligibility checks passed.`n" -ForegroundColor Green

# ---------- Membership snapshot for rollback ----------
$rollbackCsv = Join-Path $LogPath "ConvertDLToM365Group_Rollback_$($dl.Alias)_$timestamp.csv"

$snapshot = $members | ForEach-Object {
    [PSCustomObject]@{
        UserPrincipalName    = if ($_.WindowsLiveID) { $_.WindowsLiveID } else { $_.PrimarySmtpAddress }
        DisplayName          = $_.DisplayName
        PrimarySmtpAddress   = $_.PrimarySmtpAddress
        RecipientTypeDetails = $_.RecipientTypeDetails
    }
}

if ($snapshot) {
    $snapshot | Export-Csv -Path $rollbackCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Rollback snapshot saved: $rollbackCsv" -ForegroundColor Green
}
else {
    Write-Warning "No members to snapshot."
}

# Capture key DL properties for rollback reference
$dlSnapshot = [PSCustomObject]@{
    DisplayName        = $dl.DisplayName
    Alias              = $dl.Alias
    PrimarySmtpAddress = $dl.PrimarySmtpAddress
    ManagedBy          = ($dl.ManagedBy -join '; ')
    EmailAddresses     = ($dl.EmailAddresses -join '; ')
    WhenCreated        = $dl.WhenCreated
}
$dlPropsCsv = Join-Path $LogPath "ConvertDLToM365Group_DLProperties_$($dl.Alias)_$timestamp.csv"
$dlSnapshot | Export-Csv -Path $dlPropsCsv -NoTypeInformation -Encoding UTF8
Write-Host "DL properties saved:    $dlPropsCsv`n" -ForegroundColor Green

# ---------- Confirmation gate ----------
Write-Host "Ready to convert:" -ForegroundColor Yellow
Write-Host "  DL:           $($dl.DisplayName) [$($dl.PrimarySmtpAddress)]"
Write-Host "  Member count: $($members.Count)"
Write-Host "  This will DELETE the DL and create a new M365 Group with the same SMTP." -ForegroundColor Yellow
Write-Host "  This action is NOT reversible through the UI.`n" -ForegroundColor Yellow

if (-not $Force -and -not $WhatIfPreference) {
    $confirm = Read-Host "Type 'CONVERT' to proceed (anything else cancels)"
    if ($confirm -ne 'CONVERT') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        Stop-Transcript | Out-Null
        return
    }
}

# ---------- Execute upgrade ----------
if ($PSCmdlet.ShouldProcess($dl.PrimarySmtpAddress, "Upgrade DL to M365 Group")) {
    try {
        Write-Host "`nRunning Upgrade-DistributionGroup..." -ForegroundColor Yellow
        Upgrade-DistributionGroup -DlIdentities $dl.PrimarySmtpAddress -ErrorAction Stop
        Write-Host "Upgrade command submitted." -ForegroundColor Green
    }
    catch {
        Write-Error "Upgrade-DistributionGroup failed: $_"
        Stop-Transcript | Out-Null
        return
    }
}
else {
    Write-Host "`n[WhatIf] Would run: Upgrade-DistributionGroup -DlIdentities $($dl.PrimarySmtpAddress)" -ForegroundColor Cyan
    Stop-Transcript | Out-Null
    return
}

# ---------- Post-upgrade verification ----------
Write-Host "`n--- Post-upgrade verification ---" -ForegroundColor Cyan
Write-Host "Waiting 30 seconds for replication..." -ForegroundColor Yellow
Start-Sleep -Seconds 30

try {
    $newGroup = Get-UnifiedGroup -Identity $dl.PrimarySmtpAddress -ErrorAction Stop
    Write-Host "  [VERIFIED] New M365 Group exists: $($newGroup.DisplayName)" -ForegroundColor Green

    $newMembers = Get-UnifiedGroupLinks -Identity $newGroup.Identity -LinkType Members -ResultSize Unlimited
    Write-Host "  [VERIFIED] New group has $($newMembers.Count) member(s)" -ForegroundColor Green

    if ($newMembers.Count -ne $members.Count) {
        Write-Warning "  Member count differs from snapshot ($($members.Count) -> $($newMembers.Count)). Review the rollback CSV."
    }

    $verificationCsv = Join-Path $LogPath "ConvertDLToM365Group_Verification_$($dl.Alias)_$timestamp.csv"
    $newMembers | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails |
        Export-Csv -Path $verificationCsv -NoTypeInformation -Encoding UTF8
    Write-Host "  Verification CSV: $verificationCsv" -ForegroundColor Green
}
catch {
    Write-Warning "Could not verify new M365 Group yet — replication may take longer. Re-check in a few minutes with: Get-UnifiedGroup -Identity '$($dl.PrimarySmtpAddress)'"
}

Write-Host "`n=== Conversion Complete ===" -ForegroundColor Cyan
Stop-Transcript | Out-Null