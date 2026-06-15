<#
.SYNOPSIS
    Identifies stale (inactive) Entra ID user accounts based on sign-in activity.

.DESCRIPTION
    Read-only audit script. Reports users whose last sign-in (interactive or non-interactive)
    exceeds a configurable inactivity threshold. Tiers results by age, separately flags
    privileged accounts, and highlights divergence between interactive and non-interactive
    sign-in dates.

    Requires Entra ID P1 or P2 licensing for sign-in activity data via Graph.

.PARAMETER InactiveDays
    Minimum days since last sign-in to include in the report. Default is 90.

.PARAMETER UseNonInteractive
    When set, uses non-interactive sign-in (the more lenient signal) for the staleness
    threshold. Default uses interactive sign-in (the security-relevant signal).

.PARAMETER IncludeDisabledAccounts
    When set, includes accounts that are already disabled.

.PARAMETER IncludeGuests
    When set, includes guest (B2B) users. Default is members only.

.PARAMETER LogPath
    Folder where the run log and report CSV will be written. Defaults to script directory.

.EXAMPLE
    .\Get-StaleUsers.ps1

.EXAMPLE
    .\Get-StaleUsers.ps1 -InactiveDays 180 -IncludeGuests

.EXAMPLE
    .\Get-StaleUsers.ps1 -InactiveDays 60 -UseNonInteractive
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [int]$InactiveDays = 90,

    [Parameter(Mandatory = $false)]
    [switch]$UseNonInteractive,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDisabledAccounts,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeGuests,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "StaleUsers_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

$threshold = (Get-Date).AddDays(-$InactiveDays).ToUniversalTime()
$signalLabel = if ($UseNonInteractive) { 'non-interactive' } else { 'interactive' }

Write-Host "=== Stale User Account Report ===" -ForegroundColor Cyan
Write-Host "Inactivity threshold:      $InactiveDays days (older than $($threshold.ToString('yyyy-MM-dd')) UTC)"
Write-Host "Staleness signal:          $signalLabel sign-in"
Write-Host "Include disabled accounts: $IncludeDisabledAccounts"
Write-Host "Include guest users:       $IncludeGuests"
Write-Host "Log file:                  $logFile`n"

# ---------- Module check ----------
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.DirectoryManagement'
)
$missing = $requiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
if ($missing) {
    Write-Error "Missing required modules: $($missing -join ', ')"
    Write-Host "Install with: Install-Module $($missing -join ', ') -Scope CurrentUser" -ForegroundColor Yellow
    Stop-Transcript | Out-Null
    return
}

# ---------- Connect to Microsoft Graph ----------
$requiredScopes = @(
    'AuditLog.Read.All',
    'User.Read.All',
    'Directory.Read.All'
)

try {
    $context = Get-MgContext -ErrorAction SilentlyContinue
    $needsConnect = $false

    if (-not $context) {
        $needsConnect = $true
    } else {
        $missingScopes = $requiredScopes | Where-Object { $context.Scopes -notcontains $_ }
        if ($missingScopes) {
            Write-Host "Reconnecting to add missing scopes: $($missingScopes -join ', ')" -ForegroundColor Yellow
            Disconnect-MgGraph | Out-Null
            $needsConnect = $true
        }
    }

    if ($needsConnect) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
    } else {
        Write-Host "Already connected to Microsoft Graph as $($context.Account)." -ForegroundColor Green
    }
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    Stop-Transcript | Out-Null
    return
}

# ---------- Helper: tier classification ----------
function Get-StalenessTier {
    param([datetime]$LastSignIn)
    $days = (New-TimeSpan -Start $LastSignIn -End (Get-Date)).Days
    switch ($days) {
        { $_ -ge 365 } { return '365+ days' }
        { $_ -ge 180 } { return '180-364 days' }
        { $_ -ge 90  } { return '90-179 days'  }
        { $_ -ge 60  } { return '60-89 days'   }
        { $_ -ge 30  } { return '30-59 days'   }
        default        { return 'under 30 days' }
    }
}

# ---------- Get directory roles for privileged-account flagging ----------
Write-Host "Building privileged-user list..." -ForegroundColor Yellow
$privilegedUserIds = [System.Collections.Generic.HashSet[string]]::new()
try {
    $roles = Get-MgDirectoryRole -All -ErrorAction Stop
    foreach ($role in $roles) {
        try {
            $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction Stop
            foreach ($m in $members) {
                if ($m.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user') {
                    $null = $privilegedUserIds.Add($m.Id)
                }
            }
        } catch {
            Write-Warning "  Could not enumerate members of role $($role.DisplayName): $_"
        }
    }
    Write-Host "Found $($privilegedUserIds.Count) privileged user(s).`n" -ForegroundColor Green
}
catch {
    Write-Warning "Could not enumerate directory roles; IsPrivileged flag may be incomplete: $_"
}

# ---------- Retrieve user data ----------
Write-Host "Retrieving users with sign-in activity..." -ForegroundColor Yellow
Write-Host "(This can take several minutes on large tenants.)`n" -ForegroundColor Yellow

# signInActivity must be requested explicitly via -Property
$selectProps = @(
    'Id',
    'UserPrincipalName',
    'DisplayName',
    'AccountEnabled',
    'UserType',
    'CreatedDateTime',
    'Mail',
    'SignInActivity'
)

try {
    $allUsers = Get-MgUser -All -Property $selectProps -ErrorAction Stop
}
catch {
    Write-Error "Failed to retrieve users: $_"
    Stop-Transcript | Out-Null
    return
}

# License-check sniff: if EVERY user has null SignInActivity, the tenant likely lacks Entra ID P1/P2.
$signInDataPresent = ($allUsers | Where-Object { $_.SignInActivity }).Count
if ($signInDataPresent -eq 0 -and $allUsers.Count -gt 0) {
    Write-Warning "No sign-in activity data is available. This typically means the tenant does not have Entra ID P1 or P2 licensing."
    Write-Host "Without P1/P2, this script cannot produce meaningful results. Consider using mailbox last-logon as a fallback (Get-InactiveMailboxes.ps1)." -ForegroundColor Yellow
    Stop-Transcript | Out-Null
    return
}

Write-Host "Retrieved $($allUsers.Count) user(s). $signInDataPresent have sign-in activity data.`n" -ForegroundColor Green

# ---------- Process ----------
$results = [System.Collections.Generic.List[object]]::new()

foreach ($u in $allUsers) {

    # Apply user-type filter
    if (-not $IncludeGuests -and $u.UserType -eq 'Guest') { continue }
    if (-not $IncludeDisabledAccounts -and -not $u.AccountEnabled) { continue }

    $lastInteractive    = $u.SignInActivity.LastSignInDateTime
    $lastNonInteractive = $u.SignInActivity.LastNonInteractiveSignInDateTime

    # Pick the staleness signal
    $primarySignIn = if ($UseNonInteractive) {
        if ($lastNonInteractive) { $lastNonInteractive } else { $lastInteractive }
    } else {
        $lastInteractive
    }

    # Decide inclusion
    $include = $false
    if ($null -eq $primarySignIn) {
        # Never signed in - include only if account is older than the threshold
        if ($u.CreatedDateTime -and $u.CreatedDateTime -lt $threshold) {
            $include = $true
        }
    }
    elseif ($primarySignIn -lt $threshold) {
        $include = $true
    }

    if (-not $include) { continue }

    $tier = if ($null -eq $primarySignIn) { 'NeverSignedIn' } else { Get-StalenessTier -LastSignIn $primarySignIn }
    $daysInactive = if ($primarySignIn) { (New-TimeSpan -Start $primarySignIn -End (Get-Date)).Days } else { $null }

    # Flag interactive vs non-interactive divergence (only relevant when using interactive signal)
    $hasNonInteractiveActivity = $false
    if (-not $UseNonInteractive -and $lastNonInteractive -and $lastNonInteractive -ge $threshold) {
        $hasNonInteractiveActivity = $true
    }

    $results.Add([PSCustomObject]@{
        DisplayName               = $u.DisplayName
        UserPrincipalName         = $u.UserPrincipalName
        UserType                  = $u.UserType
        AccountEnabled            = $u.AccountEnabled
        IsPrivileged              = $privilegedUserIds.Contains($u.Id)
        StalenessTier             = $tier
        DaysInactive              = $daysInactive
        LastInteractiveSignIn     = $lastInteractive
        LastNonInteractiveSignIn  = $lastNonInteractive
        HasRecentNonInteractive   = $hasNonInteractiveActivity
        AccountCreated            = $u.CreatedDateTime
        Mail                      = $u.Mail
    })
}

# ---------- Export ----------
$reportCsv = Join-Path $LogPath "StaleUsers_$timestamp.csv"

if ($results.Count -eq 0) {
    Write-Host "`nNo stale users found matching the criteria." -ForegroundColor Green
    "DisplayName,UserPrincipalName,UserType,AccountEnabled,IsPrivileged,StalenessTier,DaysInactive,LastInteractiveSignIn,LastNonInteractiveSignIn,HasRecentNonInteractive,AccountCreated,Mail" |
        Out-File -FilePath $reportCsv -Encoding UTF8
    Stop-Transcript | Out-Null
    return
}

$results | Sort-Object IsPrivileged -Descending, DaysInactive -Descending |
    Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8

# ---------- Summary ----------
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host ("Total stale accounts:        {0}" -f $results.Count)

$privCount  = ($results | Where-Object IsPrivileged -eq $true).Count
$divergent  = ($results | Where-Object HasRecentNonInteractive -eq $true).Count
$disabled   = ($results | Where-Object AccountEnabled -eq $false).Count

if ($privCount -gt 0) {
    Write-Host ("Privileged accounts stale:   {0}  *** REVIEW IMMEDIATELY ***" -f $privCount) -ForegroundColor Red
}
if ($divergent -gt 0) {
    Write-Host ("With recent non-interactive: {0}  (treat with care - tokens still active)" -f $divergent) -ForegroundColor Yellow
}
Write-Host ("Already disabled:            {0}" -f $disabled)

Write-Host "`nBy staleness tier:" -ForegroundColor Yellow
$tierOrder = @('NeverSignedIn', '365+ days', '180-364 days', '90-179 days', '60-89 days', '30-59 days')
foreach ($t in $tierOrder) {
    $c = ($results | Where-Object StalenessTier -eq $t).Count
    if ($c -gt 0) { Write-Host ("  {0,-18} {1}" -f $t, $c) }
}

if ($privCount -gt 0) {
    Write-Host "`nStale privileged accounts (immediate priority):" -ForegroundColor Red
    $results | Where-Object IsPrivileged -eq $true | Sort-Object DaysInactive -Descending | ForEach-Object {
        $days = if ($_.DaysInactive) { "$($_.DaysInactive)d" } else { 'never' }
        Write-Host ("  {0,-45} {1,8}" -f $_.UserPrincipalName, $days) -ForegroundColor Red
    }
}

Write-Host "`nReport exported: $reportCsv" -ForegroundColor Green

Stop-Transcript | Out-Null