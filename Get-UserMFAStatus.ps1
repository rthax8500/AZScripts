<#
.SYNOPSIS
    Reports MFA registration and capability status for all users in Entra ID via Microsoft Graph.

.DESCRIPTION
    Read-only audit script. Pulls MFA registration details from the userRegistrationDetails
    Graph endpoint and exports a categorized CSV. Includes registered methods, default method,
    passwordless capability, and account enabled state.

    Requires Microsoft Graph PowerShell module. Run with an account that holds at least
    Global Reader, Security Reader, Authentication Admin, or Privileged Auth Admin role.

.PARAMETER UserType
    Filter by user type. Valid values: All, MfaRegistered, NotMfaRegistered. Default is All.

.PARAMETER IncludeDisabledAccounts
    When set, includes users whose Entra accounts are disabled. By default, disabled
    accounts are excluded since they don't represent active MFA gaps.

.PARAMETER LogPath
    Folder where the run log and report CSV will be written. Defaults to script directory.

.EXAMPLE
    .\Get-UserMFAStatus.ps1

.EXAMPLE
    .\Get-UserMFAStatus.ps1 -UserType NotMfaRegistered

.EXAMPLE
    .\Get-UserMFAStatus.ps1 -IncludeDisabledAccounts
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('All', 'MfaRegistered', 'NotMfaRegistered')]
    [string]$UserType = 'All',

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDisabledAccounts,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "UserMFAStatus_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host "=== User MFA Status Report ===" -ForegroundColor Cyan
Write-Host "Filter:           $UserType"
Write-Host "Include disabled: $IncludeDisabledAccounts"
Write-Host "Log file:         $logFile`n"

# ---------- Module check ----------
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Identity.SignIns',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Reports'
)
$missing = $requiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
if ($missing) {
    Write-Error "Missing required Microsoft.Graph modules: $($missing -join ', ')"
    Write-Host "Install with: Install-Module $($missing -join ', ') -Scope CurrentUser" -ForegroundColor Yellow
    Stop-Transcript | Out-Null
    return
}

# ---------- Connect to Microsoft Graph ----------
$requiredScopes = @(
    'AuditLog.Read.All',
    'UserAuthenticationMethod.Read.All',
    'User.Read.All'
)

try {
    $context = Get-MgContext -ErrorAction SilentlyContinue
    $needsConnect = $false

    if (-not $context) {
        $needsConnect = $true
    } else {
        # Verify all required scopes are granted on the current context
        $missingScopes = $requiredScopes | Where-Object { $context.Scopes -notcontains $_ }
        if ($missingScopes) {
            Write-Host "Existing Graph context is missing scopes: $($missingScopes -join ', '). Reconnecting..." -ForegroundColor Yellow
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

# ---------- Retrieve registration data ----------
Write-Host "`nRetrieving authentication method registration details..." -ForegroundColor Yellow
try {
    # The userRegistrationDetails endpoint is the canonical source for MFA status.
    # It's a "report" endpoint - data refreshes daily, so it can lag real-time changes by up to 24h.
    $registrationData = Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop
    Write-Host "Retrieved $($registrationData.Count) registration record(s)." -ForegroundColor Green
}
catch {
    Write-Error "Failed to retrieve registration data: $_"
    Stop-Transcript | Out-Null
    return
}

# ---------- Cross-reference with user accounts for AccountEnabled ----------
Write-Host "Building user account state cache..." -ForegroundColor Yellow
$userCache = @{}
try {
    Get-MgUser -All -Property Id, UserPrincipalName, AccountEnabled, UserType -ErrorAction Stop | ForEach-Object {
        $userCache[$_.Id] = $_
    }
    Write-Host "Cached $($userCache.Count) user account(s).`n" -ForegroundColor Green
}
catch {
    Write-Warning "Could not build user cache; AccountEnabled column may be incomplete: $_"
}

# ---------- Process records ----------
$results = [System.Collections.Generic.List[object]]::new()

foreach ($r in $registrationData) {

    $accountEnabled = $true
    $userType       = 'Unknown'
    if ($userCache.ContainsKey($r.Id)) {
        $accountEnabled = $userCache[$r.Id].AccountEnabled
        $userType       = $userCache[$r.Id].UserType
    }

    # Apply enabled-account filter
    if (-not $IncludeDisabledAccounts -and -not $accountEnabled) { continue }

    # Apply MFA-status filter
    if ($UserType -eq 'MfaRegistered'    -and -not $r.IsMfaRegistered) { continue }
    if ($UserType -eq 'NotMfaRegistered' -and      $r.IsMfaRegistered) { continue }

    $results.Add([PSCustomObject]@{
        DisplayName              = $r.UserDisplayName
        UserPrincipalName        = $r.UserPrincipalName
        UserType                 = $userType
        AccountEnabled           = $accountEnabled
        IsMfaRegistered          = $r.IsMfaRegistered
        IsMfaCapable             = $r.IsMfaCapable
        IsPasswordlessCapable    = $r.IsPasswordlessCapable
        IsAdmin                  = $r.IsAdmin
        IsSsprRegistered         = $r.IsSsprRegistered
        IsSsprEnabled            = $r.IsSsprEnabled
        IsSsprCapable            = $r.IsSsprCapable
        DefaultMfaMethod         = $r.DefaultMfaMethod
        RegisteredMethods        = ($r.MethodsRegistered -join '; ')
        SystemPreferredAuthMethod = ($r.SystemPreferredAuthenticationMethods -join '; ')
        UserPreferredAuthMethod  = $r.UserPreferredMethodForSecondaryAuthentication
        LastUpdated              = $r.LastUpdatedDateTime
    })
}

# ---------- Export ----------
$reportCsv = Join-Path $LogPath "UserMFAStatus_$timestamp.csv"

if ($results.Count -eq 0) {
    Write-Host "`nNo users match the specified filter." -ForegroundColor Yellow
    "DisplayName,UserPrincipalName,UserType,AccountEnabled,IsMfaRegistered,IsMfaCapable,IsPasswordlessCapable,IsAdmin,IsSsprRegistered,IsSsprEnabled,IsSsprCapable,DefaultMfaMethod,RegisteredMethods,SystemPreferredAuthMethod,UserPreferredAuthMethod,LastUpdated" |
        Out-File -FilePath $reportCsv -Encoding UTF8
    Stop-Transcript | Out-Null
    return
}

$results | Sort-Object IsMfaRegistered, UserPrincipalName | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8

# ---------- Summary ----------
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host ("Total users reported:        {0}" -f $results.Count)

$mfaRegistered = ($results | Where-Object IsMfaRegistered -eq $true).Count
$mfaCapable    = ($results | Where-Object IsMfaCapable    -eq $true).Count
$passwordless  = ($results | Where-Object IsPasswordlessCapable -eq $true).Count
$admins        = ($results | Where-Object IsAdmin -eq $true).Count
$adminsNoMfa   = ($results | Where-Object { $_.IsAdmin -eq $true -and $_.IsMfaRegistered -eq $false }).Count

Write-Host ("MFA registered:              {0}  ({1}%)" -f $mfaRegistered, [math]::Round(($mfaRegistered / $results.Count) * 100, 1))
Write-Host ("MFA capable (active):        {0}  ({1}%)" -f $mfaCapable,    [math]::Round(($mfaCapable    / $results.Count) * 100, 1))
Write-Host ("Passwordless capable:        {0}  ({1}%)" -f $passwordless,  [math]::Round(($passwordless  / $results.Count) * 100, 1))
Write-Host ("Admin accounts in scope:     {0}" -f $admins)

if ($adminsNoMfa -gt 0) {
    Write-Host ("Admins WITHOUT MFA:          {0}  *** REVIEW IMMEDIATELY ***" -f $adminsNoMfa) -ForegroundColor Red
}

Write-Host "`nDefault method breakdown:" -ForegroundColor Yellow
$results | Where-Object IsMfaRegistered -eq $true | Group-Object DefaultMfaMethod | Sort-Object Count -Descending | ForEach-Object {
    $method = if ([string]::IsNullOrEmpty($_.Name)) { '(none set)' } else { $_.Name }
    Write-Host ("  {0,-30} {1}" -f $method, $_.Count)
}

if ($adminsNoMfa -gt 0) {
    Write-Host "`nAdmins without MFA (immediate priority):" -ForegroundColor Red
    $results | Where-Object { $_.IsAdmin -eq $true -and $_.IsMfaRegistered -eq $false } | ForEach-Object {
        Write-Host ("  {0}" -f $_.UserPrincipalName) -ForegroundColor Red
    }
}

Write-Host "`nReport exported: $reportCsv" -ForegroundColor Green

Stop-Transcript | Out-Null