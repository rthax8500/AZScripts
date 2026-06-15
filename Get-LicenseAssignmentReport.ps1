<#
.SYNOPSIS
    Generates a comprehensive license assignment report showing direct vs. group-based assignments.

.DESCRIPTION
    Read-only audit script. For each licensed user, reports which licenses they hold,
    how each was assigned (direct vs. group-based), the source group(s) for inherited
    licenses, and any disabled service plans. Includes friendly SKU name translation.

    Useful for license cleanup, audit, cost optimization, and group-based licensing
    troubleshooting.

.PARAMETER LicensedOnly
    When set, includes only users with at least one license. Default behavior includes
    all users (so unlicensed users are also visible for context).

.PARAMETER IncludeDisabledAccounts
    When set, includes accounts that are already disabled.

.PARAMETER FilterBySku
    Optional. Restrict output to users holding a specific SKU. Pass either the SkuPartNumber
    (e.g., 'SPE_E5') or friendly name (e.g., 'Microsoft 365 E5').

.PARAMETER LogPath
    Folder where the run log and report CSV will be written. Defaults to script directory.

.EXAMPLE
    .\Get-LicenseAssignmentReport.ps1

.EXAMPLE
    .\Get-LicenseAssignmentReport.ps1 -LicensedOnly -FilterBySku 'SPE_E5'
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$LicensedOnly,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDisabledAccounts,

    [Parameter(Mandatory = $false)]
    [string]$FilterBySku,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "LicenseAssignmentReport_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host "=== License Assignment Report ===" -ForegroundColor Cyan
Write-Host "Licensed users only:       $LicensedOnly"
Write-Host "Include disabled accounts: $IncludeDisabledAccounts"
if ($FilterBySku) { Write-Host "SKU filter:                $FilterBySku" }
Write-Host "Log file:                  $logFile`n"

# ---------- Module check ----------
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
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
    'User.Read.All',
    'Group.Read.All',
    'Organization.Read.All',
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
        Write-Host "Already connected as $($context.Account)." -ForegroundColor Green
    }
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    Stop-Transcript | Out-Null
    return
}

# ---------- Friendly SKU name map ----------
# Microsoft's internal SkuPartNumbers are not always recognizable. This map covers the
# most common ones; unknown SKUs fall back to the SkuPartNumber as-is in the report.
$skuFriendlyName = @{
    'STANDARDPACK'                     = 'Office 365 E1'
    'ENTERPRISEPACK'                   = 'Office 365 E3'
    'ENTERPRISEPREMIUM'                = 'Office 365 E5'
    'ENTERPRISEPREMIUM_NOPSTNCONF'     = 'Office 365 E5 (no PSTN)'
    'SPE_E3'                           = 'Microsoft 365 E3'
    'SPE_E5'                           = 'Microsoft 365 E5'
    'SPE_F1'                           = 'Microsoft 365 F1'
    'SPE_F3'                           = 'Microsoft 365 F3'
    'SPB'                              = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_ESSENTIALS'         = 'Microsoft 365 Business Basic'
    'O365_BUSINESS_PREMIUM'            = 'Microsoft 365 Business Standard'
    'M365_BUSINESS_BASIC'              = 'Microsoft 365 Business Basic'
    'M365_BUSINESS_STANDARD'           = 'Microsoft 365 Business Standard'
    'M365_BUSINESS_PREMIUM'            = 'Microsoft 365 Business Premium'
    'EXCHANGESTANDARD'                 = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'               = 'Exchange Online (Plan 2)'
    'EXCHANGEDESKLESS'                 = 'Exchange Online Kiosk'
    'POWER_BI_STANDARD'                = 'Power BI (free)'
    'POWER_BI_PRO'                     = 'Power BI Pro'
    'POWER_BI_PREMIUM_PER_USER'        = 'Power BI Premium Per User'
    'TEAMS_EXPLORATORY'                = 'Teams Exploratory'
    'TEAMS_FREE'                       = 'Teams Free'
    'EMS'                              = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                       = 'Enterprise Mobility + Security E5'
    'AAD_PREMIUM'                      = 'Entra ID P1'
    'AAD_PREMIUM_P2'                   = 'Entra ID P2'
    'PROJECTPROFESSIONAL'              = 'Project Plan 3'
    'PROJECTPREMIUM'                   = 'Project Plan 5'
    'VISIOCLIENT'                      = 'Visio Plan 2'
    'WIN10_PRO_ENT_SUB'                = 'Windows 10/11 Enterprise E3'
    'WIN10_VDA_E5'                     = 'Windows 10/11 Enterprise E5'
    'DEFENDER_ENDPOINT_P1'             = 'Defender for Endpoint P1'
    'DEFENDER_ENDPOINT_P2'             = 'Defender for Endpoint P2'
    'IDENTITY_THREAT_PROTECTION'       = 'Microsoft 365 E5 Security'
    'INFORMATION_PROTECTION_COMPLIANCE'= 'Microsoft 365 E5 Compliance'
    'FLOW_FREE'                        = 'Power Automate (free)'
    'POWERAUTOMATE_ATTENDED_RPA'       = 'Power Automate per user with RPA'
    'POWER_BI_INDIVIDUAL_USE'          = 'Power BI for Individual Use'
}

function Get-FriendlySkuName {
    param([string]$SkuPartNumber)
    if ($skuFriendlyName.ContainsKey($SkuPartNumber)) { return $skuFriendlyName[$SkuPartNumber] }
    return $SkuPartNumber
}

# ---------- Build subscribed SKU lookup ----------
Write-Host "Retrieving tenant SKU catalog..." -ForegroundColor Yellow
try {
    $tenantSkus = Get-MgSubscribedSku -All -ErrorAction Stop
    $skuLookup = @{}
    foreach ($s in $tenantSkus) {
        $skuLookup[$s.SkuId] = $s
    }
    Write-Host "Tenant has $($tenantSkus.Count) SKU(s) subscribed.`n" -ForegroundColor Green
}
catch {
    Write-Error "Failed to retrieve tenant SKUs: $_"
    Stop-Transcript | Out-Null
    return
}

# ---------- Resolve filter to SkuId if FilterBySku was provided ----------
$filterSkuId = $null
if ($FilterBySku) {
    $match = $tenantSkus | Where-Object {
        $_.SkuPartNumber -eq $FilterBySku -or (Get-FriendlySkuName -SkuPartNumber $_.SkuPartNumber) -eq $FilterBySku
    } | Select-Object -First 1
    if (-not $match) {
        Write-Error "FilterBySku value '$FilterBySku' did not match any subscribed SKU. Available SkuPartNumbers: $(($tenantSkus.SkuPartNumber | Sort-Object) -join ', ')"
        Stop-Transcript | Out-Null
        return
    }
    $filterSkuId = $match.SkuId
    Write-Host "Filter resolved to SkuId: $filterSkuId ($($match.SkuPartNumber))`n" -ForegroundColor Green
}

# ---------- Build group lookup for resolving group-based assignment sources ----------
Write-Host "Building group cache for assignment source resolution..." -ForegroundColor Yellow
$groupCache = @{}
try {
    Get-MgGroup -All -Property Id, DisplayName, AssignedLicenses -ErrorAction Stop | ForEach-Object {
        $groupCache[$_.Id] = $_
    }
    Write-Host "Cached $($groupCache.Count) group(s).`n" -ForegroundColor Green
}
catch {
    Write-Warning "Could not build group cache; SourceGroups column may be incomplete: $_"
}

# ---------- Retrieve users ----------
Write-Host "Retrieving users with license assignment data..." -ForegroundColor Yellow
$selectProps = @(
    'Id',
    'UserPrincipalName',
    'DisplayName',
    'AccountEnabled',
    'UserType',
    'Department',
    'JobTitle',
    'AssignedLicenses',
    'LicenseAssignmentStates',
    'UsageLocation'
)

try {
    $allUsers = Get-MgUser -All -Property $selectProps -ErrorAction Stop
    Write-Host "Retrieved $($allUsers.Count) user(s).`n" -ForegroundColor Green
}
catch {
    Write-Error "Failed to retrieve users: $_"
    Stop-Transcript | Out-Null
    return
}

# ---------- Process ----------
$results = [System.Collections.Generic.List[object]]::new()

foreach ($u in $allUsers) {

    if (-not $IncludeDisabledAccounts -and -not $u.AccountEnabled) { continue }
    if ($LicensedOnly -and ($null -eq $u.AssignedLicenses -or $u.AssignedLicenses.Count -eq 0)) { continue }
    if ($filterSkuId -and ($u.AssignedLicenses.SkuId -notcontains $filterSkuId)) { continue }

    if ($null -eq $u.AssignedLicenses -or $u.AssignedLicenses.Count -eq 0) {
        $results.Add([PSCustomObject]@{
            DisplayName        = $u.DisplayName
            UserPrincipalName  = $u.UserPrincipalName
            AccountEnabled     = $u.AccountEnabled
            UserType           = $u.UserType
            Department         = $u.Department
            JobTitle           = $u.JobTitle
            UsageLocation      = $u.UsageLocation
            SkuPartNumber      = '(no licenses)'
            FriendlyName       = '(no licenses)'
            AssignmentMethod   = $null
            SourceGroups       = $null
            DisabledServicePlans = $null
            HasErrors          = $false
            ErrorDetail        = $null
        })
        continue
    }

    foreach ($lic in $u.AssignedLicenses) {

        $skuPartNumber = if ($skuLookup.ContainsKey($lic.SkuId)) { $skuLookup[$lic.SkuId].SkuPartNumber } else { '(unknown SKU)' }
        $friendly      = Get-FriendlySkuName -SkuPartNumber $skuPartNumber

        # Determine assignment method by examining LicenseAssignmentStates
        $statesForThisSku = $u.LicenseAssignmentStates | Where-Object { $_.SkuId -eq $lic.SkuId }

        $isDirect       = $statesForThisSku | Where-Object { $null -eq $_.AssignedByGroup }
        $groupStates    = $statesForThisSku | Where-Object {  $null -ne $_.AssignedByGroup }

        $method = if ($isDirect -and $groupStates) { 'Direct + Group' }
                  elseif ($isDirect)               { 'Direct' }
                  elseif ($groupStates)            { 'Group' }
                  else                             { 'Unknown' }

        # Resolve source group names
        $sourceGroupNames = @()
        foreach ($gs in $groupStates) {
            $sourceGroupNames += if ($groupCache.ContainsKey($gs.AssignedByGroup)) {
                $groupCache[$gs.AssignedByGroup].DisplayName
            } else {
                "(group $($gs.AssignedByGroup))"
            }
        }

        # Translate disabled service plan IDs to friendly names where possible
        $disabledPlanNames = @()
        if ($lic.DisabledPlans -and $lic.DisabledPlans.Count -gt 0 -and $skuLookup.ContainsKey($lic.SkuId)) {
            $skuServicePlans = $skuLookup[$lic.SkuId].ServicePlans
            foreach ($disabledId in $lic.DisabledPlans) {
                $planName = ($skuServicePlans | Where-Object ServicePlanId -eq $disabledId).ServicePlanName
                $disabledPlanNames += if ($planName) { $planName } else { "$disabledId" }
            }
        }

        # Capture error states (e.g., disabled SKUs, missing usage location)
        $errorState = $statesForThisSku | Where-Object { $_.State -eq 'Error' } | Select-Object -First 1
        $hasError   = [bool]$errorState
        $errorDetail = if ($errorState) { $errorState.Error } else { $null }

        $results.Add([PSCustomObject]@{
            DisplayName          = $u.DisplayName
            UserPrincipalName    = $u.UserPrincipalName
            AccountEnabled       = $u.AccountEnabled
            UserType             = $u.UserType
            Department           = $u.Department
            JobTitle             = $u.JobTitle
            UsageLocation        = $u.UsageLocation
            SkuPartNumber        = $skuPartNumber
            FriendlyName         = $friendly
            AssignmentMethod     = $method
            SourceGroups         = ($sourceGroupNames | Sort-Object -Unique) -join '; '
            DisabledServicePlans = ($disabledPlanNames | Sort-Object) -join '; '
            HasErrors            = $hasError
            ErrorDetail          = $errorDetail
        })
    }
}

# ---------- Export ----------
$reportCsv = Join-Path $LogPath "LicenseAssignmentReport_$timestamp.csv"

if ($results.Count -eq 0) {
    Write-Host "`nNo records to report." -ForegroundColor Yellow
    Stop-Transcript | Out-Null
    return
}

$results | Sort-Object UserPrincipalName, FriendlyName | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8

# ---------- Summary ----------
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
$uniqueUsers = ($results | Select-Object -Unique UserPrincipalName).Count
Write-Host ("Unique users in report:      {0}" -f $uniqueUsers)
Write-Host ("Total assignment rows:       {0}" -f $results.Count)

$errorRows = ($results | Where-Object HasErrors -eq $true).Count
if ($errorRows -gt 0) {
    Write-Host ("Rows with errors:            {0}  *** REVIEW ***" -f $errorRows) -ForegroundColor Red
}

Write-Host "`nLicense usage by SKU (consumed vs. enabled):" -ForegroundColor Yellow
foreach ($s in ($tenantSkus | Sort-Object SkuPartNumber)) {
    $friendly = Get-FriendlySkuName -SkuPartNumber $s.SkuPartNumber
    $consumed = $s.ConsumedUnits
    $enabled  = $s.PrepaidUnits.Enabled
    $available = $enabled - $consumed
    Write-Host ("  {0,-40} {1,5} / {2,-5}  ({3} available)" -f $friendly, $consumed, $enabled, $available)
}

Write-Host "`nAssignment method breakdown:" -ForegroundColor Yellow
$results | Where-Object AssignmentMethod | Group-Object AssignmentMethod | Sort-Object Count -Descending | ForEach-Object {
    Write-Host ("  {0,-18} {1}" -f $_.Name, $_.Count)
}

if ($errorRows -gt 0) {
    Write-Host "`nUsers with assignment errors:" -ForegroundColor Red
    $results | Where-Object HasErrors -eq $true | Select-Object UserPrincipalName, FriendlyName, ErrorDetail -Unique | ForEach-Object {
        Write-Host ("  {0,-40} {1,-30} -> {2}" -f $_.UserPrincipalName, $_.FriendlyName, $_.ErrorDetail) -ForegroundColor Red
    }
}

Write-Host "`nReport exported: $reportCsv" -ForegroundColor Green

Stop-Transcript | Out-Null