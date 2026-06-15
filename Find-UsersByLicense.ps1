<#
.SYNOPSIS
    Finds users matching a specific license SKU or service plan, including support for
    "missing" lookups.

.DESCRIPTION
    Read-only audit script. Inverse companion to Get-LicenseAssignmentReport.ps1.

    Supports two lookup modes:
      - SKU mode: matches users holding a full license (e.g., 'SPE_E5').
      - Service plan mode: matches users with a specific service plan ENABLED inside
        any license (e.g., 'TEAMS1', 'EXCHANGE_S_ENTERPRISE').

    Use -Missing to invert the result (find users WITHOUT the license/plan).

.PARAMETER Sku
    SKU to search for. Pass either SkuPartNumber ('SPE_E5') or friendly name
    ('Microsoft 365 E5'). Mutually exclusive with -ServicePlan.

.PARAMETER ServicePlan
    Service plan to search for. Pass the ServicePlanName (e.g., 'TEAMS1', 'POWER_BI_PRO').
    Mutually exclusive with -Sku.

.PARAMETER Missing
    Inverts the search: returns users who do NOT have the specified license/plan.

.PARAMETER LimitToLicensedUsers
    For -Missing searches: only consider users who have at least one license.
    Useful for "licensed users missing X" rather than "all users missing X".

.PARAMETER IncludeDisabledAccounts
    When set, includes accounts that are already disabled.

.PARAMETER FilterByDepartment
    Optional. Restrict the search to users in a specific department.

.PARAMETER LogPath
    Folder where the run log and report CSV will be written. Defaults to script directory.

.EXAMPLE
    .\Find-UsersByLicense.ps1 -Sku 'SPE_E5'

.EXAMPLE
    .\Find-UsersByLicense.ps1 -ServicePlan 'POWER_BI_PRO' -Missing -FilterByDepartment 'Finance'

.EXAMPLE
    .\Find-UsersByLicense.ps1 -ServicePlan 'TEAMS1' -Missing -LimitToLicensedUsers
#>

[CmdletBinding(DefaultParameterSetName = 'Sku')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Sku')]
    [string]$Sku,

    [Parameter(Mandatory = $true, ParameterSetName = 'ServicePlan')]
    [string]$ServicePlan,

    [Parameter(Mandatory = $false)]
    [switch]$Missing,

    [Parameter(Mandatory = $false)]
    [switch]$LimitToLicensedUsers,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDisabledAccounts,

    [Parameter(Mandatory = $false)]
    [string]$FilterByDepartment,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "FindUsersByLicense_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

$searchMode  = if ($PSCmdlet.ParameterSetName -eq 'Sku') { 'SKU' } else { 'ServicePlan' }
$searchValue = if ($searchMode -eq 'SKU') { $Sku } else { $ServicePlan }
$matchLogic  = if ($Missing) { 'MISSING' } else { 'HAS' }

Write-Host "=== Find Users By License ===" -ForegroundColor Cyan
Write-Host "Mode:                       $searchMode"
Write-Host "Searching for:              $matchLogic '$searchValue'"
Write-Host "Limit to licensed users:    $LimitToLicensedUsers"
Write-Host "Include disabled accounts:  $IncludeDisabledAccounts"
if ($FilterByDepartment) { Write-Host "Department filter:          $FilterByDepartment" }
Write-Host "Log file:                   $logFile`n"

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
$requiredScopes = @('User.Read.All', 'Organization.Read.All', 'Directory.Read.All')

try {
    $context = Get-MgContext -ErrorAction SilentlyContinue
    $needsConnect = $false
    if (-not $context) { $needsConnect = $true }
    else {
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

# ---------- Friendly SKU name map (subset, same as previous script) ----------
$skuFriendlyName = @{
    'STANDARDPACK' = 'Office 365 E1';            'ENTERPRISEPACK' = 'Office 365 E3'
    'ENTERPRISEPREMIUM' = 'Office 365 E5';       'SPE_E3' = 'Microsoft 365 E3'
    'SPE_E5' = 'Microsoft 365 E5';               'SPE_F1' = 'Microsoft 365 F1'
    'SPE_F3' = 'Microsoft 365 F3';               'SPB' = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_ESSENTIALS' = 'Microsoft 365 Business Basic'
    'O365_BUSINESS_PREMIUM' = 'Microsoft 365 Business Standard'
    'EXCHANGESTANDARD' = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE' = 'Exchange Online (Plan 2)'
    'POWER_BI_PRO' = 'Power BI Pro';             'POWER_BI_STANDARD' = 'Power BI (free)'
    'EMS' = 'Enterprise Mobility + Security E3'; 'EMSPREMIUM' = 'Enterprise Mobility + Security E5'
    'AAD_PREMIUM' = 'Entra ID P1';               'AAD_PREMIUM_P2' = 'Entra ID P2'
    'PROJECTPROFESSIONAL' = 'Project Plan 3';    'PROJECTPREMIUM' = 'Project Plan 5'
    'VISIOCLIENT' = 'Visio Plan 2'
}
function Get-FriendlySkuName {
    param([string]$SkuPartNumber)
    if ($skuFriendlyName.ContainsKey($SkuPartNumber)) { return $skuFriendlyName[$SkuPartNumber] }
    return $SkuPartNumber
}

# ---------- Load tenant SKU catalog ----------
Write-Host "Retrieving tenant SKU catalog..." -ForegroundColor Yellow
try {
    $tenantSkus = Get-MgSubscribedSku -All -ErrorAction Stop
}
catch {
    Write-Error "Failed to retrieve tenant SKUs: $_"
    Stop-Transcript | Out-Null
    return
}
$skuLookup = @{}
foreach ($s in $tenantSkus) { $skuLookup[$s.SkuId] = $s }

# ---------- Resolve search target to ID(s) ----------
$targetSkuId      = $null
$targetPlanName   = $null
$targetSkusContaining = @()

if ($searchMode -eq 'SKU') {
    $match = $tenantSkus | Where-Object {
        $_.SkuPartNumber -eq $Sku -or (Get-FriendlySkuName -SkuPartNumber $_.SkuPartNumber) -eq $Sku
    } | Select-Object -First 1
    if (-not $match) {
        Write-Error "SKU '$Sku' did not match any subscribed SKU. Available SkuPartNumbers: $(($tenantSkus.SkuPartNumber | Sort-Object) -join ', ')"
        Stop-Transcript | Out-Null
        return
    }
    $targetSkuId = $match.SkuId
    Write-Host "Resolved to SkuId: $targetSkuId  ($($match.SkuPartNumber))`n" -ForegroundColor Green
}
else {
    # Service plan mode - find which SKUs include this plan
    foreach ($s in $tenantSkus) {
        if ($s.ServicePlans.ServicePlanName -contains $ServicePlan) {
            $targetSkusContaining += $s
        }
    }
    if ($targetSkusContaining.Count -eq 0) {
        $allPlans = $tenantSkus.ServicePlans.ServicePlanName | Sort-Object -Unique
        Write-Error "Service plan '$ServicePlan' was not found in any subscribed SKU."
        Write-Host "Available service plan names in this tenant:" -ForegroundColor Yellow
        $allPlans | ForEach-Object { Write-Host "  $_" }
        Stop-Transcript | Out-Null
        return
    }
    $targetPlanName = $ServicePlan
    Write-Host "Service plan '$ServicePlan' found in $($targetSkusContaining.Count) SKU(s):" -ForegroundColor Green
    $targetSkusContaining | ForEach-Object {
        Write-Host ("  {0} ({1})" -f (Get-FriendlySkuName -SkuPartNumber $_.SkuPartNumber), $_.SkuPartNumber)
    }
    Write-Host ""
}

# ---------- Retrieve users ----------
Write-Host "Retrieving users..." -ForegroundColor Yellow
$selectProps = @(
    'Id','UserPrincipalName','DisplayName','AccountEnabled','UserType',
    'Department','JobTitle','UsageLocation','Mail',
    'AssignedLicenses','LicenseAssignmentStates'
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

# ---------- Helper: does user have the target SKU? ----------
function Test-UserHasSku {
    param($User, [string]$SkuId)
    return ($User.AssignedLicenses.SkuId -contains $SkuId)
}

# ---------- Helper: does user have the target service plan ENABLED? ----------
# A service plan is "enabled" for a user when:
#   1) The user holds at least one SKU that contains the plan, AND
#   2) The plan is NOT in the user's DisabledPlans for that SKU.
function Test-UserHasServicePlanEnabled {
    param($User, [string]$PlanName, $SkusContainingPlan)

    foreach ($userLic in $User.AssignedLicenses) {
        $sku = $SkusContainingPlan | Where-Object SkuId -eq $userLic.SkuId | Select-Object -First 1
        if (-not $sku) { continue }   # User's license isn't one that contains the plan

        # Find the plan's ID inside this SKU
        $planId = ($sku.ServicePlans | Where-Object ServicePlanName -eq $PlanName).ServicePlanId
        if (-not $planId) { continue }

        # Plan is enabled if it's not in the user's DisabledPlans for this license
        if ($userLic.DisabledPlans -notcontains $planId) {
            return $true
        }
    }
    return $false
}

# ---------- Process ----------
$results = [System.Collections.Generic.List[object]]::new()

foreach ($u in $allUsers) {

    if (-not $IncludeDisabledAccounts -and -not $u.AccountEnabled) { continue }
    if ($FilterByDepartment -and $u.Department -ne $FilterByDepartment) { continue }
    if ($LimitToLicensedUsers -and ($null -eq $u.AssignedLicenses -or $u.AssignedLicenses.Count -eq 0)) { continue }

    # Determine match
    $userHasIt = if ($searchMode -eq 'SKU') {
        Test-UserHasSku -User $u -SkuId $targetSkuId
    } else {
        Test-UserHasServicePlanEnabled -User $u -PlanName $targetPlanName -SkusContainingPlan $targetSkusContaining
    }

    # Apply -Missing inversion
    $included = if ($Missing) { -not $userHasIt } else { $userHasIt }
    if (-not $included) { continue }

    # For HAS results, capture the assignment method/source for context
    $assignmentMethod = $null
    $sourceGroups     = $null

    if (-not $Missing -and $searchMode -eq 'SKU') {
        $states = $u.LicenseAssignmentStates | Where-Object SkuId -eq $targetSkuId
        $isDirect    = $states | Where-Object { $null -eq $_.AssignedByGroup }
        $groupStates = $states | Where-Object { $null -ne $_.AssignedByGroup }
        $assignmentMethod = if ($isDirect -and $groupStates) { 'Direct + Group' }
                            elseif ($isDirect) { 'Direct' }
                            elseif ($groupStates) { 'Group' }
                            else { 'Unknown' }
        if ($groupStates) {
            $sourceGroups = ($groupStates.AssignedByGroup | Sort-Object -Unique) -join '; '
        }
    }

    # User's full license list for context
    $userLicenseList = @()
    if ($u.AssignedLicenses) {
        foreach ($l in $u.AssignedLicenses) {
            if ($skuLookup.ContainsKey($l.SkuId)) {
                $userLicenseList += Get-FriendlySkuName -SkuPartNumber $skuLookup[$l.SkuId].SkuPartNumber
            }
        }
    }

    $results.Add([PSCustomObject]@{
        DisplayName        = $u.DisplayName
        UserPrincipalName  = $u.UserPrincipalName
        AccountEnabled     = $u.AccountEnabled
        UserType           = $u.UserType
        Department         = $u.Department
        JobTitle           = $u.JobTitle
        UsageLocation      = $u.UsageLocation
        Mail               = $u.Mail
        AllLicenses        = ($userLicenseList | Sort-Object -Unique) -join '; '
        AssignmentMethod   = $assignmentMethod
        SourceGroupIds     = $sourceGroups
    })
}

# ---------- Export ----------
$verbForFile = if ($Missing) { 'Missing' } else { 'Has' }
$cleanTarget = $searchValue -replace '[^a-zA-Z0-9_]', '_'
$reportCsv = Join-Path $LogPath "FindUsersByLicense_${verbForFile}_${cleanTarget}_$timestamp.csv"

if ($results.Count -eq 0) {
    Write-Host "`nNo users matched the criteria." -ForegroundColor Yellow
    "DisplayName,UserPrincipalName,AccountEnabled,UserType,Department,JobTitle,UsageLocation,Mail,AllLicenses,AssignmentMethod,SourceGroupIds" |
        Out-File -FilePath $reportCsv -Encoding UTF8
    Stop-Transcript | Out-Null
    return
}

$results | Sort-Object Department, UserPrincipalName | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8

# ---------- Summary ----------
Write-Host "=== Summary ===" -ForegroundColor Cyan
Write-Host ("Search:               {0} '{1}'" -f $matchLogic, $searchValue)
Write-Host ("Matching users:       {0}" -f $results.Count)

if ($searchMode -eq 'SKU' -and -not $Missing) {
    $sku = $tenantSkus | Where-Object SkuId -eq $targetSkuId
    Write-Host ("Tenant capacity:      {0} consumed / {1} enabled" -f $sku.ConsumedUnits, $sku.PrepaidUnits.Enabled)
}

if ($results | Where-Object Department) {
    Write-Host "`nBy department:" -ForegroundColor Yellow
    $results | Group-Object Department | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object {
        $dept = if ($_.Name) { $_.Name } else { '(none)' }
        Write-Host ("  {0,-30} {1}" -f $dept, $_.Count)
    }
}

if (-not $Missing -and $searchMode -eq 'SKU') {
    Write-Host "`nAssignment method breakdown:" -ForegroundColor Yellow
    $results | Where-Object AssignmentMethod | Group-Object AssignmentMethod | ForEach-Object {
        Write-Host ("  {0,-18} {1}" -f $_.Name, $_.Count)
    }
}

Write-Host "`nReport exported: $reportCsv" -ForegroundColor Green

Stop-Transcript | Out-Null