<#
.SYNOPSIS
    Copies membership from a source group to a target group in Exchange Online.

.DESCRIPTION
    Retrieves all members of the source group and adds them to the target group,
    skipping users already present in the target. Supports DLs, mail-enabled security
    groups, and Microsoft 365 groups on either side.

    Includes confirmation prompt, -WhatIf support, optional exclusion CSV, and full
    transcript + results CSV logging.

.PARAMETER SourceGroup
    The source group's primary SMTP address, alias, or DisplayName.

.PARAMETER TargetGroup
    The target group's primary SMTP address, alias, or DisplayName.

.PARAMETER ExcludeFromCsv
    Optional path to a CSV with a 'UserPrincipalName' column listing users to skip.

.PARAMETER LogPath
    Folder where the run log will be written. Defaults to script directory.

.PARAMETER Force
    Skip the interactive confirmation prompt.

.EXAMPLE
    .\Copy-GroupMembership.ps1 -SourceGroup "TeamA@contoso.com" -TargetGroup "TeamB@contoso.com" -WhatIf

.EXAMPLE
    .\Copy-GroupMembership.ps1 -SourceGroup "TeamA@contoso.com" -TargetGroup "TeamB@contoso.com" -ExcludeFromCsv "C:\Temp\skip.csv"
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceGroup,    # <-- REPLACE at runtime: e.g. "OldGroup@yourdomain.com"

    [Parameter(Mandatory = $true)]
    [string]$TargetGroup,    # <-- REPLACE at runtime: e.g. "NewGroup@yourdomain.com"

    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ExcludeFromCsv,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Split-Path -Parent $MyInvocation.MyCommand.Path),

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# ---------- Setup ----------
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $LogPath "CopyGroupMembership_$timestamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

Write-Host "=== Copy Group Membership ===" -ForegroundColor Cyan
Write-Host "Source:   $SourceGroup"
Write-Host "Target:   $TargetGroup"
if ($ExcludeFromCsv) { Write-Host "Exclude:  $ExcludeFromCsv" }
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

# ---------- Helper: resolve group + type ----------
function Resolve-Group {
    param([string]$Identity)

    try {
        $g = Get-DistributionGroup -Identity $Identity -ErrorAction Stop
        return [PSCustomObject]@{ Group = $g; Type = $g.RecipientTypeDetails }
    } catch {
        try {
            $g = Get-UnifiedGroup -Identity $Identity -ErrorAction Stop
            return [PSCustomObject]@{ Group = $g; Type = 'GroupMailbox' }
        } catch {
            return $null
        }
    }
}

# ---------- Validate source ----------
$src = Resolve-Group -Identity $SourceGroup
if (-not $src) {
    Write-Error "Source group '$SourceGroup' not found."
    Stop-Transcript | Out-Null
    return
}
Write-Host "Source resolved: $($src.Group.DisplayName) [$($src.Type)]" -ForegroundColor Green

# ---------- Validate target ----------
$tgt = Resolve-Group -Identity $TargetGroup
if (-not $tgt) {
    Write-Error "Target group '$TargetGroup' not found."
    Stop-Transcript | Out-Null
    return
}
Write-Host "Target resolved: $($tgt.Group.DisplayName) [$($tgt.Type)]`n" -ForegroundColor Green

# ---------- Get source members ----------
try {
    if ($src.Type -eq 'GroupMailbox') {
        $sourceMembers = Get-UnifiedGroupLinks -Identity $SourceGroup -LinkType Members -ResultSize Unlimited -ErrorAction Stop
    } else {
        $sourceMembers = Get-DistributionGroupMember -Identity $SourceGroup -ResultSize Unlimited -ErrorAction Stop
    }
} catch {
    Write-Error "Failed to retrieve source members: $_"
    Stop-Transcript | Out-Null
    return
}

# ---------- Get target members (for duplicate skip) ----------
try {
    if ($tgt.Type -eq 'GroupMailbox') {
        $targetMembers = Get-UnifiedGroupLinks -Identity $TargetGroup -LinkType Members -ResultSize Unlimited -ErrorAction Stop
    } else {
        $targetMembers = Get-DistributionGroupMember -Identity $TargetGroup -ResultSize Unlimited -ErrorAction Stop
    }
    $existingTargetSmtp = $targetMembers | Select-Object -ExpandProperty PrimarySmtpAddress
} catch {
    Write-Error "Failed to retrieve target members: $_"
    Stop-Transcript | Out-Null
    return
}

# ---------- Load exclusion list (optional) ----------
$excludeList = @()
if ($ExcludeFromCsv) {
    $excludeData = Import-Csv -Path $ExcludeFromCsv
    if (-not ($excludeData | Get-Member -Name 'UserPrincipalName' -MemberType NoteProperty)) {
        Write-Error "Exclusion CSV must contain a 'UserPrincipalName' column."
        Stop-Transcript | Out-Null
        return
    }
    $excludeList = $excludeData.UserPrincipalName | ForEach-Object { $_.Trim().ToLower() }
    Write-Host "Loaded $($excludeList.Count) excluded UPN(s) from CSV.`n" -ForegroundColor Yellow
}

# ---------- Confirmation gate ----------
$plannedCount = ($sourceMembers | Where-Object {
    ($existingTargetSmtp -notcontains $_.PrimarySmtpAddress) -and
    ($excludeList -notcontains $_.PrimarySmtpAddress.ToLower())
}).Count

Write-Host "Source members:           $($sourceMembers.Count)"
Write-Host "Already in target:        $(($sourceMembers | Where-Object { $existingTargetSmtp -contains $_.PrimarySmtpAddress }).Count)"
Write-Host "Excluded by CSV:          $($excludeList.Count)"
Write-Host "Planned additions:        $plannedCount`n"

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

foreach ($member in $sourceMembers) {
    $upn  = if ($member.WindowsLiveID) { $member.WindowsLiveID } else { $member.PrimarySmtpAddress }
    $smtp = $member.PrimarySmtpAddress
    if ([string]::IsNullOrWhiteSpace($smtp)) { continue }

    $status = 'Unknown'
    $detail = ''

    try {
        if ($excludeList -contains $smtp.ToLower()) {
            $status = 'Skipped'
            $detail = 'Excluded by CSV'
        }
        elseif ($existingTargetSmtp -contains $smtp) {
            $status = 'Skipped'
            $detail = 'Already in target'
        }
        elseif ($PSCmdlet.ShouldProcess($upn, "Add to $TargetGroup")) {
            if ($tgt.Type -eq 'GroupMailbox') {
                Add-UnifiedGroupLinks -Identity $TargetGroup -LinkType Members -Links $smtp -ErrorAction Stop
            } else {
                Add-DistributionGroupMember -Identity $TargetGroup -Member $smtp -BypassSecurityGroupManagerCheck -ErrorAction Stop
            }
            $status = 'Added'
            $detail = 'Success'
        }
        else {
            $status = 'WhatIf'
            $detail = 'Would add (dry run)'
        }
    }
    catch {
        $status = 'Failed'
        $detail = $_.Exception.Message
    }

    $results.Add([PSCustomObject]@{
        UserPrincipalName  = $upn
        PrimarySmtpAddress = $smtp
        DisplayName        = $member.DisplayName
        Status             = $status
        Detail             = $detail
    })

    Write-Host ("[{0,-7}] {1}  -  {2}" -f $status, $upn, $detail)
}

# ---------- Summary ----------
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
$results | Group-Object Status | ForEach-Object {
    Write-Host ("{0,-8}: {1}" -f $_.Name, $_.Count)
}

$resultsCsv = Join-Path $LogPath "CopyGroupMembership_Results_$timestamp.csv"
$results | Export-Csv -Path $resultsCsv -NoTypeInformation -Encoding UTF8
Write-Host "`nResults exported: $resultsCsv" -ForegroundColor Green

Stop-Transcript | Out-Null