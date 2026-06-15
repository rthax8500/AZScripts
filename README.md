PowerShell Microsoft 365 Administration Toolkit

A collection of PowerShell scripts for automating common administrative tasks across
Microsoft 365, Exchange Online, and Entra ID — focused on mailbox management, group
and membership operations, and identity/license reporting. Built to replace repetitive
manual work in the admin portals with repeatable, parameter-driven commands that log
what they do.

What's Inside

Mailbox & Exchange Online

ScriptPurposeConvert-UserToSharedMailbox.ps1Converts a user mailbox into a shared mailbox.
Grant-MailboxFullAccess.ps1Grants a user Full Access permission to another mailbox.
Set-OutOfOfficeBulk.ps1Sets automatic replies (Out of Office) across multiple mailboxes at once.
Get-MailboxSizeReport.ps1Generates a report of mailbox sizes across the tenant.
Get-InactiveMailboxes.ps1Identifies mailboxes with no recent activity.

Groups & Membership

ScriptPurpose
Add-BulkUsersToMailSecurityGroup.ps1 Bulk-adds users to a mail-enabled security group from a CSV. (See Detailed Usage below.)
Remove-BulkUsersFromMailSecurityGroup.ps1 Bulk-removes users from a mail-enabled security group from a CSV. (See Detailed Usage below.)
Convert-DLToM365Group.ps1Converts a classic distribution list to a Microsoft 365 group.
Copy-GroupMembership.ps1 Copies the membership of one group onto another.
Export-GroupMembership.ps1Exports a group's membership to a file (e.g. CSV) for review or backup.
Find-EmptyOrOrphanedGroups.ps1Flags groups that are empty or have no owner.

Identity & Licensing

ScriptPurposeGet-LicenseAssignmentReport.ps1 Reports license assignments across the tenant.
Find-UsersByLicense.ps1 Finds all users assigned a specific license.
Get-UserMFAStatus.ps1 Reports each user's multi-factor authentication status.
Get-StaleUsers.ps1 Identifies inactive or stale user accounts.

Requirements


Windows PowerShell 5.1 or PowerShell 7+
The following modules, depending on the script:

ExchangeOnlineManagement — for mailbox and group operations





powershell    Install-Module ExchangeOnlineManagement -Scope CurrentUser


[Microsoft.Graph OR MSOnline/AzureAD — confirm which the identity scripts use] — for MFA, license, and stale-user reporting
An account with the appropriate admin role in your tenant (e.g. Exchange Administrator, User Administrator)


General Usage

Connect to the relevant service first, then run the script you need. For example:

"Connect-ExchangeOnline -UserPrincipalName "[admin@yourdomain.com]"

.\Get-MailboxSizeReport.ps1 -OutputPath "[C:\Reports\MailboxSizes.csv]"



Detailed Usage — Add-BulkUsersToMailSecurityGroup.ps1

Adds a list of users (supplied in a CSV) to a mail-enabled security group in Exchange
Online. It verifies each user exists, skips anyone already a member, performs the
additions, and writes a full transcript log plus a results report.

CSV format

The CSV must contain a single header column named exactly UserPrincipalName. Each row
is one user's sign-in address. Blank rows are ignored.

UserPrincipalName
jane.doe@yourdomain.com
john.smith@yourdomain.com
alex.lee@yourdomain.com

Parameters 
.\-GroupIdentity - Target mail-enabled security groups. Accepts primary SMTP address, alias, or Display names
.\-CsvPath - Full path to the CSV file, must exist and contain UserPrincipalName column
.\-LogPath Folder for the log and results file, Default to the working folder. 

Examples

powershell# Standard run
.\Add-BulkUsersToMailSecurityGroup.ps1 -GroupIdentity "[GroupAlias@yourdomain.com]" -CsvPath "[C:\Temp\users.csv]"

# Dry run — previews changes, makes none
.\Add-BulkUsersToMailSecurityGroup.ps1 -GroupIdentity "[GroupAlias@yourdomain.com]" -CsvPath "[C:\Temp\users.csv]" -WhatIf

# Custom log location
.\Add-BulkUsersToMailSecurityGroup.ps1 -GroupIdentity "[GroupAlias@yourdomain.com]" -CsvPath "[C:\Temp\users.csv]" -LogPath "[C:\Temp\Logs]"

Output

Two timestamped files are written to -LogPath:


AddToGroup_<timestamp>.log — full transcript of the run.
AddToGroup_Results_<timestamp>.csv — one row per user with UserPrincipalName, Status, and Detail.


Status values: Added (success), Skipped (already a member), WhatIf (dry-run preview), Failed (see Detail for the reason).

Safety


Run with -WhatIf first to preview changes before committing them.
The script refuses to act on any group that is not a mail-enabled security group, preventing accidental edits to the wrong group type.



Detailed Usage — Remove-BulkUsersFromMailSecurityGroup.ps1

Removes a list of users (supplied in a CSV) from a mail-enabled security group in Exchange
Online. It validates each user, skips anyone not currently in the group, logs every action,
and — because removals are destructive — requires explicit confirmation before making live
changes.

CSV format

Identical to the Add script: a single header column named exactly UserPrincipalName, one
UPN per row. Blank rows are ignored.

UserPrincipalName
jane.doe@yourdomain.com
john.smith@yourdomain.com

Parameters 
.\-GroupIdentity - Target mail-enabled security groups. Accepts primary SMTP address, alias, or Display names
.\-CsvPath - Full path to the CSV file, must exist and contain UserPrincipalName column
.\-LogPath - Folder for the log and results file, Default to the working folder. 
.\-Force - Skips interactive confirmation prompt. Use with caution. 

Examples

powershell# Dry run — previews removals, makes none
.\Remove-BulkUsersFromMailSecurityGroup.ps1 -GroupIdentity "[GroupAlias@yourdomain.com]" -CsvPath "[C:\Temp\users.csv]" -WhatIf

# Standard run — prompts for confirmation before removing
.\Remove-BulkUsersFromMailSecurityGroup.ps1 -GroupIdentity "[GroupAlias@yourdomain.com]" -CsvPath "[C:\Temp\users.csv]"

# Unattended run — skips the prompt (use carefully)
.\Remove-BulkUsersFromMailSecurityGroup.ps1 -GroupIdentity "[GroupAlias@yourdomain.com]" -CsvPath "[C:\Temp\users.csv]" -Force

Output

Two timestamped files are written to -LogPath:


RemoveFromGroup_<timestamp>.log — full transcript of the run.
RemoveFromGroup_Results_<timestamp>.csv — one row per user with UserPrincipalName, Status, and Detail.


Status values: Removed (success), Skipped (not a current member), WhatIf (dry-run preview), Failed (see Detail for the reason).

Safety

This script applies three layers of protection, escalating with the risk of the operation:


Group-type check — refuses to act on any group that is not a mail-enabled security group.
-WhatIf dry run — preview exactly which users would be removed before committing.
Interactive confirmation — a live run requires typing YES to proceed, unless -Force is explicitly set.



Notes


Scripts that modify objects in bulk (the Add-, Remove-, and Convert- scripts)
change live tenant data. Review your input list and test against a small set — or use
a dry-run switch where available — before running at scale.



Built and maintained by Ryan Thackston.
