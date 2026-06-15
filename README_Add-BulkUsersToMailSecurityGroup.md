===============================================================================
 Add-BulkUsersToMailSecurityGroup.ps1
 README
===============================================================================

-------------------------------------------------------------------------------
 OVERVIEW
-------------------------------------------------------------------------------
This PowerShell script adds a list of users (supplied in a CSV file) to a
mail-enabled security group in Exchange Online. It connects to Exchange Online,
checks that each user exists, skips anyone who is already a member, performs the
additions, and writes a full log plus a results report.

Think of it like a guest-list manager for a mailing group: you hand it a
spreadsheet of names, and it adds each valid person, ignores anyone already on
the list, and hands you back a receipt of exactly what it did.


-------------------------------------------------------------------------------
 REQUIREMENTS
-------------------------------------------------------------------------------
- Windows PowerShell 5.1 or PowerShell 7+
- ExchangeOnlineManagement module installed
      Install-Module ExchangeOnlineManagement -Scope CurrentUser
- An account with permission to manage distribution/security groups in
  Exchange Online (e.g. Exchange Administrator or an equivalent role)
- A CSV file containing a column named exactly:  UserPrincipalName


-------------------------------------------------------------------------------
 CSV FORMAT
-------------------------------------------------------------------------------
The CSV must contain a single header column titled "UserPrincipalName".
Each row is one user's UPN (their sign-in address).

Example (users.csv):

      UserPrincipalName
      jane.doe@yourdomain.com
      john.smith@yourdomain.com
      alex.lee@yourdomain.com


-------------------------------------------------------------------------------
 PARAMETERS
-------------------------------------------------------------------------------
-GroupIdentity   (required)
      The target mail-enabled security group. Accepts the primary SMTP
      address, alias, or DisplayName.
      Example placeholder:  "GroupAlias@yourdomain.com"   <-- REPLACE

-CsvPath         (required)
      Full path to the CSV file. Must exist and contain a
      "UserPrincipalName" column.
      Example placeholder:  "C:\Temp\users.csv"           <-- REPLACE

-LogPath         (optional)
      Folder where the log and results files are written.
      Defaults to the folder the script is run from.
      Example placeholder:  "C:\Temp\Logs"                <-- REPLACE


-------------------------------------------------------------------------------
 USAGE
-------------------------------------------------------------------------------
1. Open a PowerShell window.

2. (First time only) Install the Exchange Online module:
      Install-Module ExchangeOnlineManagement -Scope CurrentUser

3. Run the script. Replace the placeholder values shown in < > brackets.

   Standard run:
      .\Add-BulkUsersToMailSecurityGroup.ps1 -GroupIdentity "<GroupAlias@yourdomain.com>" -CsvPath "<C:\Temp\users.csv>"

   Dry run (preview only, makes no changes):
      .\Add-BulkUsersToMailSecurityGroup.ps1 -GroupIdentity "<GroupAlias@yourdomain.com>" -CsvPath "<C:\Temp\users.csv>" -WhatIf

   Custom log location:
      .\Add-BulkUsersToMailSecurityGroup.ps1 -GroupIdentity "<GroupAlias@yourdomain.com>" -CsvPath "<C:\Temp\users.csv>" -LogPath "<C:\Temp\Logs>"

4. If you are not already connected, a sign-in prompt will appear.


-------------------------------------------------------------------------------
 WHAT THE SCRIPT DOES (STEP BY STEP)
-------------------------------------------------------------------------------
1. Starts a transcript log file in the log folder.
2. Connects to Exchange Online (or reuses an existing connection).
3. Confirms the target group exists AND is a mail-enabled security group.
   If it is any other group type, the script stops to avoid mistakes.
4. Loads the CSV and verifies the "UserPrincipalName" column is present.
5. Caches current group members so duplicate checks are fast.
6. For each user in the CSV:
      - Confirms the user exists in Exchange Online.
      - Skips them if they are already a member.
      - Adds them if they are not.
      - Records the outcome.
7. Prints a summary grouped by status and exports a results CSV.


-------------------------------------------------------------------------------
 OUTPUT FILES
-------------------------------------------------------------------------------
Both files are timestamped (yyyyMMdd_HHmmss) and saved to -LogPath:

   AddToGroup_<timestamp>.log
      Full transcript of the run.

   AddToGroup_Results_<timestamp>.csv
      One row per user with: UserPrincipalName, Status, Detail.

Possible Status values:
   Added    - User was successfully added.
   Skipped  - User was already a member.
   WhatIf   - Dry-run preview; nothing was changed.
   Failed   - An error occurred (the reason is in the Detail column).
   Unknown  - Status was never set (should not normally appear).


-------------------------------------------------------------------------------
 SAFETY NOTES
-------------------------------------------------------------------------------
- Use -WhatIf first to preview changes before running for real.
- The script will refuse to act on a group that is not a mail-enabled
  security group, which prevents accidental edits to the wrong group type.
- Blank rows in the CSV are ignored automatically.


-------------------------------------------------------------------------------
 TROUBLESHOOTING
-------------------------------------------------------------------------------
"Failed to connect to Exchange Online"
      Confirm the ExchangeOnlineManagement module is installed and that your
      account has the required admin role.

"Group '<name>' not found or not accessible"
      Check the spelling of -GroupIdentity. Try the primary SMTP address
      instead of the display name.

"CSV must contain a 'UserPrincipalName' column"
      Open the CSV and confirm the header is spelled exactly
      "UserPrincipalName" with no extra spaces.

A user shows Status = Failed
      Read the Detail column in the results CSV. A common cause is a UPN that
      does not exist in the tenant or a typo in the address.
===============================================================================