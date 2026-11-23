# Entra ID user offboarding

PowerShell 7 script that walks a leaver through the standard Entra ID and Exchange Online offboarding steps:

1. Disable the account (block sign-in)
2. Revoke sign-in sessions and refresh tokens
3. Remove group memberships (dynamic groups and directory roles are skipped)
4. Convert the Exchange Online mailbox to a shared mailbox

Every action is written to a log. Use `-WhatIf` to preview the work first.

This is a portfolio example. Offboarding changes a real account. Read the script, confirm you have the right person, and run `-WhatIf` before you run it for real.

## Requirements

- PowerShell 7
- Modules:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

- Entra ID role that can disable users, revoke sessions, and edit group membership (for example, User Administrator)
- Exchange Administrator, or equivalent, if you convert the mailbox
- Graph scopes requested by the script:
  - `User.ReadWrite.All`
  - `User.RevokeSessions.All`
  - `GroupMember.ReadWrite.All`
  - `Directory.Read.All`

Sign-in is interactive. The script stores no secrets.

## Usage

Preview first:

```powershell
./Invoke-EntraUserOffboarding.ps1 -UserPrincipalName 'alex.lee@contoso.com' -WhatIf
```

Run the full workflow (you will be prompted to confirm; `ConfirmImpact` is High):

```powershell
./Invoke-EntraUserOffboarding.ps1 -UserPrincipalName 'alex.lee@contoso.com'
```

Leave the mailbox unchanged, or keep specific groups:

```powershell
./Invoke-EntraUserOffboarding.ps1 -UserPrincipalName 'alex.lee@contoso.com' -SkipMailboxConversion
./Invoke-EntraUserOffboarding.ps1 -UserPrincipalName 'alex.lee@contoso.com' -ExcludeGroupId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
```

The script prints a summary object and writes `logs/EntraOffboarding-<timestamp>.log` unless you set `-LogPath`.

## What this script does not do

These are common follow-up steps in many organizations and are left out on purpose:

- Delete the account
- Remove Microsoft 365 licenses (convert the mailbox while the user is still licensed, then remove licenses according to your process)
- Wipe or retire devices
- Set an out-of-office message or mail forwarding
- Transfer OneDrive or Teams ownership
- Remove directory role assignments (those are logged as skipped)

Role-assignable groups may fail if the signed-in admin does not have permission. Those failures are logged and the rest of the run continues.

## Parameters

| Parameter | Description |
| --- | --- |
| `-UserPrincipalName` | The account to offboard (required) |
| `-LogPath` | Action log path |
| `-ExcludeGroupId` | One or more group object IDs to leave in place |
| `-SkipMailboxConversion` | Skip the Exchange Online step |
| `-TenantId` | Optional tenant ID or domain for `Connect-MgGraph` |
| `-WhatIf` | Show the planned actions without changing anything |
