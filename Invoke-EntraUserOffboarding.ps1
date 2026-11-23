#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, ExchangeOnlineManagement

<#
.SYNOPSIS
    Offboards an Entra ID user: disable the account, revoke sessions,
    remove group memberships, and convert the mailbox to shared.

.DESCRIPTION
    Performs the standard leaver steps for a cloud user:

      1. Disable the Entra ID account (block sign-in)
      2. Revoke refresh tokens and sign-in sessions
      3. Remove Microsoft 365 / security group memberships
         (dynamic groups and directory roles are skipped)
      4. Convert an Exchange Online user mailbox to a shared mailbox

    Each step is logged. The script supports -WhatIf so you can preview
    the actions first. It does not delete the account, remove licenses,
    wipe devices, or set mail forwarding.

.PARAMETER UserPrincipalName
    The user's UPN, for example first.last@contoso.com.

.PARAMETER LogPath
    Path to the action log. Defaults to a timestamped file under .\logs
    in the script directory.

.PARAMETER ExcludeGroupId
    Optional group object IDs that should not be removed.

.PARAMETER SkipMailboxConversion
    Perform the Entra ID steps only and leave Exchange Online unchanged.

.PARAMETER TenantId
    Optional tenant ID or domain name passed to Connect-MgGraph.

.EXAMPLE
    ./Invoke-EntraUserOffboarding.ps1 -UserPrincipalName 'alex.lee@contoso.com' -WhatIf

    Show the disable, revoke, group, and mailbox actions without changing anything.

.EXAMPLE
    ./Invoke-EntraUserOffboarding.ps1 -UserPrincipalName 'alex.lee@contoso.com'

    Run the full offboarding workflow. ConfirmImpact is High, so you will
    be prompted unless you also pass -Confirm:$false.

.NOTES
    Typical admin roles: User Administrator (or equivalent) plus
    Exchange Administrator if you convert the mailbox. Converting to a
    shared mailbox usually requires the mailbox to still be licensed at
    the moment of conversion; license removal is a common follow-up
    step and is intentionally left out of this script.
.LINK
    https://github.com/samtravis-sysadmin/entra-offboarding

#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$UserPrincipalName,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath,

    [Parameter()]
    [Guid[]]$ExcludeGroupId = @(),

    [Parameter()]
    [switch]$SkipMailboxConversion,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-OffboardingLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $line = '{0:yyyy-MM-dd HH:mm:ss}  [{1}]  {2}' -f (Get-Date), $Level, $Message
    Add-Content -LiteralPath $script:ResolvedLogPath -Value $line -Encoding utf8

    switch ($Level) {
        'ERROR' { Write-Warning $line }
        'WARN' { Write-Warning $Message }
        default { Write-Host $line }
    }
}

function Initialize-LogFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $directory = Split-Path -Parent $Path
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        $null = New-Item -ItemType Directory -Path $directory -Force
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        $null = New-Item -ItemType File -Path $Path -Force
    }
}

function Test-HasGraphScopes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Required
    )

    $context = Get-MgContext
    if (-not $context) {
        return $false
    }

    $granted = @($context.Scopes | ForEach-Object { ($_ -split '/')[-1] })
    foreach ($scope in $Required) {
        $shortName = ($scope -split '/')[-1]
        if ($shortName -notin $granted) {
            return $false
        }
    }
    return $true
}

function Connect-OffboardingGraph {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$TenantId
    )

    $requiredScopes = @(
        'User.ReadWrite.All',
        'User.RevokeSessions.All',
        'GroupMember.ReadWrite.All',
        'Directory.Read.All'
    )

    if (Test-HasGraphScopes -Required $requiredScopes) {
        Write-OffboardingLog -Message "Already connected to Microsoft Graph as $((Get-MgContext).Account)."
        return
    }

    $connectParams = @{
        Scopes    = $requiredScopes
        NoWelcome = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $connectParams['TenantId'] = $TenantId
    }

    Write-OffboardingLog -Message 'Connecting to Microsoft Graph.'
    Connect-MgGraph @connectParams
}

function Connect-OffboardingExchange {
    [CmdletBinding()]
    param()

    $connections = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
    if ($connections.Count -gt 0) {
        Write-OffboardingLog -Message 'Already connected to Exchange Online.'
        return
    }

    Write-OffboardingLog -Message 'Connecting to Exchange Online.'
    Connect-ExchangeOnline -ShowBanner:$false
}

function Get-RemovableGroupMembership {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter()]
        [Guid[]]$ExcludeGroupId = @()
    )

    $exclude = @($ExcludeGroupId | ForEach-Object { $_.ToString() })
    $memberships = @(Get-MgUserMemberOf -UserId $UserId -All)
    $removable = [System.Collections.Generic.List[object]]::new()

    foreach ($item in $memberships) {
        $additional = $item.AdditionalProperties
        $odataType = $null
        if ($additional -and $additional.ContainsKey('@odata.type')) {
            $odataType = [string]$additional['@odata.type']
        }

        if ($odataType -ne '#microsoft.graph.group') {
            $displayName = if ($additional -and $additional.ContainsKey('displayName')) { [string]$additional['displayName'] } else { $item.Id }
            Write-OffboardingLog -Message "Skipping non-group membership '$displayName' ($odataType)."
            continue
        }

        $displayName = if ($additional -and $additional.ContainsKey('displayName')) { [string]$additional['displayName'] } else { $item.Id }
        $groupTypes = @()
        if ($additional -and $additional.ContainsKey('groupTypes') -and $null -ne $additional['groupTypes']) {
            $groupTypes = @($additional['groupTypes'])
        }

        if ($groupTypes -contains 'DynamicMembership') {
            Write-OffboardingLog -Message "Skipping dynamic group '$displayName' ($($item.Id))."
            continue
        }

        if ($item.Id -in $exclude) {
            Write-OffboardingLog -Message "Skipping excluded group '$displayName' ($($item.Id))."
            continue
        }

        $removable.Add([pscustomobject]@{
                Id          = $item.Id
                DisplayName = $displayName
            })
    }

    return $removable
}

# --- Main ------------------------------------------------------------------

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path -Path $PSScriptRoot -ChildPath "logs/EntraOffboarding-$timestamp.log"
}
if (-not [System.IO.Path]::IsPathRooted($LogPath)) {
    $LogPath = Join-Path -Path (Get-Location) -ChildPath $LogPath
}
$script:ResolvedLogPath = $LogPath
Initialize-LogFile -Path $script:ResolvedLogPath

$summary = [pscustomobject]@{
    UserPrincipalName     = $UserPrincipalName
    UserId                = $null
    AccountDisabled       = $false
    SessionsRevoked       = $false
    GroupsRemoved         = [System.Collections.Generic.List[string]]::new()
    GroupsFailed          = [System.Collections.Generic.List[string]]::new()
    MailboxConverted      = $false
    MailboxStatus         = $null
    WhatIf                = [bool]$WhatIfPreference
    LogPath               = $script:ResolvedLogPath
    Errors                = [System.Collections.Generic.List[string]]::new()
}

Write-OffboardingLog -Message "Starting Entra ID offboarding for '$UserPrincipalName'."
if ($WhatIfPreference) {
    Write-OffboardingLog -Message 'Running in WhatIf mode. No changes will be made.'
}

try {
    Connect-OffboardingGraph -TenantId $TenantId

    $user = Get-MgUser -UserId $UserPrincipalName -Property 'Id,DisplayName,UserPrincipalName,AccountEnabled' -ErrorAction Stop
    $summary.UserId = $user.Id
    Write-OffboardingLog -Message "Resolved '$($user.DisplayName)' ($($user.UserPrincipalName)), object ID $($user.Id). AccountEnabled=$($user.AccountEnabled)."

    if ($user.AccountEnabled -eq $false) {
        Write-OffboardingLog -Level 'WARN' -Message 'Account is already disabled. Continuing with the remaining steps.'
        $summary.AccountDisabled = $true
    }
    elseif ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Disable Entra ID account (block sign-in)')) {
        try {
            Update-MgUser -UserId $user.Id -BodyParameter @{ accountEnabled = $false }
            $summary.AccountDisabled = $true
            Write-OffboardingLog -Message 'Disabled the account (sign-in blocked).'
        }
        catch {
            $summary.Errors.Add("Disable account: $($_.Exception.Message)")
            Write-OffboardingLog -Level 'ERROR' -Message "Failed to disable the account: $($_.Exception.Message)"
        }
    }
    else {
        Write-OffboardingLog -Message "WHATIF: would disable '$UserPrincipalName'."
    }

    if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Revoke sign-in sessions and refresh tokens')) {
        try {
            $null = Revoke-MgUserSignInSession -UserId $user.Id
            $summary.SessionsRevoked = $true
            Write-OffboardingLog -Message 'Revoked sign-in sessions and refresh tokens.'
        }
        catch {
            $summary.Errors.Add("Revoke sessions: $($_.Exception.Message)")
            Write-OffboardingLog -Level 'ERROR' -Message "Failed to revoke sessions: $($_.Exception.Message)"
        }
    }
    else {
        Write-OffboardingLog -Message "WHATIF: would revoke sessions for '$UserPrincipalName'."
    }

    $groups = Get-RemovableGroupMembership -UserId $user.Id -ExcludeGroupId $ExcludeGroupId
    if ($groups.Count -eq 0) {
        Write-OffboardingLog -Message 'No removable (non-dynamic) group memberships were found.'
    }

    foreach ($group in $groups) {
        $label = "$($group.DisplayName) ($($group.Id))"
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Remove from group $label")) {
            try {
                Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop
                $summary.GroupsRemoved.Add($label)
                Write-OffboardingLog -Message "Removed from group $label."
            }
            catch {
                $summary.GroupsFailed.Add($label)
                $summary.Errors.Add("Remove group $label`: $($_.Exception.Message)")
                Write-OffboardingLog -Level 'ERROR' -Message "Failed to remove from group $label`: $($_.Exception.Message)"
            }
        }
        else {
            Write-OffboardingLog -Message "WHATIF: would remove '$UserPrincipalName' from $label."
        }
    }

    if ($SkipMailboxConversion) {
        $summary.MailboxStatus = 'Skipped by operator'
        Write-OffboardingLog -Message 'Skipping mailbox conversion because -SkipMailboxConversion was set.'
    }
    else {
        try {
            Connect-OffboardingExchange
            $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
            $summary.MailboxStatus = [string]$mailbox.RecipientTypeDetails

            if ($mailbox.RecipientTypeDetails -eq 'SharedMailbox') {
                Write-OffboardingLog -Message 'Mailbox is already a shared mailbox.'
                $summary.MailboxConverted = $true
            }
            elseif ($mailbox.RecipientTypeDetails -ne 'UserMailbox') {
                $message = "Mailbox type is '$($mailbox.RecipientTypeDetails)', not UserMailbox. Leaving it unchanged."
                $summary.MailboxStatus = $message
                Write-OffboardingLog -Level 'WARN' -Message $message
            }
            elseif ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Convert Exchange Online mailbox to shared')) {
                Set-Mailbox -Identity $UserPrincipalName -Type Shared -ErrorAction Stop
                $summary.MailboxConverted = $true
                $summary.MailboxStatus = 'SharedMailbox'
                Write-OffboardingLog -Message 'Converted the mailbox to a shared mailbox.'
            }
            else {
                Write-OffboardingLog -Message "WHATIF: would convert the mailbox for '$UserPrincipalName' to shared."
            }
        }
        catch {
            $summary.Errors.Add("Mailbox conversion: $($_.Exception.Message)")
            $summary.MailboxStatus = "Error: $($_.Exception.Message)"
            Write-OffboardingLog -Level 'ERROR' -Message "Mailbox conversion failed: $($_.Exception.Message)"
        }
    }
}
catch {
    $summary.Errors.Add($_.Exception.Message)
    Write-OffboardingLog -Level 'ERROR' -Message "Offboarding stopped: $($_.Exception.Message)"
    throw
}
finally {
    Write-OffboardingLog -Message ("Finished. Disabled={0}; SessionsRevoked={1}; GroupsRemoved={2}; GroupsFailed={3}; Mailbox={4}. Log: '{5}'." -f $summary.AccountDisabled, $summary.SessionsRevoked, $summary.GroupsRemoved.Count, $summary.GroupsFailed.Count, $summary.MailboxStatus, $script:ResolvedLogPath)
}

$summary
