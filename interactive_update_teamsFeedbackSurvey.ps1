<#
.SYNOPSIS
Interactive helper to assign Microsoft Teams feedback policies to users or groups.

.DESCRIPTION
- Ensures required modules (MicrosoftTeams, AzureAD) are installed and imported.
- Connects to Microsoft Teams and Azure AD.
- Lets the admin choose the target scope:
    - Single user (UPN)
    - Azure AD group (by name; script resolves ObjectId and members)
- Lets the admin choose which TeamsFeedbackPolicy to assign:
    - Global
    - Tag:Enabled
    - Tag:Disabled
    - Tag:UserChoice
- Shows a summary of WHO will be affected (lists all UPNs) and HOW (policy), then asks for confirmation.
- Applies the changes and logs every step to both the console and a transcript log file.
#>

[CmdletBinding()]
param()

#region Logging / Helpers

# Determine script directory for log placement
$scriptDir = if ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    Get-Location
}

$logFile = Join-Path $scriptDir ("TeamsFeedbackPolicy_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

Write-Host "Starting transcript logging to: $logFile"
Start-Transcript -Path $logFile -Append | Out-Null

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO ] $Message"
}

function Write-Warn {
    param([string]$Message)
    Write-Warning "[WARN ] $Message"
}

function Write-Err {
    param([string]$Message)
    Write-Error "[ERROR] $Message"
}

#endregion Logging / Helpers

#region Module Handling

function EnsureModule {
    param(
        [Parameter(Mandatory=$true)][string]$ModuleName
    )

    Write-Info "Checking for module '$ModuleName'..."
    $module = Get-Module -ListAvailable -Name $ModuleName

    if (-not $module) {
        Write-Warn "Module '$ModuleName' not found. Attempting to install from PSGallery..."
        try {
            Install-Module -Name $ModuleName -Force -AllowClobber -ErrorAction Stop
            Write-Info "Module '$ModuleName' installed successfully."
        }
        catch {
            Write-Err "Failed to install module '$ModuleName'. Error: $_"
            Write-Err "Please ensure you are running as Administrator and have access to the PowerShell Gallery."
            Stop-Transcript | Out-Null
            exit 1
        }
    }
    else {
        Write-Info "Module '$ModuleName' is already installed."
    }

    Write-Info "Importing module '$ModuleName'..."
    try {
        Import-Module $ModuleName -ErrorAction Stop
        Write-Info "Module '$ModuleName' imported successfully."
    }
    catch {
        Write-Err "Failed to import module '$ModuleName'. Error: $_"
        Stop-Transcript | Out-Null
        exit 1
    }
}

EnsureModule -ModuleName "MicrosoftTeams"
EnsureModule -ModuleName "AzureAD"

#endregion Module Handling

#region Connections

Write-Info "Connecting to Microsoft Teams..."
try {
    Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
    Write-Info "Connected to Microsoft Teams."
}
catch {
    Write-Err "Failed to connect to Microsoft Teams. Error: $_"
    Stop-Transcript | Out-Null
    exit 1
}

Write-Info "Connecting to Azure AD..."
try {
    Connect-AzureAD -ErrorAction Stop | Out-Null
    Write-Info "Connected to Azure AD."
}
catch {
    Write-Err "Failed to connect to Azure AD. Error: $_"
    Stop-Transcript | Out-Null
    exit 1
}

#endregion Connections

#region Policy Choice

function Select-FeedbackPolicy {
    Write-Host ""
    Write-Host "================= Feedback Policy Selection ================="
    Write-Host "Please choose which Teams feedback policy to assign:"
    Write-Host "1) Global      - Use tenant-wide default feedback behavior."
    Write-Host "2) Tag:Enabled - Always allow feedback surveys."
    Write-Host "3) Tag:Disabled - Disable all feedback & surveys."
    Write-Host "4) Tag:UserChoice - Let users control feedback in their client."
    Write-Host "============================================================="
    $selection = Read-Host "Enter option number (1-4)"

    switch ($selection) {
        '1' { return 'Global' }
        '2' { return 'Tag:Enabled' }
        '3' { return 'Tag:Disabled' }
        '4' { return 'Tag:UserChoice' }
        default {
            Write-Warn "Invalid selection '$selection'. Please try again."
            return Select-FeedbackPolicy
        }
    }
}

#endregion Policy Choice

#region Scope Choice

function Select-Scope {
    Write-Host ""
    Write-Host "====================== Scope Selection ======================"
    Write-Host "Select WHO will be affected by this change:"
    Write-Host "1) Single user (enter a single UPN)"
    Write-Host "2) Azure AD group (enter group display name; script resolves users)"
    Write-Host "============================================================="
    $scope = Read-Host "Enter option number (1-2)"

    switch ($scope) {
        '1' { return 'User' }
        '2' { return 'Group' }
        default {
            Write-Warn "Invalid selection '$scope'. Please try again."
            return Select-Scope
        }
    }
}

function Get-UserScopeUPNs {
    $upn = Read-Host "Enter the UPN of the user (e.g., user@domain.com)"
    $upn = $upn.Trim()

    if (-not $upn) {
        Write-Warn "UPN cannot be empty. Please try again."
        return Get-UserScopeUPNs
    }

    Write-Info "Validating user '$upn' in Azure AD..."
    try {
        $user = Get-AzureADUser -ObjectId $upn -ErrorAction Stop
    }
    catch {
        Write-Err "User '$upn' not found in Azure AD or lookup failed. Error: $_"
        $retry = Read-Host "Do you want to try entering the UPN again? (Y/N)"
        if ($retry -match '^(Y|y)') {
            return Get-UserScopeUPNs
        }
        else {
            Write-Err "Aborting by user choice."
            Stop-Transcript | Out-Null
            exit 1
        }
    }

    Write-Info "User '$($user.UserPrincipalName)' found."
    return ,$user.UserPrincipalName  # return as single-element array
}

function Get-GroupScopeUPNs {
    $groupName = Read-Host "Enter the Azure AD group display name (full or partial)"

    if (-not $groupName.Trim()) {
        Write-Warn "Group name cannot be empty. Please try again."
        return Get-GroupScopeUPNs
    }

    Write-Info "Searching for Azure AD groups matching '$groupName'..."
    $groups = Get-AzureADGroup -SearchString $groupName

    if (-not $groups) {
        Write-Err "No groups found matching '$groupName'."
        $retry = Read-Host "Do you want to try entering the group name again? (Y/N)"
        if ($retry -match '^(Y|y)') {
            return Get-GroupScopeUPNs
        }
        else {
            Write-Err "Aborting by user choice."
            Stop-Transcript | Out-Null
            exit 1
        }
    }

    if ($groups.Count -gt 1) {
        Write-Host ""
        Write-Host "Multiple groups found. Please select the correct one:"
        $index = 1
        foreach ($g in $groups) {
            Write-Host ("{0}) {1} (ObjectId: {2})" -f $index, $g.DisplayName, $g.ObjectId)
            $index++
        }
        $choice = Read-Host "Enter the number of the group you want to use"
        if (-not ($choice -as [int]) -or $choice -lt 1 -or $choice -gt $groups.Count) {
            Write-Err "Invalid group selection. Aborting."
            Stop-Transcript | Out-Null
            exit 1
        }
        $group = $groups[$choice - 1]
    }
    else {
        $group = $groups
        Write-Info ("Using group: {0} (ObjectId: {1})" -f $group.DisplayName, $group.ObjectId)
    }

    Write-Info "Resolving members of group '$($group.DisplayName)'..."
    try {
        $members = Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true
    }
    catch {
        Write-Err "Failed to retrieve members for group '$($group.DisplayName)'. Error: $_"
        Stop-Transcript | Out-Null
        exit 1
    }

    $userMembers = $members | Where-Object { $_.ObjectType -eq 'User' }

    if (-not $userMembers) {
        Write-Warn "Group '$($group.DisplayName)' has no user members."
    }

    $upns = $userMembers | Select-Object -ExpandProperty UserPrincipalName

    Write-Info ("Group '{0}' has {1} user member(s)." -f $group.DisplayName, $upns.Count)
    Write-Info "Group ObjectId (for reference/logging): $($group.ObjectId)"

    return $upns
}

#endregion Scope Choice

#region Main Flow

Write-Info "Retrieving available Teams feedback policies..."
$policies = Get-CsTeamsFeedbackPolicy
Write-Info ("Found {0} Teams feedback policy object(s)." -f $policies.Count)

$scopeSelection  = Select-Scope
$targetUPNs      = @()

switch ($scopeSelection) {
    'User' {
        Write-Info "Scope selected: Single user."
        $targetUPNs = Get-UserScopeUPNs
    }
    'Group' {
        Write-Info "Scope selected: Azure AD group."
        $targetUPNs = Get-GroupScopeUPNs
    }
}

if (-not $targetUPNs -or $targetUPNs.Count -eq 0) {
    Write-Err "No target users resolved. Nothing to do. Exiting."
    Stop-Transcript | Out-Null
    exit 1
}

$policyName = Select-FeedbackPolicy
Write-Info ("Policy selected: {0}" -f $policyName)

# Summary
Write-Host ""
Write-Host "========================== SUMMARY =========================="
Write-Host ("Scope       : {0}" -f $scopeSelection)
Write-Host ("Policy      : {0}" -f $policyName)
Write-Host ("User Count  : {0}" -f $targetUPNs.Count)
Write-Host "Users (UPN) :"
$targetUPNs | ForEach-Object { Write-Host "  - $_" }
Write-Host "============================================================="
Write-Host ""

$confirm = Read-Host "Do you want to apply this policy to ALL of the users listed above? (Y/N)"
if ($confirm -notmatch '^(Y|y)$') {
    Write-Warn "Operation cancelled by user before making any changes."
    Stop-Transcript | Out-Null
    exit 0
}

# Apply policy
Write-Info ("Applying Teams feedback policy '{0}' to {1} user(s)..." -f $policyName, $targetUPNs.Count)

$successCount = 0
$failureCount = 0

foreach ($upn in $targetUPNs) {
    Write-Info ("Granting policy '{0}' to user '{1}'..." -f $policyName, $upn)
    try {
        Grant-CsTeamsFeedbackPolicy -Identity $upn -PolicyName $policyName -ErrorAction Stop
        $successCount++
    }
    catch {
        $failureCount++
        Write-Err ("Failed to assign policy '{0}' to user '{1}'. Error: {2}" -f $policyName, $upn, $_)
    }
}

Write-Host ""
Write-Host "======================= RESULT SUMMARY ======================"
Write-Host ("Policy       : {0}" -f $policyName)
Write-Host ("Users OK     : {0}" -f $successCount)
Write-Host ("Users Failed : {0}" -f $failureCount)
Write-Host ("Log file     : {0}" -f $logFile)
Write-Host "============================================================="
Write-Host ""

Write-Info "Verifying a sample user (if available)..."
if ($targetUPNs.Count -gt 0) {
    $sample = $targetUPNs[0]
    try {
        $userInfo = Get-CsOnlineUser -Identity $sample -ErrorAction Stop | Select-Object DisplayName, UserPrincipalName, TeamsFeedbackPolicy
        Write-Host "Sample user policy after change:"
        $userInfo | Format-Table -AutoSize | Out-Host
    }
    catch {
        Write-Warn "Could not verify sample user '$sample' with Get-CsOnlineUser. Error: $_"
    }
}

Write-Info "All done."
Stop-Transcript | Out-Null
