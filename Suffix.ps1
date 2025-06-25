[CmdletBinding(SupportsShouldProcess = $true)] # Enables -WhatIf and -Confirm parameters
param (
    [Parameter(Mandatory = $true)]
    [string]$TargetOU,

    [Parameter(Mandatory = $true)]
    [AllowEmptyString()] # Explicitly allow "" as a value
    [string]$OldUPNSuffix,

    [Parameter(Mandatory = $true)]
    [string]$NewUPNSuffix
)

if (-not (Get-Module ActiveDirectory -ErrorAction SilentlyContinue)) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        Write-Error "Active Directory module is not available. Please install RSAT-AD-Tools or run this on a Domain Controller."
        exit 1
    }
}

Write-Warning "IMPORTANT: Ensure the new UPN suffix '$NewUPNSuffix' is ALREADY registered in your ON-PREMISES AD forest (Active Directory Domains and Trusts) AND VERIFIED in Azure AD."

$usersToProcess = @()

try {
    if ([string]::IsNullOrWhiteSpace($OldUPNSuffix)) {
        Write-Verbose "OldUPNSuffix is blank. Searching for users in '$TargetOU' with a missing or incomplete UPN (no '@' symbol)."
        # Get all users in the OU and filter in PowerShell for more complex logic.
        $allUsersInOU = Get-ADUser -Filter * -SearchBase $TargetOU -SearchScope Subtree -Properties SamAccountName, UserPrincipalName -ErrorAction Stop
        $usersToProcess = $allUsersInOU | Where-Object { -not $_.UserPrincipalName -or $_.UserPrincipalName -notlike '*@*' }
    }
    else {
        Write-Verbose "Attempting to retrieve users from OU: '$TargetOU' whose UPN ends with '@$OldUPNSuffix'."
        $filterString = "UserPrincipalName -like '*@$($OldUPNSuffix)'"
        $usersToProcess = Get-ADUser -Filter $filterString -SearchBase $TargetOU -SearchScope Subtree -Properties SamAccountName, UserPrincipalName -ErrorAction Stop
    }
}
catch {
    Write-Error "Failed to retrieve users from OU '$TargetOU'. Error: $($_.Exception.Message)"
    exit 1
}


if ($usersToProcess.Count -eq 0) {
    if ([string]::IsNullOrWhiteSpace($OldUPNSuffix)) {
        Write-Host "No users found in OU '$TargetOU' with a blank or incomplete UPN. Exiting." -ForegroundColor Yellow
    }
    else {
        Write-Host "No users found in OU '$TargetOU' with the UPN suffix '$OldUPNSuffix'. Exiting." -ForegroundColor Yellow
    }
    exit 0
}

Write-Host "Found $($usersToProcess.Count) user(s) matching the criteria in OU '$TargetOU' to potentially update."
Write-Host "Review the users. If -WhatIf is not used, changes will be attempted." -ForegroundColor Magenta
Write-Host "If Set-ADUser prompts, run this script with -Confirm:`$false to suppress those." -ForegroundColor Magenta

foreach ($user in $usersToProcess) {
    $currentUPN = $user.UserPrincipalName
    $usernamePart = $null

    if ($currentUPN -like '*@*') {
        # Standard case: user@domain.com
        $usernamePart = ($currentUPN -split '@', 2)[0]
    }
    elseif (-not [string]::IsNullOrWhiteSpace($currentUPN)) {
        # Case where UPN is present but has no suffix, e.g., "john.smith"
        $usernamePart = $currentUPN
        Write-Verbose "User $($user.SamAccountName) has a UPN ('$currentUPN') without a suffix. Using it as the username part."
    }
    else {
        # Fallback case: UPN is null or empty. Use SamAccountName as the most reliable identifier.
        $usernamePart = $user.SamAccountName
        Write-Verbose "User $($user.SamAccountName) has a blank UPN. Using SamAccountName ('$($user.SamAccountName)') as the username part."
    }

    if (-not $usernamePart) {
        Write-Warning "Could not determine a username part for user $($user.SamAccountName) (DN: $($user.DistinguishedName)). Skipping."
        continue
    }

    $newProposedUPN = "$($usernamePart)@$($NewUPNSuffix)"

    if ($currentUPN -eq $newProposedUPN) {
        Write-Verbose "User $($user.SamAccountName) (UPN: $currentUPN) already has the target UPN. Skipping."
        continue
    }

    Write-Host "Processing User: $($user.SamAccountName) (DN: $($user.DistinguishedName))" -ForegroundColor Cyan
    Write-Host "  Current UPN: $(if ([string]::IsNullOrEmpty($currentUPN)) {'<BLANK>'} else {$currentUPN})"
    Write-Host "  Proposed New UPN: $newProposedUPN"

    try {
        Set-ADUser -Identity $user -UserPrincipalName $newProposedUPN -ErrorAction Stop

        if (-not $PSCmdlet.ShouldProcess("User: $($user.SamAccountName)", "Update UPN")) {
            # This block is only entered if -WhatIf is used. Set-ADUser already showed its "What if:" message.
        } else {
             Write-Host "  SUCCESS: UPN for $($user.SamAccountName) updated to $newProposedUPN." -ForegroundColor Green
        }
    }
    catch {
        Write-Error "  FAILED to update UPN for $($user.SamAccountName). Error: $($_.Exception.Message)"
    }
    Write-Host ("-" * 50)
}

Write-Host "Script finished."