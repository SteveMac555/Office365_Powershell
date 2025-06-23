[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter the number of days of inactivity.")]
    [ValidateRange(1, 3650)]
    [int]$InactiveDays,

    [Parameter(Mandatory = $false, HelpMessage = "Path to export the CSV file.")]
    [string]$OutputPath = "C:\temp\InactiveLicensedUsers.csv"
)

try {
    $OutputDirectory = Split-Path -Path $OutputPath -Parent

    if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) {
        if (-not (Test-Path -Path $OutputDirectory)) {
            Write-Host "Output directory '$OutputDirectory' does not exist. Creating it..." -ForegroundColor Yellow
            New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        }
    }

    $RequiredScopes = @("User.Read.All", "Reports.Read.All")
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
    Connect-MgGraph -Scopes $RequiredScopes -NoWelcome

    $context = Get-MgContext
    if ($null -eq $context.Account) {
        throw "Failed to connect to Microsoft Graph. Please check your credentials and permissions."
    }
    Write-Host "Successfully connected as $($context.Account)" -ForegroundColor Green

    $cutoffDate = (Get-Date).AddDays(-$InactiveDays)
    Write-Host "Searching for licensed users inactive since $($cutoffDate.ToString('yyyy-MM-dd'))..."

    $properties = "id,displayName,userPrincipalName,userType,createdDateTime,signInActivity"
    $licensedUsers = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -Property $properties -All -ConsistencyLevel eventual -CountVariable userCount

    if ($null -eq $licensedUsers) {
        Write-Warning "No licensed users found in the tenant."
        return
    }

    $totalUsers = $licensedUsers.Count
    Write-Host "Found $totalUsers licensed users. Analyzing sign-in data..."

    $inactiveUserReport = @()
    $count = 0

    foreach ($user in $licensedUsers) {
        $count++
        Write-Progress -Activity "Analyzing User Sign-In Data" -Status "Processing user $count of $totalUsers ($($user.UserPrincipalName))" -PercentComplete (($count / $totalUsers) * 100)

        $lastSignIn = $user.SignInActivity.LastSignInDateTime
        $isInactive = $false
        $lastSignInStatus = ""

        if ($null -eq $lastSignIn) {
            # User has never had an interactive sign-in recorded.
            if ($user.CreatedDateTime -lt $cutoffDate) {
                $isInactive = $true
                $lastSignInStatus = "Never or No Data"
            }
        }
        elseif ($lastSignIn -lt $cutoffDate) {
            # Last sign-in is older than our cutoff date.
            $isInactive = $true
            $lastSignInStatus = $lastSignIn.ToString("yyyy-MM-dd HH:mm:ss")
        }

        if ($isInactive) {
            $inactiveUserReport += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName       = $user.DisplayName
                LastSignIn        = $lastSignInStatus
                AccountCreated    = $user.CreatedDateTime.ToString("yyyy-MM-dd")
                UserType          = $user.UserType
            }
        }
    }

    Write-Progress -Activity "Analysis Completed." -Completed

    if ($inactiveUserReport.Count -gt 0) {
        Write-Host "`nFound $($inactiveUserReport.Count) inactive licensed users." -ForegroundColor Green
        $inactiveUserReport | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Report successfully exported to: $(Resolve-Path -Path $OutputPath)" -ForegroundColor Green
    }
    else {
        Write-Host "`nNo inactive licensed users found matching the criteria." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    if (Get-MgContext) {
        Write-Host "Disconnecting from Microsoft Graph..."
        Disconnect-MgGraph
    }
}