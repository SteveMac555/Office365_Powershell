[CmdletBinding(SupportsShouldProcess = $true)] # Enables -WhatIf and -Confirm
param (
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [string]$HostnameColumnName = "ComputerName",

    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO" # INFO, WARNING, ERROR
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$Timestamp [$Level] $Message"
    Write-Host $LogEntry
    if ($LogPath) {
        try {
            Add-Content -Path $LogPath -Value $LogEntry -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to write to log file '$LogPath': $($_.Exception.Message)"
        }
    }
}

if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
    Write-Log -Message "ActiveDirectory PowerShell module is not available. Please install RSAT for AD DS or run this on a Domain Controller." -Level "ERROR"
    exit 1
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Log -Message "Failed to import ActiveDirectory module: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}

if (-not (Test-Path -Path $CsvPath -PathType Leaf)) {
    Write-Log -Message "CSV file not found at path: $CsvPath" -Level "ERROR"
    exit 1
}

try {
    $computersToRemove = Import-Csv -Path $CsvPath -ErrorAction Stop
}
catch {
    Write-Log -Message "Failed to import CSV file '$CsvPath': $($_.Exception.Message)" -Level "ERROR"
    exit 1
}

if ($null -eq $computersToRemove -or $computersToRemove.Count -eq 0) {
    Write-Log -Message "No computer names found in the CSV file or the file is empty." -Level "WARNING"
    exit 0
}

if (-not ($computersToRemove[0].PSObject.Properties.Name -contains $HostnameColumnName)) {
    Write-Log -Message "The specified hostname column '$HostnameColumnName' was not found in the CSV file." -Level "ERROR"
    Write-Log -Message "Available columns are: $($computersToRemove[0].PSObject.Properties.Name -join ', ')" -Level "INFO"
    exit 1
}

Write-Log -Message "Starting process to remove computer objects..." -Level "INFO"
$ProgressCount = 0
$TotalComputers = $computersToRemove.Count

foreach ($row in $computersToRemove) {
    $ProgressCount++
    $computerName = $row.$HostnameColumnName.Trim()

    if ([string]::IsNullOrWhiteSpace($computerName)) {
        Write-Log -Message "Skipping empty or whitespace hostname in CSV row $ProgressCount." -Level "WARNING"
        continue
    }

    Write-Progress -Activity "Processing AD Computer Objects" -Status "Processing $computerName ($ProgressCount of $TotalComputers)" -PercentComplete (($ProgressCount / $TotalComputers) * 100)

    Write-Log -Message "Attempting to find AD computer object: $computerName" -Level "INFO"
    try {
        $adComputer = Get-ADComputer -Identity $computerName -ErrorAction Stop
        
        if ($adComputer) {
            Write-Log -Message "Found AD Computer: $($adComputer.DistinguishedName)" -Level "INFO"
            
            try {
                Remove-ADComputer -Identity $adComputer.DistinguishedName -ErrorAction Stop # -Confirm is handled by [CmdletBinding]
                Write-Log -Message "Successfully removed (or would remove if -WhatIf) computer object: $($adComputer.DistinguishedName)" -Level "INFO"
            }
            catch {
                Write-Log -Message "Failed to remove computer object '$($adComputer.DistinguishedName)': $($_.Exception.Message)" -Level "ERROR"
            }
        }
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Log -Message "Computer object '$computerName' not found in Active Directory." -Level "WARNING"
    }
    catch {
        Write-Log -Message "An error occurred while processing '$computerName': $($_.Exception.Message)" -Level "ERROR"
    }
}

Write-Progress -Activity "Processing AD Computer Objects" -Completed
Write-Log -Message "Finished processing computer objects." -Level "INFO"