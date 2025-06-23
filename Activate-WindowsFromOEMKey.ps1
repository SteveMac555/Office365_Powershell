Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [string]$Type = "INFO"
    )
    $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [$($MyInvocation.MyCommand.Name)] - $Type - $Message"
    Write-Host $LogEntry
}

try {
    Write-Log -Message "Starting OEM Activation process."
    Write-Log -Message "Checking current Windows activation status..."
    $activation = Get-CimInstance -ClassName SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.ApplicationID -eq "55c92734-d682-4d71-983e-d6ec3f16059f" } | Select-Object -First 1

    if ($activation.LicenseStatus -eq 1) {
        Write-Log -Message "Windows is already activated. No action needed." -Type "SUCCESS"
        Exit 0 # Exit successfully
    }
    else {
        Write-Log -Message "Windows is not activated. Proceeding to find OEM key."
    }

    Write-Log -Message "Querying firmware for OEM Product Key..."
    $oemKey = (Get-CimInstance -ClassName SoftwareLicensingService).OA3xOriginalProductKey

    if (-not [string]::IsNullOrWhiteSpace($oemKey)) {
        $maskedKey = $oemKey -replace '.{5}-.{5}-.{5}-.{5}', 'XXXXX-XXXXX-XXXXX-XXXXX'
        Write-Log -Message "Found OEM Product Key: $maskedKey"

        Write-Log -Message "Attempting to install the OEM product key..."
        $ipkProcess = Start-Process -FilePath "cscript.exe" -ArgumentList "//nologo C:\Windows\System32\slmgr.vbs /ipk $oemKey" -Wait -PassThru -WindowStyle Hidden
        
        if ($ipkProcess.ExitCode -eq 0) {
            Write-Log -Message "Successfully installed product key." -Type "SUCCESS"
        }
        else {
            Write-Log -Message "Failed to install the product key. SLMGR exit code: $($ipkProcess.ExitCode)." -Type "ERROR"
        }

        Write-Log -Message "Attempting to activate Windows online..."
        $atoProcess = Start-Process -FilePath "cscript.exe" -ArgumentList "//nologo C:\Windows\System32\slmgr.vbs /ato" -Wait -PassThru -WindowStyle Hidden

        if ($atoProcess.ExitCode -eq 0) {
            Write-Log -Message "Windows successfully activated online." -Type "SUCCESS"
        }
        else {
            Write-Log -Message "Online activation failed. SLMGR exit code: $($atoProcess.ExitCode). Windows will attempt to activate automatically later." -Type "WARNING"
        }
    }
    else {
        Write-Log -Message "No OEM Product Key found in the system firmware. Cannot proceed." -Type "WARNING"
    }

    Write-Log -Message "OEM Activation script finished."
    Exit 0

}
catch {
    Write-Log -Message "A critical error occurred: $($_.Exception.Message)" -Type "FATAL"
    Exit 1
}