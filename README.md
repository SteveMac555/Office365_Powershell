# Office365 & Windows Powershell
Collection of PowerShell Scripts to Manage/Use with Offic365 &amp; Entra and Windows, focused on a cloud-first approach.


.\Get-InactiveLicensedUsers.ps1 -InactiveDays 180 -OutputPath "users.csv"

.\Remove-ADComputers.ps1 -CsvPath "pcs.csv" -WhatIf -Confirm

.\Suffix.ps1 -TargetOU "OU=NewStaff,DC=contoso,DC=com" -OldUPNSuffix "" -NewUPNSuffix "staff.contoso.com" -Confirm:$false -Verbose
