<#
.SYNOPSIS
    Prevents PC from sleeping during long import operations.

.DESCRIPTION
    Sends synthetic F15 keypresses every 45 seconds to reset Windows sleep timer.
    Run in a separate PowerShell window while import script runs.

.USAGE
    .\Keep-Awake.ps1
    # Press Ctrl+C to stop
#>

$wsh = New-Object -ComObject Wscript.Shell
Write-Host "Keep-awake active. Press Ctrl+C to stop." -ForegroundColor Green

while ($true) {
    $wsh.SendKeys('+{F15}')
    Start-Sleep -Seconds 45
}
