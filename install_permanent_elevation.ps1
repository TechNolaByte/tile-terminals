#Requires -Version 5.1
<#
.SYNOPSIS
    First run  -> installs a scheduled task + desktop shortcut that runs
                  tile-terminals.ps1 elevated with no UAC prompt.
    Second run -> uninstalls both.
#>

# Requires admin to create a scheduled task
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Start-Process powershell -Verb RunAs `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    exit
}

$TaskName     = "TileTerminals_Elevated"
$ScriptPath   = Join-Path $PSScriptRoot "tile-terminals.ps1"
$ShortcutPath = Join-Path ([Environment]::GetFolderPath('Desktop')) "TileTerminals.lnk"
$Desc         = "Run tile-terminals.ps1 elevated, no UAC prompt"

# ---------------------------------------------------------------------------
# TOGGLE: if task already exists, uninstall everything and exit
# ---------------------------------------------------------------------------
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Scheduled task '$TaskName' removed." -ForegroundColor Yellow

    if (Test-Path $ShortcutPath) {
        Remove-Item $ShortcutPath -Force
        Write-Host "Shortcut removed: $ShortcutPath" -ForegroundColor Yellow
    }
    Write-Host "Uninstall complete." -ForegroundColor Green
    exit 0
}

# ---------------------------------------------------------------------------
# INSTALL
# ---------------------------------------------------------------------------
if (-not (Test-Path $ScriptPath -PathType Leaf)) {
    Write-Error "tile-terminals.ps1 not found at: $ScriptPath"
    exit 1
}

# Run as the current interactive user, elevated, no UAC prompt.
# SYSTEM (Session 0) cannot touch the user's desktop windows, so we MUST
# use the logged-in user account with RunLevel Highest.
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

$action = New-ScheduledTaskAction `
    -Execute  "PowerShell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -NonInteractive -File `"$ScriptPath`""

# Interactive logon with no stored password: task fires in the user's session.
$principal = New-ScheduledTaskPrincipal `
    -UserId   $currentUser `
    -LogonType Interactive `
    -RunLevel Highest

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -MultipleInstances IgnoreNew

# No trigger: on-demand only.
# Fix for the original error: use direct params, not -InputObject, so that
# -Description is accepted (it belongs to a different parameter set).
Register-ScheduledTask `
    -TaskName    $TaskName `
    -Description $Desc `
    -Action      $action `
    -Principal   $principal `
    -Settings    $settings `
    -Force

Write-Host "Scheduled task '$TaskName' created." -ForegroundColor Green
Write-Host ("To run manually: schtasks /run /tn `"{0}`"" -f $TaskName) -ForegroundColor DarkGray

# Desktop shortcut
$shell    = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($ShortcutPath)
$shortcut.TargetPath   = "$env:SystemRoot\System32\schtasks.exe"
$shortcut.Arguments    = "/run /tn `"$TaskName`""
$shortcut.Description  = $Desc
$shortcut.IconLocation = "PowerShell.exe,0"
$shortcut.Save()

Write-Host "Shortcut created:  $ShortcutPath" -ForegroundColor Green
Write-Host ""
Write-Host "Run this script again to uninstall." -ForegroundColor DarkGray