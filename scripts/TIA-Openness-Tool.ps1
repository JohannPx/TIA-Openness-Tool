# PowerShell Script: TIA Portal Openness Tool
# Multi-version TIA Portal DataBlock Exporter
# Version: 1.0.0 - WPF GUI
# Author: JPR
# Date: 2026-02-27

# =================== GENERAL SETTINGS ===================
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Console UTF-8
try { chcp 65001 > $null } catch {}
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# STA check (WPF requires STA thread)
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Start-Process powershell.exe -ArgumentList "-Sta -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -NoNewWindow -Wait
    exit
}

# Show startup message before hiding console
Write-Host ""
Write-Host "  TIA Portal Openness Tool" -ForegroundColor Cyan
Write-Host "  Chargement de l'interface..." -ForegroundColor Gray
Write-Host ""

# Brief pause so the user can read the startup message
Start-Sleep -Seconds 1

# Hide the console window (only the WPF GUI will be visible)
Add-Type -Name Win32 -Namespace Native -MemberDefinition @'
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
$consoleHwnd = [Native.Win32]::GetConsoleWindow()
if ($consoleHwnd -ne [IntPtr]::Zero) {
    [Native.Win32]::ShowWindow($consoleHwnd, 0) | Out-Null  # 0 = SW_HIDE
}

# =================== MODULE LOADING ===================
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulesDir = Join-Path $ScriptDir "modules"

if (Test-Path $modulesDir) {
    $moduleOrder = @(
        "AppState.ps1"
        "Localization.ps1"
        "TiaVersions.ps1"
        "TiaConnection.ps1"
        "TiaDataBlocks.ps1"
        "TiaExport.ps1"
        "TiaExportTable.ps1"
        "UIHelpers.ps1"
        "UI.ps1"
    )
    foreach ($mod in $moduleOrder) {
        $modPath = Join-Path $modulesDir $mod
        if (Test-Path $modPath) {
            . $modPath
        }
    }
}

# =================== LAUNCH ===================
try {
    $window = Initialize-MainWindow
    $window.ShowDialog() | Out-Null
} catch {
    $errMsg = $_.Exception.Message
    $inner = $_.Exception.InnerException
    while ($inner) {
        $errMsg += "`n-> $($inner.Message)"
        $inner = $inner.InnerException
    }
    try {
        [System.Windows.MessageBox]::Show(
            "$(T 'MsgError'): $errMsg`n`n$($_.ScriptStackTrace)",
            (T "MsgError"), "OK", "Error")
    } catch {
        Write-Host "Erreur: $errMsg" -ForegroundColor Red
        Write-Host $_.ScriptStackTrace
    }
}
