# TiaConnection.ps1 - Scan, connect, disconnect to TIA Portal instances
# Ported from TIA_APP_V1\Services\TiaConnectionService.cs and Functions\ConnectionFunction.cs

# C# helper for P/Invoke (EnumWindows callback requires compiled delegate)
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Management;

public static class TiaProcessHelper {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int maxCount);
    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc proc, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    public static string GetProjectNameFromWindowTitle(int processId) {
        string result = string.Empty;
        EnumWindows((hWnd, lParam) => {
            int wPid;
            GetWindowThreadProcessId(hWnd, out wPid);
            if (wPid == processId) {
                var sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, sb.Capacity);
                string title = sb.ToString();
                if (!string.IsNullOrEmpty(title) && title.Contains("TIA Portal")) {
                    var m = Regex.Match(title, @"\[([^\]]+)\]");
                    if (m.Success) { result = m.Groups[1].Value; return false; }
                    string cleaned = Regex.Replace(title, @"TIA Portal|V\d+|\-", "").Trim();
                    if (!string.IsNullOrEmpty(cleaned)) { result = cleaned; return false; }
                }
            }
            return true;
        }, IntPtr.Zero);
        return result;
    }

    public static string GetProjectNameFromCommandLine(int processId) {
        try {
            using (var searcher = new ManagementObjectSearcher(
                "SELECT CommandLine FROM Win32_Process WHERE ProcessId = " + processId)) {
                foreach (ManagementObject obj in searcher.Get()) {
                    string cmd = obj["CommandLine"] != null ? obj["CommandLine"].ToString() : null;
                    if (!string.IsNullOrEmpty(cmd)) {
                        var m = Regex.Match(cmd, @"([^\\]+)\.ap\d+", RegexOptions.IgnoreCase);
                        if (m.Success) return m.Groups[1].Value;
                    }
                }
            }
        } catch {}
        return string.Empty;
    }
}
'@ -ReferencedAssemblies @('System.Management') -ErrorAction SilentlyContinue

function Get-ProjectNameFromProcess {
    param([int]$ProcessId)

    # Method 1: Command line (WMI)
    $name = [TiaProcessHelper]::GetProjectNameFromCommandLine($ProcessId)
    if (-not [string]::IsNullOrEmpty($name)) { return $name }

    # Method 2: Window title (P/Invoke)
    $name = [TiaProcessHelper]::GetProjectNameFromWindowTitle($ProcessId)
    if (-not [string]::IsNullOrEmpty($name)) { return $name }

    return ""
}

function Invoke-TiaScan {
    $state = Get-AppState
    if (-not $state.DllLoaded) {
        throw (T "MsgDllRequired")
    }

    $instances = @()
    $processes = [Siemens.Engineering.TiaPortal]::GetProcesses()

    foreach ($process in $processes) {
        $projectName = Get-ProjectNameFromProcess -ProcessId $process.Id
        $version = Get-TiaVersionFromProcess -ProcessId $process.Id

        $displayText = if ([string]::IsNullOrEmpty($projectName)) {
            "TIA Portal - PID: $($process.Id)"
        } else {
            "TIA Portal - $projectName"
        }

        if ($version) { $displayText += " ($version)" }

        $instances += @{
            ProcessId   = $process.Id
            DisplayText = $displayText
            ProjectName = $projectName
            Version     = $version
        }
    }

    Set-AppStateValue -Key "TiaInstances" -Value $instances
    return $instances
}

function Connect-TiaInstance {
    param([int]$ProcessId)

    $state = Get-AppState
    if (-not $state.DllLoaded) {
        return @{ Success = $false; Message = T "MsgDllRequired" }
    }

    try {
        $processes = [Siemens.Engineering.TiaPortal]::GetProcesses()
        $targetProcess = $null
        foreach ($p in $processes) {
            if ($p.Id -eq $ProcessId) { $targetProcess = $p; break }
        }

        if (-not $targetProcess) {
            return @{ Success = $false; Message = T "MsgProcessNotFound" }
        }

        $tiaPortal = $targetProcess.Attach()

        if ($tiaPortal.Projects.Count -eq 0) {
            return @{ Success = $false; Message = T "MsgNoProject" }
        }

        $project = $tiaPortal.Projects[0]

        # Enumerate PLC devices
        $plcList = @()
        foreach ($device in $project.Devices) {
            foreach ($deviceItem in $device.DeviceItems) {
                try {
                    $swContainer = $deviceItem.GetService([Siemens.Engineering.HW.Features.SoftwareContainer])
                    if ($swContainer -and $swContainer.Software -is [Siemens.Engineering.SW.PlcSoftware]) {
                        $plcList += $swContainer.Software
                    }
                } catch {}
            }
        }

        if ($plcList.Count -eq 0) {
            return @{ Success = $false; Message = T "MsgNoPlc" }
        }

        # Update state
        Set-AppStateValue -Key "TiaPortal" -Value $tiaPortal
        Set-AppStateValue -Key "CurrentProject" -Value $project
        Set-AppStateValue -Key "PlcSoftwareList" -Value $plcList
        Set-AppStateValue -Key "IsConnected" -Value $true
        Set-AppStateValue -Key "ConnectedProcessId" -Value $ProcessId
        Set-AppStateValue -Key "ProjectName" -Value $project.Name

        return @{
            Success     = $true
            Message     = (T "MsgConnectOk") -f $project.Name, $plcList.Count
            ProjectName = $project.Name
            PlcCount    = $plcList.Count
        }
    } catch {
        return @{ Success = $false; Message = (T "MsgConnectError") -f $_.Exception.Message }
    }
}

function Disconnect-TiaInstance {
    $state = Get-AppState
    try {
        if ($state.TiaPortal) {
            $state.TiaPortal.Dispose()
        }
    } catch {}

    Set-AppStateValue -Key "TiaPortal" -Value $null
    Set-AppStateValue -Key "CurrentProject" -Value $null
    Set-AppStateValue -Key "PlcSoftwareList" -Value @()
    Set-AppStateValue -Key "IsConnected" -Value $false
    Set-AppStateValue -Key "ConnectedProcessId" -Value 0
    Set-AppStateValue -Key "ProjectName" -Value ""
    Set-AppStateValue -Key "AllDataBlocks" -Value @()
    Set-AppStateValue -Key "FilteredDataBlocks" -Value @()
}
