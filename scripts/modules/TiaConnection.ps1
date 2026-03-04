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
        # Method 1: TIA Openness API ProjectPath property (most reliable)
        $projectName = ""
        try {
            $projPath = $process.ProjectPath
            if ($projPath -and $projPath.Exists) {
                $projectName = [System.IO.Path]::GetFileNameWithoutExtension($projPath.Name)
            }
        } catch {}

        # Method 2: Fallback to WMI / window title
        if ([string]::IsNullOrEmpty($projectName)) {
            $projectName = Get-ProjectNameFromProcess -ProcessId $process.Id
        }

        $version = Get-TiaVersionFromProcess -ProcessId $process.Id

        $displayText = if ([string]::IsNullOrEmpty($projectName)) {
            "TIA Portal - PID: $($process.Id)"
        } else {
            "TIA Portal - $projectName (PID: $($process.Id))"
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

function Get-SoftwareContainer {
    # Calls the generic method DeviceItem.GetService<SoftwareContainer>() via reflection
    # PowerShell 5.1 cannot call parameterless generic methods directly
    param([object]$DeviceItem)

    $method = $DeviceItem.GetType().GetMethod('GetService')
    if (-not $method -or -not $method.IsGenericMethod) { return $null }
    $generic = $method.MakeGenericMethod([Siemens.Engineering.HW.Features.SoftwareContainer])
    return $generic.Invoke($DeviceItem, $null)
}

function Find-PlcSoftwareInDevice {
    param([object]$DeviceItems)

    $result = @()
    foreach ($item in $DeviceItems) {
        try {
            $swContainer = Get-SoftwareContainer -DeviceItem $item
            if ($swContainer -and $swContainer.Software -is [Siemens.Engineering.SW.PlcSoftware]) {
                $result += @{
                    PlcSoftware = $swContainer.Software
                    DeviceItem  = $item
                }
            }
        } catch {}
        # Recurse into child DeviceItems
        if ($item.DeviceItems.Count -gt 0) {
            $result += @(Find-PlcSoftwareInDevice -DeviceItems $item.DeviceItems)
        }
    }
    return $result
}

function Build-PlcDeviceInfoList {
    param(
        [array]$PlcResults,
        [object]$Project
    )

    $plcInfoList = @()
    $idx = 1

    foreach ($plcData in $PlcResults) {
        $deviceItem = $plcData.DeviceItem
        $plcInfo = @{
            PlcIndex  = $idx
            Name      = "PLC_$idx"
            IpAddress = ""
            Rack      = 0
            Slot      = 2
            Tsap      = "03.02"
        }

        # Try to get device name (navigate up to the Device level)
        try {
            $plcInfo.Name = $deviceItem.Name
        } catch {}

        # Try to read IP address from DeviceItem or its sub-items
        try {
            # Search network interfaces in sub-items
            foreach ($subItem in $deviceItem.DeviceItems) {
                try {
                    $addresses = $subItem.Addresses
                    foreach ($addr in $addresses) {
                        try {
                            $ip = $addr.GetAttribute("IpAddress")
                            if ($ip) {
                                $plcInfo.IpAddress = $ip
                                break
                            }
                        } catch {}
                    }
                    if ($plcInfo.IpAddress) { break }
                } catch {}
            }
        } catch {}

        # Try to read Rack / Slot from attributes
        try {
            $rack = $deviceItem.GetAttribute("RackNumber")
            if ($null -ne $rack) { $plcInfo.Rack = [int]$rack }
        } catch {}
        try {
            $slot = $deviceItem.GetAttribute("SlotNumber")
            if ($null -ne $slot) { $plcInfo.Slot = [int]$slot }
        } catch {}

        # Compute TSAP from Rack/Slot: format "03.XX" where XX = rack*32 + slot in hex
        $tsapByte = ($plcInfo.Rack * 32 + $plcInfo.Slot)
        $plcInfo.Tsap = "03." + $tsapByte.ToString("X2")

        $plcInfoList += $plcInfo
        $idx++
    }

    return $plcInfoList
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

        # Enumerate PLC devices (recursive traversal of DeviceItems tree)
        $plcResults = @()
        foreach ($device in $project.Devices) {
            $plcResults += @(Find-PlcSoftwareInDevice -DeviceItems $device.DeviceItems)
        }

        if ($plcResults.Count -eq 0) {
            return @{ Success = $false; Message = T "MsgNoPlc" }
        }

        # Extract PlcSoftware objects for backward compatibility
        $plcList = @($plcResults | ForEach-Object { $_.PlcSoftware })

        # Build PLC device info (name, IP, TSAP)
        $plcDeviceInfoList = Build-PlcDeviceInfoList -PlcResults $plcResults -Project $project

        # Update state
        Set-AppStateValue -Key "TiaPortal" -Value $tiaPortal
        Set-AppStateValue -Key "CurrentProject" -Value $project
        Set-AppStateValue -Key "PlcSoftwareList" -Value $plcList
        Set-AppStateValue -Key "PlcDeviceInfoList" -Value $plcDeviceInfoList
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
    Set-AppStateValue -Key "PlcDeviceInfoList" -Value @()
    Set-AppStateValue -Key "IsConnected" -Value $false
    Set-AppStateValue -Key "ConnectedProcessId" -Value 0
    Set-AppStateValue -Key "ProjectName" -Value ""
    Set-AppStateValue -Key "AllDataBlocks" -Value @()
    Set-AppStateValue -Key "FilteredDataBlocks" -Value @()
}
