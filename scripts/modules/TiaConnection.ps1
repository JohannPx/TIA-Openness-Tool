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

function Get-ScanEmptyDiagnostic {
    # Quand le scan Openness ne retourne aucune instance, determine la cause probable
    # pour guider l'utilisateur. GetProcesses() ne voit que les instances de la version
    # de DLL chargee et echoue silencieusement (liste vide) dans les autres cas :
    #   - une instance d'une AUTRE version est ouverte -> selectionner la bonne version ;
    #   - une instance de MEME version est ouverte mais inaccessible -> ecart de privileges ;
    #   - aucune instance n'est reellement ouverte -> message standard.
    # Retourne le texte localise et le type de banniere a afficher.
    $running = @(Get-RunningTiaPortalVersions)
    if ($running.Count -eq 0) {
        return @{ Text = (T "LblNoInstance"); Type = "warning" }
    }

    $loaded = (Get-AppState).SelectedVersion
    $runningVersions = @($running | ForEach-Object { $_.Version } | Sort-Object -Unique)

    if ($runningVersions -notcontains $loaded) {
        return @{ Text = ((T "LblScanVersionMismatch") -f ($runningVersions -join ", "), $loaded); Type = "warning" }
    }

    return @{ Text = ((T "LblScanRunningNotAccessible") -f $loaded); Type = "warning" }
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

function Get-OnlineProvider {
    # Recupere le service OnlineProvider d'un DeviceItem (CPU) via reflexion.
    # Comme GetService<T>(), PowerShell 5.1 ne peut pas appeler la methode generique directement.
    param([object]$DeviceItem)

    $method = $DeviceItem.GetType().GetMethod('GetService')
    if (-not $method -or -not $method.IsGenericMethod) { return $null }
    $generic = $method.MakeGenericMethod([Siemens.Engineering.Online.OnlineProvider])
    return $generic.Invoke($DeviceItem, $null)
}

function Test-DeviceItemOnline {
    # Indique si l'automate (CPU) est en ligne dans TIA Portal.
    # En cas d'indetermination (provider absent, erreur), retourne $false : on laisse alors
    # l'export tenter et retomber sur le message d'erreur par bloc (filet de securite).
    param([object]$DeviceItem)

    if (-not $DeviceItem) { return $false }
    try {
        $provider = Get-OnlineProvider -DeviceItem $DeviceItem
        if (-not $provider) { return $false }
        return ($provider.State -eq [Siemens.Engineering.Online.OnlineState]::Online)
    } catch {
        return $false
    }
}

function Get-OnlinePlcNames {
    # Retourne les noms des automates en ligne parmi ceux concernes par les blocs selectionnes.
    param([array]$SelectedBlocks)

    $infoList = (Get-AppState).PlcDeviceInfoList
    if (-not $infoList) { return @() }

    $plcIndexes = @($SelectedBlocks | ForEach-Object { $_.PlcIndex } | Sort-Object -Unique)
    $onlineNames = @()
    foreach ($idx in $plcIndexes) {
        $info = $infoList | Where-Object { $_.PlcIndex -eq $idx } | Select-Object -First 1
        if ($info -and $info.ContainsKey('DeviceItem') -and (Test-DeviceItemOnline -DeviceItem $info.DeviceItem)) {
            $onlineNames += $info.Name
        }
    }
    return @($onlineNames)
}

function Get-AllProjectDevices {
    # Collecte tous les devices du projet, y compris ceux rangés dans des groupes.
    # project.Devices ne retourne que les devices à la racine : les automates placés
    # dans un dossier (DeviceUserGroup) ou dans "Ungrouped devices" seraient ignorés.
    param([object]$Project)

    $devices = @()
    $devices += @($Project.Devices)

    # Dossier "Ungrouped devices"
    try {
        if ($Project.UngroupedDevicesGroup) {
            $devices += @($Project.UngroupedDevicesGroup.Devices)
        }
    } catch {}

    # Groupes de devices créés par l'utilisateur (récursif via les sous-groupes)
    try {
        foreach ($group in $Project.DeviceGroups) {
            $devices += @(Get-DevicesInGroup -Group $group)
        }
    } catch {}

    return $devices
}

function Get-DevicesInGroup {
    # Parcours récursif d'un groupe de devices : devices directs + sous-groupes.
    param([object]$Group)

    $devices = @()
    $devices += @($Group.Devices)
    try {
        foreach ($subGroup in $Group.Groups) {
            $devices += @(Get-DevicesInGroup -Group $subGroup)
        }
    } catch {}
    return $devices
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
            PlcIndex   = $idx
            Name       = "PLC_$idx"
            IpAddress  = ""
            Rack       = 0
            Slot       = 2
            Tsap       = "03.02"
            DeviceItem = $deviceItem   # CPU, requis pour la detection de l'etat en ligne
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

        # Enumerate PLC devices (recursive traversal of DeviceItems tree).
        # Get-AllProjectDevices inclut aussi les devices rangés dans des groupes.
        $plcResults = @()
        foreach ($device in (Get-AllProjectDevices -Project $project)) {
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
        # Build a full message by drilling into inner exceptions (Attach() wraps the real cause)
        $exMsg = $_.Exception.Message
        $inner = $_.Exception.InnerException
        while ($inner) {
            $exMsg += "`n" + $inner.Message
            $inner = $inner.InnerException
        }

        # Specific case: TIA Openness security error — the Windows user is not a member of
        # the local "Siemens TIA Openness" group. Surface an actionable, localized message
        # with a ready-to-paste command pre-filled with the actual account.
        if ($exMsg -match 'Siemens TIA Openness' -or $exMsg -match 'Security error' -or $exMsg -match 'not\s+(a\s+)?member\s+of\s+the\s+windows\s+group') {
            # The exception names the exact process owner: "Owner 'DOMAIN\User' of this process ...".
            # Extract it so the remediation command needs no manual editing; fall back to the
            # account currently running this tool if the owner can't be parsed.
            if ($exMsg -match "Owner\s+'([^']+)'") {
                $account = $Matches[1]
            } else {
                $account = "$env:USERDOMAIN\$env:USERNAME"
            }
            $command = 'net localgroup "Siemens TIA Openness" "' + $account + '" /add'
            return @{
                Success          = $false
                OpennessSecurity = $true
                Account          = $account
                Command          = $command
                Message          = (T "MsgOpennessSecurityError") -f $account, $exMsg
                Detail           = $exMsg
            }
        }

        return @{ Success = $false; Message = (T "MsgConnectError") -f $exMsg }
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
