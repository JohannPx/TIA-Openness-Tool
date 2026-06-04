# AppState.ps1 - Central state management
# Single source of truth for all application data

$Script:AppState = @{
    # TIA Portal Version
    SelectedVersion    = $null           # e.g. "V19"
    InstalledVersions  = @()             # List of @{Version; MajorNumber; DllPath}
    DllPath            = $null           # Full path to loaded Siemens.Engineering.dll
    DllLoaded          = $false          # Whether DLL was loaded successfully

    # Connection
    TiaPortal          = $null           # [Siemens.Engineering.TiaPortal] object
    CurrentProject     = $null           # [Siemens.Engineering.Project] object
    PlcSoftwareList    = @()             # List of [PlcSoftware] objects
    IsConnected        = $false
    ConnectedProcessId = 0
    ProjectName        = ""

    # Instances found by scan
    TiaInstances       = @()             # List of hashtables: ProcessId, DisplayText, ProjectName, Version

    # DataBlocks
    AllDataBlocks      = @()             # Full list of DB hashtables
    FilteredDataBlocks = @()             # After filter applied
    HideInstanceDBs    = $true           # Filter toggle

    # PLC Device Info (populated during connection)
    PlcDeviceInfoList  = @()             # List of @{PlcIndex; Name; IpAddress; Rack; Slot; Tsap}

    # Export
    ExportFolder       = $null           # User-chosen or default export path
    ExportFormat       = "CSV"           # "CSV" (Siemens) or "EWON" (var_lst)
    LastExportResult   = $null           # Hashtable: SuccessCount, ErrorCount, Errors, OutputFolder

    # Ewon var_lst config
    EwonRepere         = ""              # Variable prefix (e.g. "i30")
    EwonTopic          = "A"            # Topic: A, B, or C
    EwonPageId         = 1              # Page: 1-11

    # Runtime
    IsExporting        = $false

    # Language
    Language            = "FR"
}

function Get-AppState { return $Script:AppState }

function Set-AppStateValue {
    param([string]$Key, $Value)
    $Script:AppState[$Key] = $Value
}

# Version injected at build time by CI (sed replaces @APP_VERSION@ with the resolved version).
# Stays as the literal placeholder in dev mode, which triggers the version.json / manifest.json lookup.
$Script:InjectedAppVersion = "@APP_VERSION@"

function Get-AppVersion {
    # 1. Build-time injected version (production bundle / .exe)
    if ($Script:InjectedAppVersion -and $Script:InjectedAppVersion -notmatch '^@.*@$') {
        return $Script:InjectedAppVersion
    }

    # 2. version.json maintained by the C# wrapper after install/update
    $versionFile = Join-Path $env:LOCALAPPDATA "TiaOpennessTool\version.json"
    if (Test-Path $versionFile) {
        try {
            $v = (Get-Content $versionFile -Raw | ConvertFrom-Json).version
            if ($v -and $v -ne "0.0.0") { return $v }
        } catch {}
    }

    # 3. Dev: read manifest.json (script lives in scripts/modules/)
    $candidates = @(
        (Join-Path $PSScriptRoot "..\..\manifest.json"),
        (Join-Path $PSScriptRoot "..\manifest.json"),
        (Join-Path $PSScriptRoot "manifest.json")
    )
    foreach ($p in $candidates) {
        if (Test-Path $p) {
            try {
                $v = (Get-Content $p -Raw | ConvertFrom-Json).version
                if ($v) { return $v }
            } catch {}
        }
    }

    return "dev"
}
