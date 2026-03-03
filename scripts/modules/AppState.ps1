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
    LastExportResult   = $null           # Hashtable: SuccessCount, ErrorCount, Errors, OutputFolder

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
